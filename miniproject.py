import google.generativeai as genai
import os
import datetime
import os.path
import pickle
import json
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from O365 import Account, Connection
from O365.utils import FileSystemTokenBackend
import textwrap
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from threading import Thread


credentials = (os.environ["MICROSOFT_CLIENT_ID"], os.environ["MICROSOFT_CLIENT_SECRET"])

TOKEN_FILE = "o365_token_public.txt"  # File to store the cached token

# Initialize the FileSystemTokenBackend to manage caching
token_backend = FileSystemTokenBackend(token_filename=TOKEN_FILE)

# Initialize the Account. Passing only the client ID implies a Public Client.
account = Account(
    credentials,  # Client secret is omitted or explicitly set to None
    scopes=["basic", "message_all"],
    token_backend=token_backend,
)


if account.is_authenticated:
    print("Authenticated successfully using cached token.")

else:
    # 2. Perform the initial interactive login (Authorization Code Flow)
    print("Starting interactive authentication flow...")

    # This will open a browser for the user to log in and give consent.
    # The resulting URL/code is automatically processed by the library.
    if account.authenticate(redirect_uri="http://localhost:8080"):
        print(
            "Initial interactive authentication successful. Token saved for future use."
        )
    else:
        print("Authentication failed.")
        exit()


# 🛑 IMPORTANT: If you change these scopes, delete the 'token.pickle' file!
# 'readonly' scope allows viewing, but not editing, events.
SCOPES = ["https://www.googleapis.com/auth/calendar"]
CREDENTIALS_FILE = "credentials.json"  # Rename your downloaded file if necessary

# Best practice to load key from environment variable
# In your terminal:
# export GOOGLE_API_KEY='Your-API-Key'
genai.configure(api_key=os.environ["GOOGLE_API_KEY"])

calendarName = "PBC_AI"
testEvent = {
    "summary": "Test event",
    "location": "Trinity college, Cambridge",
    "description": "test",
    "start": {"dateTime": "2025-10-11T15:00:00+01:00", "timeZone": "Europe/London"},
    "end": {"dateTime": "2025-10-11T16:00:00+01:00", "timeZone": "Europe/London"},
}


def get_date() -> str:
    """
    Gets the current date and time
    """
    date_and_time = datetime.datetime.now()
    local_now = date_and_time.astimezone()
    local_tz = local_now.tzinfo
    local_tzname = local_tz.tzname(local_now)
    print("AI got date")
    print("-" * 30)

    return (
        date_and_time.strftime("%A, %B %d, %Y at %I:%M %p")
        + f" The current timezone is {local_tzname}."
    )


def get_calendar_service():
    """
    Handles the authentication flow and returns the Google Calendar API service object.
    It uses a 'token.pickle' file to cache credentials for future runs.
    """
    creds = None
    # 1. Check for existing cached token file
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    # 2. If no valid token, start the OAuth flow
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            # Refresh token if expired but refresh token is available
            creds.refresh(Request())
        else:
            # Run the full flow to get new credentials
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)

            # This opens a browser window for user login/consent
            creds = flow.run_local_server(port=0)

        # 3. Save the new credentials (access token and refresh token)
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)

    # 4. Build the API service object
    service = build("calendar", "v3", credentials=creds)
    return service


def parse_output(out: str) -> list[dict[str, str]]:
    """
    Converts the json returned by the AI to a list of dictionaries representing the events, for the google calendar API.
    """
    try:
        event_dict = json.loads(out)
        if type(event_dict) == dict:
            event_dict = [event_dict]

    except json.JSONDecodeError as e:
        print(f"Error decoding JSON: {e}")
        # Handle the error, e.g., by exiting the function
        event_dict = None
    return event_dict


service = get_calendar_service()


def get_calendar_id(name: str) -> str:
    """
    Gets the google_id of the callendar called 'name'
    """
    calendar_list = service.calendarList().list().execute()

    calendars = calendar_list["items"]

    for c in calendars:
        if c["summary"] == name:
            return c["id"]

    for c in calendars:
        print(f"The calendar {c['summary']} has the id: {c['id']}")


def add_event(app_instance, event_body: dict[str, str], calendar_id: str) -> None:
    """
    Adds all events in event_list to the google calendar represented by calendar_id
    """
    app_instance.log_message(f"\nTrying event {event_body['summary']}")
    try:
        event_result = (
            service.events().insert(calendarId=calendar_id, body=event_body).execute()
        )

        app_instance.log_message("\n✅ Event successfully created!")
        app_instance.log_message(f"Event ID: {event_result.get('id')}")
        app_instance.log_message(f"View on Calendar: {event_result.get('htmlLink')}")

    except Exception as e:
        app_instance.log_message(f"\n❌ ERROR inserting event: {e}")
        app_instance.log_message(
            "Please check your authorization scope (needs write access) and calendar ID."
        )


def fetch_recent_emails_o365(app_instance, count=10):
    """
    Simulates fetching recent emails. Logs status to the GUI.
    """
    app_instance.log_message("--- (Simulating O365 Fetch) ---")

    mailbox = account.mailbox(resource=os.environ["TARGET_EMAIL"])

    # Get the Inbox folder
    inbox = mailbox.inbox_folder()

    # Retrieve messages from the Inbox.
    # get_messages() returns an iterator of Message objects.
    # You can use 'limit' to restrict the number of messages fetched.
    print("Fetching messages from Inbox...")
    return inbox.get_messages(limit=count)


model = genai.GenerativeModel(
    "gemini-2.5-flash-lite",
    system_instruction="""
You are a summarizer. You will recieve an email and your goal is to provide the data of all events, that the email mentions in a format, that will then get passed on to google calendar.
Return the answer as a list if events in JSON. Do not include the any formatting, such as "\'\'\'json". Do not include anything other, than the JSON. Do it in one shot, do not ask follow up questions. 
If there is no date, use today's. If information is not provided, write not provided. Do not ask the user for clarification, the date and year are provided in the email.
If no end time is specified assume the event takes an hour. If something is missing, do not include it in the JSON.
Only include a list of events, do not store evrything in a dictionary, with the key "events".
Keep the description field short. A max of a few sentences.
Here is an example output (keep the field names the same as in the example):
{
    "summary" : "Test event",
    "location": "Trinity college, Cambridge",
    "description": "test",
    "start": {
        "dateTime": "2025-10-11T15:00:00+01:00",
        "timeZone": "Europe/London"
    },
    "end": {
        "dateTime": "2025-10-11T16:00:00+01:00",
        "timeZone": "Europe/London"
    },

}
                                                                      """,
)


def extract_events_with_llm(app_instance, prompt):
    chat = model.start_chat(enable_automatic_function_calling=True)

    result = chat.send_message(prompt)
    text = result.text
    if text[0] == "`":
        text = text[7:-3]
    text = text.strip()
    if text.startswith('{"events":'):
        text = text.removeprefix('{"events":').removesuffix("}").strip()
    app_instance.log_message(text)
    app_instance.log_message("-" * 30)
    event_list = parse_output(text)
    return event_list


"""chat=model.start_chat(enable_automatic_function_calling=True)


mailbox = account.mailbox(resource=os.environ['TARGET_EMAIL'])    

# Get the Inbox folder
inbox = mailbox.inbox_folder()

# Retrieve messages from the Inbox. 
# get_messages() returns an iterator of Message objects.
# You can use 'limit' to restrict the number of messages fetched.
print("Fetching messages from Inbox...")
messages = inbox.get_messages(limit=10) # Get up to the 10 newest messages

# Iterate through the messages and print their details
service = get_calendar_service()
calId = get_calendar_id(calendarName)
for message in messages:
    try:
        chat = model.start_chat(enable_automatic_function_calling=True)
        prompt = message.get_body_text() + f"\nThe current date and time is {message.received}"
        print("-"*30)
        print(prompt)
        print("-"*30)
        result = chat.send_message(prompt)
        text = result.text
        if text[0] == '`':
            text = text[7:-3]
        text =text.strip()
        if text.startswith('{"events":'):
            text = text.removeprefix('{"events":').removesuffix('}').strip()
        print(text)
        print("-"*30)
        event_list = parse_output(text)
        for event in event_list:
            print(event)
            print('\n')
        print("-"*30)
        add_event(event_list, calId)
    except FileNotFoundError:
        print(f"\n[ERROR] The file '{CREDENTIALS_FILE}' was not found.")
        print("Please ensure your downloaded OAuth JSON file is in the same directory and named correctly.")
    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")
"""

"""history = ""
while(True):


    if prompt == "exit":
        break
    result = model.generate_content(history + prompt)

    print("-"*30)
    print(result.text)
    history += summarizer.generate_content("User: " + prompt).text+'\n'
    history += summarizer.generate_content("AI: " + result.text).text+'\n'
    print("-"*30)
    print(history)
    print("-"*30)"""


class EmailCalendarApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Email to Calendar Assistant")
        self.geometry("800x600")

        self.emails = []
        self.check_vars = []

        self.calendar_id = get_calendar_id(calendarName)

        self.create_widgets()
        self.load_emails()

        self.protocol("WM_DELETE_WINDOW", self.quit_app)

    def create_widgets(self):
        # Configure layout (two main frames: Emails/Controls and Log)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Main Frame for Email Display and Controls
        main_frame = ttk.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        # Email Display Frame (Scrollable area for checkboxes)
        email_frame = ttk.LabelFrame(
            main_frame, text="Recent Emails (Select for Scanning)", padding="10"
        )
        email_frame.grid(row=0, column=0, sticky="nsew", pady=10)
        email_frame.grid_rowconfigure(0, weight=1)
        email_frame.grid_columnconfigure(0, weight=1)

        # Canvas and Scrollbar for email list
        self.canvas = tk.Canvas(email_frame, borderwidth=0)
        self.v_scrollbar = ttk.Scrollbar(
            email_frame, orient="vertical", command=self.canvas.yview
        )
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.v_scrollbar.grid(row=0, column=1, sticky="ns")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )

        # Control Button
        self.scan_button = ttk.Button(
            main_frame,
            text="Scan Selected Emails to Calendar",
            command=self.start_scan_thread,
        )
        self.scan_button.grid(row=1, column=0, pady=10)

        # Log Frame
        log_frame = ttk.LabelFrame(self, text="Application Log", padding="10")
        log_frame.grid(row=1, column=0, sticky="nsew")
        self.grid_rowconfigure(1, weight=0)  # Log frame should not stretch vertically

        self.log_text = scrolledtext.ScrolledText(
            log_frame, wrap=tk.WORD, width=90, height=10, state="disabled"
        )
        self.log_text.pack(expand=True, fill="both")

    def log_message(self, message):
        """Adds a message to the log area safely from any thread."""
        # Use self.after() to schedule the GUI update on the main thread
        self.after(0, self._safe_log, message)

    def _safe_log(self, message):
        """Internal method executed safely on the main thread to update the GUI."""
        # All GUI modifications must happen here
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)  # Scroll to bottom
        self.log_text.config(state="disabled")
        self.update()  # Force up

    def load_emails(self):
        """Fetches emails and populates the scrollable frame with checkboxes."""
        fetched_emails = fetch_recent_emails_o365(self)
        self.emails = list(fetched_emails)
        self.check_vars = []

        # Clear previous widgets
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        if not self.emails:
            ttk.Label(
                self.scrollable_frame, text="No emails found or failed to fetch."
            ).pack(padx=5, pady=5, anchor="w")
            return

        for i, email in enumerate(self.emails):
            var = tk.BooleanVar(value=False)
            self.check_vars.append(var)

            # Create a label with the subject and a snippet
            subject_text = f"[{i + 1}] {email.subject[4::]}"
            snippet = textwrap.shorten(
                email.get_body_text().split("\n")[0], width=70, placeholder="..."
            )

            # Use a Checkbutton to combine the selection and display
            cb = ttk.Checkbutton(
                self.scrollable_frame,
                text=f"{subject_text} - {snippet}",
                variable=var,
                command=self.update_selection_count,
            )
            cb.pack(padx=5, pady=2, anchor="w")

        self.update_selection_count()

    def quit_app(self):
        """Cleanly closes the application window and stops the main loop."""
        self.log_message("Application closing...")
        self.destroy()

    def update_selection_count(self):
        """Updates the scan button text based on selection count."""
        count = sum(var.get() for var in self.check_vars)
        if count == 0:
            self.scan_button.config(
                text="Scan Selected Emails to Calendar", state="disabled"
            )
        else:
            self.scan_button.config(
                text=f"Scan {count} Email(s) to Calendar", state="normal"
            )

    def start_scan_thread(self):
        """Starts the scanning process in a separate thread."""
        self.scan_button.config(state="disabled", text="Processing...")
        Thread(target=self.process_selected_emails).start()

    def process_selected_emails(self):
        """Handles the main logic for scanning, LLM extraction, and calendar insertion."""

        self.log_message("\n" + "=" * 50)
        self.log_message("SCANNING STARTED")
        self.log_message("=" * 50)

        selected_emails = [
            self.emails[i] for i, var in enumerate(self.check_vars) if var.get()
        ]

        if not selected_emails:
            self.log_message("No emails were selected for scanning.")
            self.scan_button.config(
                state="normal", text="Scan Selected Emails to Calendar"
            )
            return

        total_events_created = 0

        for email in selected_emails:
            self.log_message(f"\n--- Scanning Email: {email.subject} ---")

            # 1. LLM Event Extraction
            extracted_events = extract_events_with_llm(self, email.get_body_text())

            if extracted_events:
                self.log_message(
                    f"✅ Found {len(extracted_events)} potential event(s)."
                )

                for event in extracted_events:
                    # 2. Google Calendar Event Creation
                    add_event(self, event, self.calendar_id)
                    total_events_created += 1
            else:
                self.log_message("❌ No calendar events were extracted by the LLM.")

        self.log_message("\n" + "=" * 50)
        self.log_message(
            f"PROCESS COMPLETE. Total events scheduled: {total_events_created}"
        )
        self.log_message("=" * 50)

        # Reset GUI state
        self.scan_button.config(state="normal", text="Scan Selected Emails to Calendar")
        self.update_selection_count()


if __name__ == "__main__":
    app = EmailCalendarApp()
    app.mainloop()
