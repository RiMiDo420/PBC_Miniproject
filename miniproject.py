"""
Email to Calendar Assistant 📧📅

This module implements a desktop application using Tkinter to automate the process
of extracting event information from O365 emails using a Gemini LLM and scheduling
them into a Google Calendar.

It handles:
1.  **Authentication:** Interactive OAuth2 flows for both Microsoft (O365) and Google 
    Calendar services, caching tokens for future runs.
2.  **Email Fetching:** Retrieves recent emails from a specified O365 mailbox.
3.  **LLM Integration:** Uses the Gemini 2.5 Flash Lite model with a specific 
    system instruction to parse email body text and output structured JSON event data.
4.  **Calendar Integration:** Connects to the Google Calendar API to insert the 
    parsed events into a designated calendar.
5.  **GUI:** Provides a simple, thread-safe graphical interface for selecting emails, 
    initiating the scanning process, and logging the application's activity.

Dependencies:
- google-genai
- google-api-python-client
- google-auth-oauthlib
- O365
- tkinter (built-in)

Environment Variables Required:
- MICROSOFT_CLIENT_ID
- MICROSOFT_CLIENT_SECRET
- TARGET_EMAIL (for O365 mailbox)
- GOOGLE_API_KEY
"""

import os
import os.path
import datetime
import pickle
import json
import textwrap
import tkinter as tk
from tkinter import ttk, scrolledtext, simpledialog
from threading import Thread
import platform
from dotenv import load_dotenv
from zoneinfo import ZoneInfo

import google.generativeai as genai
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from O365 import Account
from O365.utils import FileSystemTokenBackend

load_dotenv()


# Sets the name of the caledndar, to which the events will be pushed
# This calendar must be a calendar you have already started in your google account
# Otherwise it will list your calendars in the console, but will not work

event_schema ={
  "type": "array",
  "description": "A list of extracted calendar events with full datetime, timezone, location, and description.",
  "items": {
    "type": "object",
    "description": "A single calendar event.",
    "properties": {
      "summary": {
        "type": "string",
        "description": "A brief, descriptive title for the event."
      },
      "start": {
        "type": "object",
        "description": "The starting datetime of the event.",
        "properties": {
          "dateTime": {
            "type": "string",
            "format": "date-time",
            "description": "The date and time in ISO 8601 format (e.g., YYYY-MM-DDTHH:MM:SS)."
          },
          "timeZone": {
            "type": "string",
            "description": "The IANA time zone identifier (e.g., 'America/Los_Angeles')."
          }
        },
        "required": ["dateTime", "timeZone"]
      },
      "end": {
        "type": "object",
        "description": "The ending datetime of the event.",
        "properties": {
          "dateTime": {
            "type": "string",
            "format": "date-time",
            "description": "The date and time in ISO 8601 format (e.g., YYYY-MM-DDTHH:MM:SS)."
          },
          "timeZone": {
            "type": "string",
            "description": "The IANA time zone identifier (e.g., 'America/Los_Angeles')."
          }
        },
        "required": ["dateTime", "timeZone"]
      },
      "location": {
        "type": "string",
        "description": "The physical or virtual location of the event (e.g., address, conference room, or video link)."
      },
      "description": {
        "type": "string",
        "description": "A detailed body or notes for the event, including context or agenda."
      }
    },
    "required": ["summary", "start", "end"]
  }
}

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
        sys.exit()


# 🛑 IMPORTANT: If you change these scopes, delete the 'token.pickle' file!
SCOPES = ["https://www.googleapis.com/auth/calendar"]
CREDENTIALS_FILE = "credentials.json"

genai.configure(api_key=os.environ["GOOGLE_API_KEY"])




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
    s = build("calendar", "v3", credentials=creds)
    return s


def parse_output(out: str) -> list[dict[str, str]]:
    """
    Converts the json returned by the AI to a list of dictionaries representing the events, 
    for the google calendar API.
    """
    try:
        event_dict = json.loads(out)
        if isinstance(event_dict, dict):
            event_dict = [event_dict]

    except json.JSONDecodeError as e:
        print(f"Error decoding JSON: {e}")
        # Handle the error, e.g., by exiting the function
        event_dict = []
    return event_dict


service = get_calendar_service()

def list_calendars(calendars) ->None:
    print("You have these calendars:")
    for c in calendars:
        print(f"{c['summary']}")

def get_user_calendars():
    calendar_list = service.calendarList().list().execute()

    return calendar_list["items"]



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
If there are multiple events, return the required items in a list.
Keep the description field short. A max of a few sentences.
We are currently in the +01:00 timezone, so make all the events in this timezone too
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
                                                                      """
)


def extract_events_with_llm(app_instance, prompt:str) -> list[dict[str, str]]:
    '''
    Extracts events from email text, using an LLM
    '''
    chat = model.start_chat(enable_automatic_function_calling=True)

    config = genai.types.GenerationConfig(
    # Use the parameter name expected by the config object
    response_mime_type="application/json", 
    response_schema=event_schema,
)   

    result = chat.send_message(prompt, generation_config=config)
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


class CalendarSelectWindow(simpledialog.Dialog):
    """
    A simple dialog window for the user to select one of their calendars.
    It calls the CalendarManager to get the list of calendars.
    """
    def __init__(self, parent, current_calendar_id):
        self.current_calendar_id = current_calendar_id
        self.selected_calendar_id = current_calendar_id
        super().__init__(parent, title="Select Destination Calendar")

    def body(self, master):
        """Creates the widgets for the dialog body."""
        ttk.Label(master, text="Select the calendar for new events:").pack(pady=5)
        
        # Fetch the calendars
        try:
            self.calendars = get_user_calendars()
        except Exception as e:
            ttk.Label(master, text=f"Error fetching calendars: {e}", foreground="red").pack(pady=5)
            self.calendars = []
            return None

        # Prepare the list for the Combobox
        calendar_names = [c['summary'] for c in self.calendars]
        self.calendar_ids = [c['id'] for c in self.calendars]

        # Find the index of the currently selected calendar
        try:
            current_index = self.calendar_ids.index(self.current_calendar_id)
        except ValueError:
            current_index = 0 # Default to the first one if not found

        self.calendar_var = tk.StringVar(master)
        
        self.calendar_combobox = ttk.Combobox(
            master,
            textvariable=self.calendar_var,
            values=calendar_names,
            state="readonly"
        )
        self.calendar_combobox.current(current_index)
        self.calendar_combobox.pack(padx=10, pady=10, fill='x')

        return self.calendar_combobox # Initial focus

    def apply(self):
        """Called when the user clicks 'OK'."""
        selected_name = self.calendar_var.get()
        # Find the corresponding ID for the selected name
        try:
            index = [c['summary'] for c in self.calendars].index(selected_name)
            self.selected_calendar_id = self.calendar_ids[index]
            self.result = self.selected_calendar_id # Set the result
        except ValueError:
            # Should not happen with a 'readonly' combobox
            self.result = self.current_calendar_id



class EmailCalendarApp(tk.Tk):
    '''
    The thing, that handles the simple UI, idk I didn't write this, gemini did...
    '''
    def __init__(self):
        super().__init__()
        self.title("Email to Calendar Assistant")
        self.geometry("800x600")

        self.emails = []
        self.check_vars = []

        self.email_limit = 10  # Default number of emails to fetch and display
        self.calendar_id = "Select calendar..." # Current target calendar ID
        self.os_name = platform.system()


        self.create_widgets()
        self.load_emails()

        self.protocol("WM_DELETE_WINDOW", self.quit_app)

    def _on_mousewheel(self, event):
        """Scrolls the canvas based on mouse wheel/trackpad movement."""
        
        # Determine scroll direction and multiplier based on OS
        if self.os_name == "Windows":
            # Windows trackpads/wheels give a delta of 120 or -120
            scroll_amount = int(-1 * (event.delta/120)) 
        elif self.os_name == "Darwin": # macOS
            # macOS trackpads often report much smaller, smooth delta values
            # The <Control-MouseWheel> binding often works here for trackpads
            if event.delta: # Check if delta is available (Windows/macOS)
                 scroll_amount = int(-1 * (event.delta/30)) # Adjust sensitivity
            else: # Fallback for buttons (Button-4/Button-5 on older systems)
                 scroll_amount = -1 if event.num == 5 else 1 
        else: # Linux/Other (Uses Button-4/Button-5)
            scroll_amount = -1 if event.num == 5 else 1
            
        self.canvas.yview_scroll(scroll_amount, "units")


    def create_widgets(self):
        '''
        ??? draws the UI I assume...
        '''
        # Configure layout (two main frames: Emails/Controls and Log)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Main Frame for Email Display and Controls
        main_frame = ttk.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        # Control Panel Frame (for new widgets)
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=0, column=0, sticky="new", pady=(0, 5), padx=5)
        # Allow the email frame to take the rest of the space (row 1)
        main_frame.grid_rowconfigure(1, weight=1) 
        
        # ----------------------------------------------------
        # 1. Email Limit Selector (Spinbox)
        # ----------------------------------------------------
        ttk.Label(control_frame, text="Emails to Show:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.email_limit_var = tk.IntVar(value=self.email_limit)
        self.email_limit_spinbox = ttk.Spinbox(
            control_frame,
            from_=5,
            to=50,
            increment=5,
            width=5,
            textvariable=self.email_limit_var,
            command=self.update_email_limit_and_reload # Command to call on value change
        )
        self.email_limit_spinbox.pack(side=tk.LEFT, padx=(0, 15))
        
        # ----------------------------------------------------
        # 2. Calendar Select Button and Display
        # ----------------------------------------------------
        self.calendar_display_label = ttk.Label(
            control_frame, 
            text=f"Target Calendar: {self.calendar_id}",
            anchor='w'
        )
        self.calendar_display_label.pack(side=tk.LEFT, padx=(5, 5))

        self.calendar_select_button = ttk.Button(
            control_frame,
            text="Change Calendar",
            command=self.open_calendar_select_window
        )
        self.calendar_select_button.pack(side=tk.LEFT, padx=(0, 5))


        # Email Display Frame (Scrollable area for checkboxes)
        email_frame = ttk.LabelFrame(
            main_frame, text="Recent Emails (Select for Scanning)", padding="10"
        )
        # Change grid row to 1 to be below the control_frame
        email_frame.grid(row=1, column=0, sticky="nsew", pady=10) 
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

                # 1. Bind to the Canvas (for when the mouse is over the blank scroll area)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Button-4>", self._on_mousewheel)
        self.canvas.bind("<Button-5>", self._on_mousewheel)
        
        # 2. Bind the inner frame. This is crucial for trackpad scrolling to work
        # when the cursor is over the list of Checkbuttons.
        self.scrollable_frame.bind("<MouseWheel>", self._on_mousewheel)
        self.scrollable_frame.bind("<Button-4>", self._on_mousewheel)
        self.scrollable_frame.bind("<Button-5>", self._on_mousewheel)

        # 3. Add the macOS/Trackpad-specific binding as an extra check
        # This often registers the "natural" two-finger scroll on macOS.
        if self.os_name == "Darwin":
             self.canvas.bind("<Control-MouseWheel>", self._on_mousewheel)
             self.scrollable_frame.bind("<Control-MouseWheel>", self._on_mousewheel)

        # Control Button
        self.scan_button = ttk.Button(
            main_frame,
            text="Scan Selected Emails to Calendar",
            command=self.start_scan_thread,
        )
        self.scan_button.grid(row=2, column=0, pady=10) # Change row to 2

        # Log Frame
        log_frame = ttk.LabelFrame(self, text="Application Log", padding="10")
        log_frame.grid(row=1, column=0, sticky="nsew")
        self.grid_rowconfigure(1, weight=0)  # Log frame should not stretch vertically

        self.log_text = scrolledtext.ScrolledText(
            log_frame, wrap=tk.WORD, width=90, height=10, state="disabled"
        )
        self.log_text.pack(expand=True, fill="both")

    def update_email_limit_and_reload(self):
        """Updates the email limit and reloads the email list."""
        try:
            new_limit = self.email_limit_var.get()
            if new_limit != self.email_limit:
                self.email_limit = new_limit
                self.log_message(f"Email limit set to {self.email_limit}. Reloading emails...")
                self.load_emails()
        except tk.TclError:
            # Handle non-integer input in the spinbox if state wasn't 'readonly'
            pass

    def open_calendar_select_window(self):
        """Opens the dialog for calendar selection."""
        # CalendarSelectWindow is a modal dialog, execution blocks until it closes
        dialog = CalendarSelectWindow(self, self.calendar_id)
        
        if dialog.result is not None:
            new_calendar_id = dialog.result
            if new_calendar_id != self.calendar_id:
                self.calendar_id = new_calendar_id
                # Attempt to get the summary for display, fallback to ID
                selected_calendar = next((c for c in get_user_calendars() if c['id'] == self.calendar_id), None)
                display_name = selected_calendar['summary'] if selected_calendar else self.calendar_id

                self.calendar_display_label.config(text=f"Target Calendar: {display_name}")
                self.log_message(f"Target calendar changed to: {display_name} ({self.calendar_id})")

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
        fetched_emails = fetch_recent_emails_o365(self, count=self.email_limit)
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

            cb.bind("<MouseWheel>", self._on_mousewheel)
            cb.bind("<Button-4>", self._on_mousewheel)
            cb.bind("<Button-5>", self._on_mousewheel)
            
            # Add the macOS/Trackpad-specific binding to the child widgets as well
            if self.os_name == "Darwin":
                 cb.bind("<Control-MouseWheel>", self._on_mousewheel)

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
            extracted_events = extract_events_with_llm(self, email.get_body_text()+f"\nThe current date is {email.received.astimezone(ZoneInfo("Europe/London"))}. The timezone is {email.received.astimezone()}")

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
