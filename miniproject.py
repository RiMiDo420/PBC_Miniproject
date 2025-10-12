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


credentials = ( os.environ["MICROSOFT_CLIENT_ID"],  os.environ["MICROSOFT_CLIENT_SECRET"])

TOKEN_FILE = 'o365_token_public.txt' # File to store the cached token

# Initialize the FileSystemTokenBackend to manage caching
token_backend = FileSystemTokenBackend(token_filename=TOKEN_FILE)

# Initialize the Account. Passing only the client ID implies a Public Client.
account = Account(
    credentials, # Client secret is omitted or explicitly set to None
    scopes=['basic', 'message_all'],
    token_backend=token_backend
)


if account.is_authenticated:
    print("Authenticated successfully using cached token.")

else:
    # 2. Perform the initial interactive login (Authorization Code Flow)
    print("Starting interactive authentication flow...")
    
    # This will open a browser for the user to log in and give consent.
    # The resulting URL/code is automatically processed by the library.
    if account.authenticate(redirect_uri='http://localhost:8080'):
        print("Initial interactive authentication successful. Token saved for future use.")
    else:
        print("Authentication failed.")
        exit()


# 🛑 IMPORTANT: If you change these scopes, delete the 'token.pickle' file!
# 'readonly' scope allows viewing, but not editing, events.
SCOPES = ['https://www.googleapis.com/auth/calendar'] 
CREDENTIALS_FILE = 'credentials.json' # Rename your downloaded file if necessary

# Best practice to load key from environment variable
# In your terminal:
# export GOOGLE_API_KEY='Your-API-Key'
genai.configure(api_key=os.environ["GOOGLE_API_KEY"])

calendarName = "PBC_AI"
testEvent = {
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

def get_date() ->str:
    '''
    Gets the current date and time
    '''
    date_and_time = datetime.datetime.now()
    local_now = date_and_time.astimezone()
    local_tz = local_now.tzinfo
    local_tzname = local_tz.tzname(local_now)
    print("AI got date")
    print("-"*30)

    return date_and_time.strftime("%A, %B %d, %Y at %I:%M %p") + f" The current timezone is {local_tzname}."
def get_calendar_service():
    """
    Handles the authentication flow and returns the Google Calendar API service object.
    It uses a 'token.pickle' file to cache credentials for future runs.
    """
    creds = None
    # 1. Check for existing cached token file
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # 2. If no valid token, start the OAuth flow
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            # Refresh token if expired but refresh token is available
            creds.refresh(Request())
        else:
            # Run the full flow to get new credentials
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE, SCOPES)
            
            # This opens a browser window for user login/consent
            creds = flow.run_local_server(port=0) 

        # 3. Save the new credentials (access token and refresh token)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    # 4. Build the API service object
    service = build('calendar', 'v3', credentials=creds)
    return service

def parse_output(out:str) -> list[dict[str, str]]:
    try:
        event_dict = json.loads(out)
        if(type(event_dict) == dict[str, str]):
            event_dict = [event_dict]
    
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON: {e}")
        # Handle the error, e.g., by exiting the function
        event_dict = None
    return event_dict


def get_calendar_id(name:str) ->str:
    calendar_list = service.calendarList().list().execute()


    calendars = calendar_list['items']


    for c in calendars:
        if c['summary'] == name:
            return c['id']
    
    for c in calendars:
        print(f"The calendar {c['summary']} has the id: {c['id']}")

def add_event(event_list: list[dict[str, str]], calendar_id:str) ->None:
    for event_body in event_list:
        print(f"\nTrying event {event_body['summary']}")
        try:
            event_result = service.events().insert(
            calendarId=calendar_id,
            body=event_body
            ).execute()

            print("\n✅ Event successfully created!")
            print(f"Event ID: {event_result.get('id')}")
            print(f"View on Calendar: {event_result.get('htmlLink')}")

        except Exception as e:
            print(f"\n❌ ERROR inserting event: {e}")
            print("Please check your authorization scope (needs write access) and calendar ID.")





model = genai.GenerativeModel('gemini-2.5-flash-lite', tools=[get_date], system_instruction='''
You are a summarizer. You will recieve an email and your goal is to provide the data of all events, that the email mentions in a format, that will then get passed on to google calendar.
Return the answer as a list if events in JSON. Do not include the any formatting, such as "\'\'\'json". Do not include anything other, than the JSON. Do it in one shot, do not ask follow up questions. 
If there is no date, use today's. If information is not provided, write not provided. Do not ask the user for clarification, check the year using get_date.
If no end time is specified assume the event takes an hour. If something is missing, do not include it in the JSON.
Only include a list of events, do not store evrything in a dictionary, with the key "events".
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
                                                                      ''')

chat=model.start_chat(enable_automatic_function_calling=True)


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


'''history = ""
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
    print("-"*30)'''


