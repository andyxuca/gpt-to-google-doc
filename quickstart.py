import os.path
import json
import markdown
from bs4 import BeautifulSoup

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Scopes for access
SCOPES = ["https://www.googleapis.com/auth/documents.readonly", "https://www.googleapis.com/auth/drive.file"]

# Folder containing documents
FOLDER_ID = "1hFB8bLPIRGk-TFrdfAUCAC0DRnqPAr5Y"

# The ID of a sample document.
document_id = None

# credentials
creds = None

def build_creds():
  global creds

  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
        "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

# Creates a new google doc titled "ABC One-Pager" in the specified folder
def create_google_doc():
  global document_id

  try:
    # Build the Google Docs and Drive services
    docs_service = build("docs", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)

    # Create a new Google Doc
    document = docs_service.documents().create(body={"title": "ABC One-Pager"}).execute()
    document_id = document.get("documentId")

    print(f"Created document with ID: {document_id}")

    # Move the document to the specified folder
    file = drive_service.files().update(
      fileId=document_id,
      addParents=FOLDER_ID,
      fields="id, parents"
    ).execute()

    print(f"Document moved to folder ID: {FOLDER_ID}")

  except HttpError as err:
    print(err)

# Adds content from json file to google doc
def add_content_to_doc(json_file):
  global document_id

  try:
    # Build the Google Docs service
    docs_service = build("docs", "v1", credentials=creds)

    # Load the JSON file
    with open(json_file, 'r') as file:
      data = json.load(file)

    # Extract the content from the "gptOutput" key
    gpt_output = data.get("gptOutput", "")

    if not gpt_output:
      print("No 'gptOutput' found in the JSON file.")
      return

    header_request = add_header_to_doc()

    # Convert Markdown to Google Docs requests
    requests = header_request + markdown_to_docs_requests(gpt_output)

    # Execute the request
    result = docs_service.documents().batchUpdate(
      documentId=document_id,
      body={'requests': requests}
    ).execute()

    print(f"Added content to document ID: {document_id}")

  except HttpError as err:
    print(err)

# Converts given markdown into google docs requests
def markdown_to_docs_requests(markdown_text):
  # Convert Markdown to HTML
  html = markdown.markdown(markdown_text)

  # Parse the HTML
  soup = BeautifulSoup(html, 'html.parser')

  # Initialize Google Doc requests list
  requests = []
  index = 1  # Start index (1 because 0 is the beginning of the document)

  for element in soup.descendants:
    # Convert Markdown to HTML
    html = markdown.markdown(markdown_text)

    # Parse the HTML
    soup = BeautifulSoup(html, 'html.parser')

    # Remove first line that doesn't include necessary information
    first_element = soup.find()
    if first_element:
        first_element.extract()

    # Initialize Google Doc requests list
    requests = []
    index = 1  # Start index (1 because 0 is the beginning of the document)

    for element in soup.descendants:
      if element.name == 'h1':
        text = element.get_text() + '\n\n'
        requests.append({
          'insertText': {
            'location': {
              'index': index,
            },
            'text': text,
          }
        })
        requests.append({
          'updateParagraphStyle': {
            'range': {
              'startIndex': index,
              'endIndex': index + len(text) - 1
            },
            'paragraphStyle': {
              'namedStyleType': 'TITLE'
            },
            'fields': 'namedStyleType'
          }
        })
        requests.append({
          'updateTextStyle': {
            'range': {
              'startIndex': index,
              'endIndex': index + len(text) - 2  # Exclude the newline characters
            },
            'textStyle': {
              'bold': True,
              'fontSize': {
                'magnitude': 20,
                'unit': 'PT'
              }
            },
            'fields': 'bold,fontSize'
          }
        })
        index += len(text)
      elif element.name == 'h2':
        text = element.get_text() + '\n\n'
        requests.append({
          'insertText': {
            'location': {
              'index': index,
            },
            'text': text,
          }
        })
        requests.append({
          'updateParagraphStyle': {
            'range': {
              'startIndex': index,
              'endIndex': index + len(text) - 1
            },
            'paragraphStyle': {
              'namedStyleType': 'SUBTITLE'
            },
            'fields': 'namedStyleType'
          }
        })
        requests.append({
          'updateTextStyle': {
            'range': {
              'startIndex': index,
              'endIndex': index + len(text) - 2  # Exclude the newline characters
            },
            'textStyle': {
              'bold': True,
              'fontSize': {
                'magnitude': 18,
                'unit': 'PT'
              }
            },
            'fields': 'bold,fontSize'
          }
        })
        index += len(text)
      elif element.name == 'h3':
        text = element.get_text() + '\n\n'
        requests.append({
          'insertText': {
            'location': {
              'index': index,
            },
            'text': text,
          }
        })
        requests.append({
          'updateParagraphStyle': {
            'range': {
              'startIndex': index,
              'endIndex': index + len(text) - 1
            },
            'paragraphStyle': {
              'namedStyleType': 'HEADING_3'
            },
            'fields': 'namedStyleType'
          }
        })
        requests.append({
          'updateTextStyle': {
            'range': {
              'startIndex': index,
              'endIndex': index + len(text) - 2  # Exclude the newline characters
            },
            'textStyle': {
              'bold': True,
              'fontSize': {
                'magnitude': 15,
                'unit': 'PT'
              },
              'foregroundColor': {
                'color': {
                  'rgbColor': {
                    'red': 0,
                    'green': 0,
                    'blue': 0
                  }
                }
              }
            },
            'fields': 'bold,fontSize,foregroundColor'
          }
        })
        index += len(text)
      elif element.name == 'p':
        # Handle paragraph with possible strong text
        for sub_element in element.children:
          if sub_element.name == 'strong':
            text = sub_element.get_text()
            requests.append({
              'insertText': {
                'location': {
                  'index': index,
                },
                'text': text,
              }
            })
            requests.append({
              'updateTextStyle': {
                'range': {
                  'startIndex': index,
                  'endIndex': index + len(text)
                },
                'textStyle': {
                  'bold': True
                },
                'fields': 'bold'
              }
            })
            index += len(text)
          else:
            text = sub_element.get_text()
            requests.append({
              'insertText': {
                'location': {
                  'index': index,
                },
                'text': text,
              }
            })
            requests.append({
              'updateTextStyle': {
                'range': {
                  'startIndex': index,
                  'endIndex': index + len(text)
                },
                'textStyle': {
                  'bold': False
                },
                'fields': 'bold'
              }
            })
            index += len(text)
        # Add a new line after the paragraph
        requests.append({
          'insertText': {
            'location': {
              'index': index,
            },
            'text': '\n\n',
          }
        })
        index += 2
      elif element.name == 'em':
        text = element.get_text()
        requests.append({
          'insertText': {
            'location': {
              'index': index,
            },
            'text': text,
          }
        })
        requests.append({
          'updateTextStyle': {
            'range': {
              'startIndex': index,
              'endIndex': index + len(text)
            },
            'textStyle': {
              'italic': True
            },
            'fields': 'italic'
          }
        })
        index += len(text)

    return requests

# Creates google doc request for a header with tuck advisor logo
def add_header_to_doc():
  """Adds a header to the document with an image above the text and returns the requests."""
  docs_service = build("docs", "v1", credentials=creds)

  # Create a header
  create_header_response = docs_service.documents().batchUpdate(
    documentId=document_id,
    body={
      "requests": [
        {
          "createHeader": {
            "type": "DEFAULT"
          }
        }
      ]
    }
  ).execute()

  # Get the header ID
  header_id = create_header_response['replies'][0]['createHeader']['headerId']

  # Define the header text
  header_text = "\nABC Acquisition Analysis\n"

  return [
    # Insert the image into the header
    {
      'insertInlineImage': {
        'location': {
          'segmentId': header_id,
          'index': 0,
        },
        'uri': 'https://drive.google.com/uc?id=1TgCS1vQE4_4J07cMbuk24TAWYrj8zVto',
        'objectSize': {
          'height': {
            'magnitude': 100,
            'unit': 'PT'
          },
          'width': {
            'magnitude': 100,
            'unit': 'PT'
          }
        }
      }
    },
    # Center the image in the header
    {
      'updateParagraphStyle': {
        'range': {
          'segmentId': header_id,
          'startIndex': 0,
          'endIndex': 1,  # Just after the image
        },
        'paragraphStyle': {
          'alignment': 'CENTER'
        },
        'fields': 'alignment'
      }
    },
    # Insert the header text below the image
    {
      'insertText': {
        'location': {
          'segmentId': header_id,
          'index': 1,  # After the image
        },
        'text': header_text,
      }
    },
    # Style the header text
    {
      'updateTextStyle': {
        'range': {
          'segmentId': header_id,
          'startIndex': 1,
          'endIndex': 1 + len(header_text),
        },
        'textStyle': {
          'bold': True,
          'fontSize': {
            'magnitude': 16,
            'unit': 'PT'
          }
        },
        'fields': 'bold,fontSize'
      }
    },
    # Center the header text
    {
      'updateParagraphStyle': {
        'range': {
          'segmentId': header_id,
          'startIndex': 1,
          'endIndex': 1 + len(header_text),
        },
        'paragraphStyle': {
          'alignment': 'CENTER'
        },
        'fields': 'alignment'
      }
    }
  ]

def main():
  build_creds()
  create_google_doc()
  if document_id:
    add_content_to_doc("abc.json")

if __name__ == "__main__":
  main()