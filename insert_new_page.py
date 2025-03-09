import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError



def extract_text_from_element(element):
    """Extracts text from a single page element.

    Args:
        element: A dictionary representing a page element.

    Returns:
        A string containing the extracted text, or an empty string if no text is found.
    """
    text = ""
    if "shape" in element:
        shape = element["shape"]
        if "text" in shape:
            text_content = shape["text"]
            if "textElements" in text_content:
                for text_element in text_content["textElements"]:
                    if "textRun" in text_element:
                        text += text_element["textRun"]["content"]
    elif "table" in element:
        table = element["table"]
        if "tableRows" in table:
            for row in table["tableRows"]:
                if "tableCells" in row:
                    for cell in row["tableCells"]:
                        if 'text' in cell:
                            text_content = cell["text"]
                            if "textElements" in text_content:
                                for text_element in text_content["textElements"]:
                                    if "textRun" in text_element:
                                        text += text_element["textRun"]["content"]
    return text


def create_new_slide(service, presentation_id):
    """Creates a new slide and returns its objectId.

    Args:
        service: The Google Slides API service object.
        presentation_id: The ID of the presentation.

    Returns:
        The objectId of the newly created slide.
    """
    requests = [
        {"createSlide": {"insertionIndex": 0}} # 1ページ目に追加する
    ]

    response = service.presentations().batchUpdate(
        presentationId=presentation_id, body={"requests": requests}
    ).execute()

    return response.get("replies")[0]["createSlide"]["objectId"]



def add_text_box(service, presentation_id, slide_id, text_box_id):
    """Adds a text box to the specified slide.

    Args:
        service: The Google Slides API service object.
        presentation_id: The ID of the presentation.
        slide_id: The ID of the slide where the text box will be added.
        text_box_id: The objectId of the text box.

    Returns:
        None
    """
    requests = [
        {
            "createShape": {
                "objectId": text_box_id,
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {
                        "height": {"magnitude": 2000000, "unit": "EMU"},
                        "width": {"magnitude": 4000000, "unit": "EMU"},
                    },
                    "transform": {
                        "scaleX": 1,
                        "scaleY": 1,
                        "translateX": 100000,
                        "translateY": 100000,
                        "unit": "EMU",
                    },
                },
            }
        }
    ]

    service.presentations().batchUpdate(
        presentationId=presentation_id, body={"requests": requests}
    ).execute()



def insert_text_into_box(service, presentation_id, text_box_id, text_to_add):
    """Inserts text into the specified text box.

    Args:
        service: The Google Slides API service object.
        presentation_id: Google SlidesのID
        text_box_id: 任意のテキストボックスのobjectId
        text_to_add: 追加したいテキスト

    Returns:
        None
    """
    requests = [
        {
            "insertText": {
                "objectId": text_box_id,
                "text": text_to_add,
                "insertionIndex": 0,
            }
        }
    ]

    service.presentations().batchUpdate(
        presentationId=presentation_id, body={"requests": requests}
    ).execute()




def add_new_slide_with_text(service, presentation_id, text_to_add):
    """新しいslideを追加し、テキストボックスを追加 (add_text_box)、テキストを挿入 (insert_text_into_box) する。


    Args:
        service: The Google Slides API service object.
        presentation_id: Google SlidesのID
        text_to_add: 追加したいテキスト


    Returns:
        None
    """
    slide_id = create_new_slide(service, presentation_id)  # 新しいスライドを作成
    text_box_id = "MyTextBox"  # 固定ID（ユニークにしたい場合はランダム生成）

    add_text_box(service, presentation_id, slide_id, text_box_id)  # テキストボックスを追加
    insert_text_into_box(service, presentation_id, text_box_id, text_to_add)  # テキストを挿入

    print(f"New slide created with ID: {slide_id}")

 

def main():

    SCOPES = ["https://www.googleapis.com/auth/presentations"]
    PRESENTATION_ID = "1GButmtqvj5LT8TzLexFRnA1-5tDsW5ft-S7E4_Sy8AE"
    creds = None

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                r"C:\Users\nepia\OneDrive\デスクトップ\Google_Slides\credentials.json",
                SCOPES
            )
            creds = flow.run_local_server(port=0)
    
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("slides", "v1", credentials=creds)
        add_new_slide_with_text(service, PRESENTATION_ID, "あっちょんぶりけ～！！ッ")
        presentation = (
        service.presentations().get(presentationId=PRESENTATION_ID).execute()
        )
        slides = presentation.get("slides")

        print(f"The presentation contains {len(slides)} slides:")
        for i, slide in enumerate(slides):
            page_elements = slide.get("pageElements")
            # print(f"Slide {i+1}: {page_elements}")
            
            if page_elements:
                for j, element in enumerate(page_elements):
                    extracted_text = extract_text_from_element(element)
                    if extracted_text:
                        print(f"Slide {i+1}: {extracted_text}")
                    else:
                        print(f"  Element {j + 1} has no text.")
            else:
                print(f"  Slide {i + 1} has no elements.")
            
    except HttpError as err:
        print(err)


if __name__ == "__main__":
    main()