import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError



def add_text_to_slide(service, params):
    """
    指定したスライド上の指定した図形にテキストを追加する関数です。
    
    パラメータは以下のような辞書で渡します:
        {
            "presentation_id": "プレゼンテーションのID",
            "shape_id": "図形のID",
            "text": "追加するテキスト",
            "insertion_index": 挿入開始位置 (オプション, デフォルトは0)
        }
    """
    insertion_index = params.get("insertion_index", 0)
    requests = [
        {
            "insertText": {
                "objectId": params["shape_id"],
                "insertionIndex": insertion_index,
                "text": params["text"]
            }
        }
    ]
    body = {"requests": requests}
    response = service.presentations().batchUpdate(
        presentationId=params["presentation_id"],
        body=body
    ).execute()
    return response



def extract_text_from_element(element):
    """
    テキストボックスなどの shape 要素の場合、その中のテキストを抽出する関数です。
    API のレスポンスでは、shape の中に "text" キーがあり、さらに "textElements" の配列が含まれています。
    その中で、"textRun" キーが存在する要素から実際のテキストを取り出します。
    """
    text_content = ""
    shape = element.get("shape")
    if shape:
        text_obj = shape.get("text", {})
        for te in text_obj.get("textElements", []):
            # textRunが存在する場合、その中の content を取得
            if "textRun" in te and "content" in te["textRun"]:
                text_content += te["textRun"]["content"]
    return text_content.strip()



def write_slides_to_markdown(presentation, output_file):
    """
    プレゼンテーション内の各スライドおよびその要素の情報を、Markdown テーブルとしてファイルに保存する関数です。
    
    テーブルの形式（1行につき1要素）:
    
    | Slide # | Slide ID | Element # | Element ID | Element Text |
    |---------|----------|-----------|------------|--------------|
    | 1       | slide_id | 1         | id1        | テキスト内容 |
    
    :param presentation: Google Slides API から取得したプレゼンテーションオブジェクト
    :param output_file: 書き出し先の Markdown ファイルパス
    """
    slides = presentation.get("slides", [])
    lines = []
    # ヘッダー行
    lines.append("| Page | Slide ID | Element # | Element ID | Element Text |")
    lines.append("|---------|----------|-----------|------------|--------------|")
    
    for slide_index, slide in enumerate(slides):
        slide_id = slide.get("objectId", "不明")
        page_elements = slide.get("pageElements", [])
        # ページ内に要素がない場合も1行出力
        if not page_elements:
            lines.append(f"| {slide_index + 1} | {slide_id} | - | - | - |")
        else:
            for element_index, element in enumerate(page_elements):
                element_id = element.get("objectId", "なし")
                element_text = extract_text_from_element(element)
                # Markdown のテーブルでは "|" で区切るので、改行などが入っている場合は適宜整形するのが望ましい
                element_text = element_text.replace("\n", " ")
                lines.append(
                    f"| {slide_index + 1} | {slide_id} | {element_index + 1} | {element_id} | {element_text} |"
                )
    
    with open(output_file, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    
    print(f"Markdownファイル '{output_file}' に保存しました。")



def main():
    """Slides API の基本的な使用法を示します。
    サンプルプレゼンテーションのスライド数と要素数を出力します。
    """
    
    # これらのスコープを変更する場合は、token.json ファイルを削除してください。
    SCOPES = ["https://www.googleapis.com/auth/presentations"]

    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        
    # 利用可能な (有効な) 認証情報がない場合は、ユーザーにログインさせます。
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # 正しい相対パスでの指定
            credentials_path = r"C:\Users\nepia\OneDrive\デスクトップ\Google_Slides\credentials.json"
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials_path, SCOPES
            )
            creds = flow.run_local_server(port=0)
        # 次回の実行のために認証情報を保存します。
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("slides", "v1", credentials=creds)

        # Google SlidesのID
        PRESENTATION_ID = "1GButmtqvj5LT8TzLexFRnA1-5tDsW5ft-S7E4_Sy8AE"
        presentation = (
            service.presentations().get(presentationId=PRESENTATION_ID).execute()
        )
        
        # slide_id, element_id, textなどをマークダウンファイルとして出力する 
        slides = presentation.get("slides")
        presentation = service.presentations().get(presentationId=PRESENTATION_ID).execute()
        write_slides_to_markdown(presentation, "slides.md")        
        
    except HttpError as err:
        print(err)


if __name__ == "__main__":
    main()