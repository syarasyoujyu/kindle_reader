import pyautogui
from PIL import Image, ImageEnhance, ImageOps
import pytesseract
import time
import os
from fpdf import FPDF
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2 import service_account
import io
from dotenv import load_dotenv


load_dotenv()

# --- 設定 ---
pytesseract.pytesseract.tesseract_cmd = os.getenv("TESSARACT_EXE_PATH", "")
SERVICE_ACCOUNT_FILE = os.getenv("SERVICE_ACCOUNT_FILE","")
SCOPES = os.getenv("SCOPES", "").split(",")
SHARE_WITH_EMAIL = os.getenv("SHARE_WITH_EMAIL","")
DOC_TITLE = os.getenv("DOC_TITLE","ツァラトゥストラ")
PAGE_NUM=int(os.getenv("PAGE_NUM","501"))

# --- Google認証とサービスの初期化 ---
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('drive', 'v3', credentials=creds)

# --- ファイルを指定のユーザーに共有する関数 ---
def share_file(file_id, user_email):
    permission = {
        'type': 'user',
        'role': 'reader',
        'emailAddress': user_email
    }
    try:
        service.permissions().create(
            fileId=file_id,
            body=permission,
            fields='id',
            sendNotificationEmail=False
        ).execute()
        print(f'🔓 {user_email} に閲覧権限を付与しました')
    except Exception as e:
        print(f"❌ 共有設定エラー: {e}")

# --- メイン処理 ---
def capture_multiple_pages(x1, y1, x2, y2, num_pages, interval_sec, doc_title):
    all_text = ''
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    image_paths = []

    for i in range(num_pages):
        time.sleep(interval_sec + 1)  # 念のため、待機時間を少し増やす
        temp_image_path = f'temp_{i}.png'
        screenshot = pyautogui.screenshot(region=(x1, y1, x2 - x1, y2 - y1))
        screenshot.save(temp_image_path)
        image_paths.append(temp_image_path)

        image = Image.open(temp_image_path).convert('L')
        image = ImageEnhance.Contrast(image).enhance(2.0)

        try:
            text = pytesseract.image_to_string(image, lang='jpn_vert', config='--psm 5') # 縦書き、PSMを5に変更
        except pytesseract.TesseractError as e:
            print(f"OCR error: {e}")
            text = ""

        all_text += f'\n\n--- Page {i + 1} ---\n{text}'

        pdf.add_page()
        pdf.image(temp_image_path, x=10, y=10, w=180)

        pyautogui.press('left')

    # Googleドキュメントアップロード（メモリ上で）
    doc_metadata = {'name': doc_title, 'mimeType': 'application/vnd.google-apps.document'}
    doc_stream = io.BytesIO(all_text.encode('utf-8'))
    doc_media = MediaIoBaseUpload(doc_stream, mimetype='text/plain')
    doc_file = service.files().create(body=doc_metadata, media_body=doc_media, fields='id').execute()
    doc_id = doc_file.get("id")  # ドキュメントIDを取得
    print(f'✅ Googleドキュメントアップロード完了: {doc_id}')
    share_file(doc_id, SHARE_WITH_EMAIL)
    doc_url = f'https://docs.google.com/document/d/{doc_id}/edit'  # URLを生成
    print(f'📄 GoogleドキュメントURL: {doc_url}')  # URLを表示

    # PDF出力 → アップロード
    temp_pdf_path = f'{doc_title}.pdf'
    pdf.output(temp_pdf_path)
    with open(temp_pdf_path, 'rb') as f:
        pdf_media = MediaIoBaseUpload(f, mimetype='application/pdf')
        pdf_metadata = {'name': doc_title + '.pdf', 'mimeType': 'application/pdf'}
        pdf_file = service.files().create(body=pdf_metadata, media_body=pdf_media, fields='id').execute()
    pdf_id = pdf_file.get("id")  # PDFのIDを取得
    print(f'✅ PDFアップロード完了: {pdf_id}')
    share_file(pdf_id, SHARE_WITH_EMAIL)
    pdf_url = f'https://drive.google.com/file/d/{pdf_id}/view?usp=sharing'  # PDFのURLを生成
    print(f'📄 PDF URL: {pdf_url}')  # PDFのURLを表示

    # 一時ファイル削除
    for path in image_paths + [temp_pdf_path]:
        try:
            os.remove(path)
        except Exception as e:
            print(f'❌ 削除失敗: {path} → {e}')

# --- 実行 ---
if __name__ == '__main__':
    capture_multiple_pages(
        x1=664, y1=93,  # スクリーンショットする左上
        x2=1302, y2=1005,  # スクリーンショットする右下
        num_pages=PAGE_NUM,
        interval_sec=1,
        doc_title=DOC_TITLE

    )