import openai
import os
from flask import Flask, request, jsonify, render_template, session, send_file, url_for
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
import requests
import pandas as pd
from fpdf import FPDF
from docx import Document
from pptx import Presentation
import tempfile

app = Flask(__name__)
app.secret_key = os.urandom(24)  # セッションを管理するための秘密鍵

# 環境変数からAPIキーやエンドポイントを取得
openai.api_type = "azure"
openai.api_base = os.getenv("OPENAI_API_BASE")
openai.api_version = "2024-08-01-preview"
openai.api_key = os.getenv("OPENAI_API_KEY")
deployment_name = os.getenv("OPENAI_DEPLOYMENT_NAME")

# Azure Cognitive Search の設定
search_service_endpoint = os.getenv("AZURE_SEARCH_ENDPOINT")
search_service_key = os.getenv("AZURE_SEARCH_KEY")
index_name = "vector-1730110777868"  # 使用するインデックス名を指定

# SearchClient の設定
search_client = SearchClient(
    endpoint=search_service_endpoint,
    index_name=index_name,
    credential=AzureKeyCredential(search_service_key)
)

# Computer Vision API の設定
vision_subscription_key = os.getenv("VISION_API_KEY")
vision_endpoint = os.getenv("VISION_ENDPOINT")

# ホームページを表示するルート
@app.route('/')
def index():
    session.clear()  # ブラウザを閉じたらセッションをリセット
    return render_template('index.html', chat_history=[])

# OCR機能を提供するルート
@app.route('/ocr', methods=['POST'])
def ocr_endpoint():
    image_url = request.json.get("image_url")  # JSONデータから画像URLを取得
    if not image_url:
        return jsonify({"error": "No image URL provided"}), 400

    try:
        text = ocr_image(image_url)
        return jsonify({"text": text})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def ocr_image(image_url):
    """
    Azure Computer Vision APIを使って画像のOCRを実行する関数
    """
    ocr_url = vision_endpoint + "/vision/v3.2/ocr"
    headers = {"Ocp-Apim-Subscription-Key": vision_subscription_key}
    params = {"language": "ja", "detectOrientation": "true"}
    data = {"url": image_url}

    # OCRリクエストを送信
    response = requests.post(ocr_url, headers=headers, params=params, json=data)
    response.raise_for_status()

    # 結果を取得し、テキスト部分を抽出
    ocr_results = response.json()
    text_results = []
    for region in ocr_results.get("regions", []):
        for line in region.get("lines", []):
            line_text = " ".join([word["text"] for word in line["words"]])
            text_results.append(line_text)
    return "\n".join(text_results)

# ファイル出力用のルート
@app.route('/download_output', methods=['POST'])
def download_output():
    try:
        prompt = request.form.get('prompt', '')
        response_content = session.get('response_content', 'No response available')
        
        # 一時ファイルを生成して内容を保存
        with tempfile.NamedTemporaryFile(delete=False, suffix=".txt") as temp_file:
            temp_file.write(f"プロンプト:\n{prompt}\n\n応答:\n{response_content}".encode('utf-8'))
            temp_file_path = temp_file.name

        # クライアントにファイルを送信
        return send_file(temp_file_path, as_attachment=True, download_name="output.txt")

    except Exception as e:
        return f"ファイルの出力中にエラーが発生しました: {str(e)}"

# ファイルとプロンプトを処理するルート
@app.route('/process_files_and_prompt', methods=['POST'])
def process_files_and_prompt():
    files = request.files.getlist('files')  # 複数ファイルの取得
    prompt = request.form.get('prompt', '')

    if 'chat_history' not in session:
        session['chat_history'] = []  # チャット履歴を初期化

    # チャット履歴を基にAIへ送信するメッセージを構成
    messages = [{"role": "system", "content": "あなたは有能なアシスタントです。"}]
    
    # チャット履歴を追加
    for entry in session['chat_history']:
        messages.append({"role": "user", "content": entry['user']})
        messages.append({"role": "assistant", "content": entry['assistant']})

    # ファイルがアップロードされている場合の処理
    file_data_text = []
    try:
        for file in files:
            if file and file.filename.endswith('.xlsx'):  # Excelファイルの処理
                df = pd.read_excel(file, engine='openpyxl')
                columns = df.columns.tolist()
                columns_text = " | ".join(columns)
                rows_text = [" | ".join([str(item) for item in row.tolist()]) for index, row in df.iterrows()]
                df_text = f"ファイル名: {file.filename}\n{columns_text}\n" + "\n".join(rows_text)
                file_data_text.append(df_text)
            elif file and file.filename.endswith('.pdf'):  # PDFファイルの処理
                text = ocr_image(file)
                pdf_text = f"ファイル名: {file.filename}\n{text}"
                file_data_text.append(pdf_text)
            elif file and file.filename.endswith('.docx'):  # Wordファイルの処理
                doc = Document(file)
                word_text = "\n".join([para.text for para in doc.paragraphs])
                word_text = f"ファイル名: {file.filename}\n{word_text}"
                file_data_text.append(word_text)
            elif file and file.filename.endswith('.pptx'):  # PPTファイルの処理
                ppt = Presentation(file)
                ppt_text = []
                for slide in ppt.slides:
                    for shape in slide.shapes:
                        if hasattr(shape, "text"):
                            ppt_text.append(shape.text)
                ppt_text = f"ファイル名: {file.filename}\n" + "\n".join(ppt_text)
                file_data_text.append(ppt_text)
            else:
                continue  # 不正なファイル形式の場合、処理をスキップ

        # ファイルがある場合、内容を結合して送信するデータに追加
        if file_data_text:
            file_contents = "\n\n".join(file_data_text)
            input_data = f"アップロードされたファイルの内容は次の通りです:\n{file_contents}\nプロンプト: {prompt}"
        else:
            input_data = f"プロンプトのみが入力されました:\nプロンプト: {prompt}"

        # AIにプロンプトとファイル内容を送信して応答を取得
        response = openai.ChatCompletion.create(
            engine=deployment_name,
            messages=messages,
            max_tokens=2000
        )

        # 応答内容を取得し、履歴に追加
        response_content = response['choices'][0]['message']['content']
        session['response_content'] = response_content  # ファイル出力用にセッションへ保存
        session['chat_history'].append({
            'user': input_data,
            'assistant': response_content
        })

        return render_template('index.html', chat_history=session['chat_history'])

    except Exception as e:
        return f"エラーが発生しました: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True)
