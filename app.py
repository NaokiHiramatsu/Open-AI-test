import os
from flask import Flask, request, jsonify, render_template, session, send_file, url_for
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
import openai
import requests
import pandas as pd
from io import BytesIO
from flask_session import Session

# Flaskアプリの設定
app = Flask(__name__)
app.secret_key = os.urandom(24)

# セッション設定
app.config['SESSION_TYPE'] = 'filesystem'
Session(app)

# 環境変数の取得とデフォルト値の設定
openai.api_base = os.getenv("OPENAI_API_BASE", "https://example.openai.azure.com")
openai.api_key = os.getenv("OPENAI_API_KEY", "your-api-key")
openai.api_version = "2024-08-01-preview"
deployment_name = os.getenv("OPENAI_DEPLOYMENT_NAME", "default-deployment")

search_service_endpoint = os.getenv("AZURE_SEARCH_ENDPOINT", "https://search-service.azure.com")
search_service_key = os.getenv("AZURE_SEARCH_KEY", "your-search-key")
index_name = os.getenv("AZURE_SEARCH_INDEX_NAME", "default-index")

vision_subscription_key = os.getenv("VISION_API_KEY", "your-vision-key")
vision_endpoint = os.getenv("VISION_ENDPOINT", "https://vision.azure.com")

# Azure Search クライアントの設定
try:
    search_client = SearchClient(
        endpoint=search_service_endpoint,
        index_name=index_name,
        credential=AzureKeyCredential(search_service_key)
    )
except Exception as e:
    print(f"Azure Search Client初期化エラー: {str(e)}")

@app.route('/')
def index():
    session.clear()  # セッションの初期化
    return render_template('index.html', chat_history=[])

@app.route('/ocr', methods=['POST'])
def ocr_endpoint():
    image_url = request.json.get("image_url")
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

    response = requests.post(ocr_url, headers=headers, params=params, json=data)
    response.raise_for_status()

    ocr_results = response.json()
    text_results = []
    for region in ocr_results.get("regions", []):
        for line in region.get("lines", []):
            line_text = " ".join([word["text"] for word in line["words"]])
            text_results.append(line_text)
    return "\n".join(text_results)

@app.route('/process_files_and_prompt', methods=['POST'])
def process_files_and_prompt():
    files = request.files.getlist('files')
    prompt = request.form.get('prompt', '')

    if 'chat_history' not in session:
        session['chat_history'] = []

    messages = [{"role": "system", "content": "あなたは有能なアシスタントです。"}]
    for entry in session['chat_history']:
        messages.append({"role": "user", "content": entry['user']})
        messages.append({"role": "assistant", "content": entry['assistant']})

    file_data_text = []
    try:
        for file in files:
            if file and file.filename.endswith('.xlsx'):
                df = pd.read_excel(file, engine='openpyxl')
                columns = df.columns.tolist()
                rows_text = df.to_string(index=False)
                df_text = f"ファイル名: {file.filename}\n列: {columns}\n内容:\n{rows_text}"
                file_data_text.append(df_text)

        file_contents = "\n\n".join(file_data_text) if file_data_text else "なし"
        input_data = f"アップロードされたファイル内容:\n{file_contents}\nプロンプト:\n{prompt}"

        search_results = search_client.search(search_text=prompt, top=3)
        relevant_docs = "\n".join([doc['content'] for doc in search_results])

        input_data_with_search = f"{input_data}\n\n関連ドキュメント:\n{relevant_docs}"
        messages.append({"role": "user", "content": input_data_with_search})

        response = openai.ChatCompletion.create(
            engine=deployment_name,
            messages=messages,
            max_tokens=2000
        )

        if not response or 'choices' not in response or not response['choices']:
            raise ValueError("OpenAI APIから有効な応答がありません。")

        response_content = response['choices'][0]['message']['content']

        session['response_content'] = response_content
        download_link = f"<a href='{url_for('download_excel')}' target='_blank'>ファイルダウンロード</a>"
        full_response = f"{response_content}<br>{download_link}"
        session['chat_history'].append({
            'user': input_data_with_search,
            'assistant': full_response
        })

        return render_template('index.html', chat_history=session['chat_history'])

    except Exception as e:
        app.logger.error(f"エラー: {str(e)}", exc_info=True)
        return jsonify({"error": f"エラーが発生しました: {str(e)}"}), 500

@app.route('/download_excel', methods=['GET'])
def download_excel():
    try:
        response_content = session.get('response_content', None)
        if not response_content:
            raise ValueError("セッションに応答データがありません。")

        rows = [row.split(",") for row in response_content.split("\n") if row]
        if len(rows) < 2 or not all(len(row) == len(rows[0]) for row in rows):
            raise ValueError("応答のフォーマットが正しくありません。")

        df = pd.DataFrame(rows[1:], columns=rows[0])

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
        output.seek(0)

        return send_file(output, as_attachment=True, download_name="output.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        app.logger.error(f"Excel出力エラー: {str(e)}", exc_info=True)
        return jsonify({"error": f"Excelファイルの出力中にエラーが発生しました: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000, debug=True)
