import os
from flask import Flask, request, jsonify, render_template, session, send_file, url_for
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
import openai
import pandas as pd
from io import BytesIO
from flask_session import Session
import requests
from fpdf import FPDF

# Flaskアプリの設定
app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config['SESSION_TYPE'] = 'filesystem'
Session(app)

# 環境変数の設定
openai.api_type = "azure"
openai.api_base = os.getenv("OPENAI_API_BASE", "https://example.openai.azure.com")
openai.api_version = "2024-08-01-preview"
openai.api_key = os.getenv("OPENAI_API_KEY", "your-api-key")
deployment_name = os.getenv("OPENAI_DEPLOYMENT_NAME", "default-deployment")

search_service_endpoint = os.getenv("AZURE_SEARCH_ENDPOINT", "https://search-service.azure.com")
search_service_key = os.getenv("AZURE_SEARCH_KEY", "your-search-key")
index_name = "vector-1730110777868"

vision_subscription_key = os.getenv("VISION_API_KEY", "your-vision-key")
vision_endpoint = os.getenv("VISION_ENDPOINT", "https://vision.azure.com")

# Azure Search クライアントの設定
search_client = SearchClient(
    endpoint=search_service_endpoint,
    index_name=index_name,
    credential=AzureKeyCredential(search_service_key)
)

@app.route('/')
def index():
    session.clear()
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
        relevant_docs = "\n".join([doc['chunk'] for doc in search_results])

        input_data_with_search = f"{input_data}\n\n関連ドキュメント:\n{relevant_docs}"
        messages.append({"role": "user", "content": input_data_with_search})

        response = openai.ChatCompletion.create(
            engine=deployment_name,
            messages=messages,
            max_tokens=2000
        )

        response_content = response['choices'][0]['message']['content']
        session['response_content'] = response_content

        format_decision = determine_output_format(response_content)
        file_data, mimetype, filename = generate_file(response_content, format_decision)

        session['chat_history'].append({
            'user': input_data_with_search,
            'assistant': response_content
        })

        return send_file(file_data, mimetype=mimetype, as_attachment=True, download_name=filename)

    except Exception as e:
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
        return jsonify({"error": f"Excelファイルの出力中にエラーが発生しました: {str(e)}"}), 500

def determine_output_format(response_content):
    """
    AIを用いて出力形式を判断する関数
    """
    prompt = f"""
    以下のテキストがどの形式に適しているかを判断してください。可能な形式としては次のものがあります:
    1. 表形式（ExcelまたはCSV）
    2. 自然言語（PDFまたはテキスト）
    3. コードスニペット（Python, JSONなど）

    応答の内容:
    「{response_content}」

    適した形式と理由を教えてください。
    """
    response = openai.Completion.create(
        engine=deployment_name,
        prompt=prompt,
        max_tokens=150,
        temperature=0.7
    )
    return response['choices'][0]['text'].strip()

def generate_file(response_content, format_decision):
    """
    フォーマット判断に基づき適切なファイルを生成する関数
    """
    if "表形式" in format_decision:
        rows = [row.split(",") for row in response_content.split("\n") if row]
        df = pd.DataFrame(rows[1:], columns=rows[0])
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
        output.seek(0)
        return output, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "output.xlsx"

    elif "自然言語" in format_decision:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        for line in response_content.split("\n"):
            pdf.cell(200, 10, txt=line, ln=True, align='L')
        output = BytesIO()
        pdf.output(output)
        output.seek(0)
        return output, "application/pdf", "output.pdf"

    elif "コードスニペット" in format_decision:
        output = BytesIO()
        output.write(response_content.encode('utf-8'))
        output.seek(0)
        return output, "text/plain", "output.txt"

    else:
        output = BytesIO()
        output.write(response_content.encode('utf-8'))
        output.seek(0)
        return output, "text/plain", "output.txt"

if __name__ == '__main__':
    app.run(debug=True)
