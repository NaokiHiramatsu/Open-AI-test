import openai
import os
from flask import Flask, request, jsonify, render_template, session, send_file, url_for
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
import pandas as pd
from fpdf import FPDF
from docx import Document
from pptx import Presentation
import tempfile
import requests

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
index_name = "vector-1730110777868"

search_client = SearchClient(
    endpoint=search_service_endpoint,
    index_name=index_name,
    credential=AzureKeyCredential(search_service_key)
)

# Computer Vision API の設定
vision_subscription_key = os.getenv("VISION_API_KEY")
vision_endpoint = os.getenv("VISION_ENDPOINT")

@app.route('/')
def index():
    session.clear()  # セッションをリセット
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

def ocr_pdf(file):
    ocr_url = vision_endpoint + "/vision/v3.2/read/analyze"
    headers = {"Ocp-Apim-Subscription-Key": vision_subscription_key, "Content-Type": "application/pdf"}

    response = requests.post(ocr_url, headers=headers, data=file.read())
    response.raise_for_status()

    operation_url = response.headers["Operation-Location"]
    analysis = {}
    while not "analyzeResult" in analysis:
        response_final = requests.get(operation_url, headers={"Ocp-Apim-Subscription-Key": vision_subscription_key})
        analysis = response_final.json()

    text_results = []
    for read_result in analysis["analyzeResult"]["readResults"]:
        for line in read_result["lines"]:
            text_results.append(line["text"])
    return "\n".join(text_results)

def generate_excel(content):
    try:
        temp_excel_path = tempfile.mktemp(suffix=".xlsx")
        if isinstance(content, list):
            df = pd.DataFrame(content)
        else:
            df = pd.DataFrame({"Content": [content]})
        df.to_excel(temp_excel_path, index=False)
        return temp_excel_path
    except Exception as e:
        print(f"Excel generation failed: {e}")
        return None

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
    for file in files:
        if file.filename.endswith('.xlsx'):
            df = pd.read_excel(file, engine='openpyxl')
            file_data_text.append(df.to_csv(index=False))
        elif file.filename.endswith('.pdf'):
            file_data_text.append(ocr_pdf(file))
        elif file.filename.endswith('.docx'):
            doc = Document(file)
            file_data_text.append("\n".join([para.text for para in doc.paragraphs]))
        elif file.filename.endswith('.pptx'):
            ppt = Presentation(file)
            ppt_text = []
            for slide in ppt.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        ppt_text.append(shape.text)
            file_data_text.append("\n".join(ppt_text))

    file_content_combined = "\n\n".join(file_data_text)
    input_data = f"アップロードされたファイルの内容:\n{file_content_combined}\nプロンプト: {prompt}"
    messages.append({"role": "user", "content": input_data})

    search_results = search_client.search(search_text=prompt, top=3)
    relevant_docs = "\n".join([doc['chunk'] for doc in search_results])
    messages.append({"role": "user", "content": f"以下に基づいて回答してください:\n{relevant_docs}"})

    response = openai.ChatCompletion.create(
        engine=deployment_name,
        messages=messages,
        max_tokens=2000
    )
    response_content = response['choices'][0]['message']['content']
    session['chat_history'].append({'user': input_data, 'assistant': response_content})

    if "Excel" in response_content:
        excel_path = generate_excel(response_content)
        if excel_path:
            return send_file(excel_path, as_attachment=True, download_name="output.xlsx")

    return render_template('index.html', chat_history=session['chat_history'])

if __name__ == '__main__':
    app.run(debug=True)
