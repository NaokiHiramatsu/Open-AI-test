# Python 3.9のスリムバージョンをベースに使用
FROM python:3.9-slim

# Tesseract OCRと日本語パッケージをインストール
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-jpn \
    libgl1-mesa-glx \
    libglib2.0-0 \
    && apt-get clean

# アプリケーションの依存関係をインストール
COPY requirements.txt /app/requirements.txt
RUN pip install -r /app/requirements.txt

# アプリケーションのコードをコピー
COPY . /app
WORKDIR /app

# Flaskアプリのポート設定（デフォルトで5000）
EXPOSE 5000

# アプリケーションを起動
CMD ["python", "app.py"]
