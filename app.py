import os
import json
import datetime
import gspread
from flask import Flask, request, abort
from dotenv import load_dotenv
from oauth2client.service_account import ServiceAccountCredentials
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
from openai import OpenAI

# Load environment variables
load_dotenv()
LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET")
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
SHEET_ID = os.getenv("SHEET_ID")
GSPREAD_JSON = os.getenv("GSPREAD_JSON")  # Google Service Account JSON content

# Initialize Flask and LINE
app = Flask(__name__)
line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# Initialize OpenAI Client
openai_client = OpenAI(api_key=OPENAI_API_KEY)

# Setup Google Sheet
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(json.loads(GSPREAD_JSON), scope)
gs_client = gspread.authorize(creds)
sheet = gs_client.open_by_key(SHEET_ID).worksheet("履約主表")

# GPT Helper
def ask_gpt(text):
    prompt = f"""
你是一個土木工程履約助理，請從以下語句中分析出：
1. 案名（可模糊比對）
2. 工作項目（可模糊比對）
3. 動作（提送 或 核定）
4. 日期（格式化為 2025-03-05）
請用 JSON 格式輸出，例如：
{{"案名": "南區開口工程", "工作項目": "提送工程預算書圖", "動作": "提送", "日期": "2025-03-05"}}
語句：{text}
"""
    chat_completion = openai_client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        temperature=0
    )
    return json.loads(chat_completion.choices[0].message.content.strip())

# Webhook 接收器
@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'OK'

# 處理文字訊息
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    text = event.message.text
    if "威威1號" not in text:
        return

    try:
        gpt_result = ask_gpt(text)
    except Exception as e:
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=f"⚠️ GPT 發生錯誤：{e}")
        )
        return

    item = gpt_result.get("工作項目")
    action = gpt_result.get("動作")
    date = gpt_result.get("日期")

    if not item or not action or not date:
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="⚠️ 語句不完整，請確認是否有案名／項目／日期")
        )
        return

    # 尋找 Google Sheet 中的對應行
    data = sheet.get_all_records()
    matched = False
    for idx, row in enumerate(data):
        if item in row["工作項目"]:
            col = "提送日" if action == "提送" else "核定日"
            sheet.update_cell(idx + 2, list(row.keys()).index(col) + 1, date)
            matched = True
            break

    if matched:
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=f"✅ 已更新『{item}』的{action}日為 {date}")
        )
    else:
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=f"❌ 找不到對應的工作項目『{item}』")
        )

# 執行主程式
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
