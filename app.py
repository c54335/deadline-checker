import os
import json
import datetime
import openai
import gspread
from flask import Flask, request, abort
from dotenv import load_dotenv
from oauth2client.service_account import ServiceAccountCredentials
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage

# Load environment variables
load_dotenv()
LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET")
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
SHEET_ID = os.getenv("SHEET_ID")
GSPREAD_JSON = os.getenv("GSPREAD_JSON")  # Google Service Account JSON content

app = Flask(__name__)

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)
openai.api_key = OPENAI_API_KEY

# Setup Google Sheet
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(json.loads(GSPREAD_JSON), scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(SHEET_ID).worksheet("履約主表")

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
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        temperature=0
    )
    try:
        return json.loads(response.choices[0].message.content.strip())
    except:
        return None

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers.get('X-Line-Signature', '')
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    text = event.message.text

    if event.source.type in ["group", "room"]:
        if hasattr(event.message, "mention") and event.message.mention:
            for mentionee in event.message.mention.mentionees:
                if mentionee.is_self:
                    break
            else:
                return
        else:
            return

    if "@威威1號" not in text:
        return

    gpt_result = ask_gpt(text)
    if not gpt_result:
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="⚠️ 無法理解這句話，請再試一次")
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

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)

