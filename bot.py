import os
from flask import Flask, request, send_file
from telegram import Update, Bot
from telegram.ext import Dispatcher, CommandHandler, MessageHandler, Filters
from io import BytesIO
from docx import Document
from docx.shared import RGBColor

# Список имён иноагентов для зачеркивания (пример)
FOREIGN_AGENT_NAMES = ["Иван Иванов", "Мария Петрова", "John Smith"]

TOKEN = os.getenv("BOT_TOKEN")
if not TOKEN:
    raise ValueError("Не найден токен BOT_TOKEN в переменных окружения")

bot = Bot(token=TOKEN)
app = Flask(__name__)
dispatcher = Dispatcher(bot, None, workers=0)

def start(update: Update, context=None):
    update.message.reply_text("Привет! Отправь мне Word-файл, я зачеркиваю в нём имена из списка.")

def handle_doc(update: Update, context=None):
    file = update.message.document.get_file()
    file_bytes = BytesIO()
    file.download(out=file_bytes)
    file_bytes.seek(0)

    doc = Document(file_bytes)

    for para in doc.paragraphs:
        original_text = para.text
        para.clear()  # Очищаем параграф, чтобы вручную вставить отформатированные куски

        i = 0
        while i < len(original_text):
            match_found = False
            for name in FOREIGN_AGENT_NAMES:
                if original_text[i:i+len(name)] == name:
                    run = para.add_run(name)
                    run.font.color.rgb = RGBColor(255, 0, 0)  # Красный цвет
                    i += len(name)
                    match_found = True
                    break
            if not match_found:
                run = para.add_run(original_text[i])
                i += 1

    output = BytesIO()
    doc.save(output)
    output.seek(0)

    update.message.reply_document(document=output, filename="highlighted.docx")

# Регистрируем обработчики
dispatcher.add_handler(CommandHandler("start", start))
dispatcher.add_handler(MessageHandler(Filters.document.mime_type("application/vnd.openxmlformats-officedocument.wordprocessingml.document"), handle_doc))

@app.route(f"/{TOKEN}", methods=["POST"])
def webhook():
    data = request.get_json(force=True)
    update = Update.de_json(data, bot)
    dispatcher.process_update(update)
    return "ok"

@app.route("/")
def index():
    return "Bot is running"

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"Starting app on port {port}")
    app.run(host="0.0.0.0", port=port)
