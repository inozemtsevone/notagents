import os
from flask import Flask, request, send_file
from telegram import Update, Bot
from telegram.ext import Dispatcher, CommandHandler, MessageHandler, Filters
from io import BytesIO
from docx import Document

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

    # Проходим по всем абзацам и заменяем имена из списка
    for para in doc.paragraphs:
        for name in FOREIGN_AGENT_NAMES:
            if name in para.text:
                # Зачеркиваем имя (ставим форматирование strikethrough)
                run = para.add_run()
                run.text = ""
                # Удалим и заменим текст с именем зачеркиванием:
                para.text = para.text.replace(name, f"̶{name}̶")  # Это упрощенный способ (не гарантирован)

    # Сохраняем изменённый документ в память
    output = BytesIO()
    doc.save(output)
    output.seek(0)

    update.message.reply_document(document=output, filename="edited.docx")

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
