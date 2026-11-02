"""
Telegram Bot для OpenJarvis + NVIDIA NIM с поддержкой анализа файлов.
Бот: @MoynoviuJARVIS_bot
Поддерживает: Excel (.xlsx, .xls), CSV, PDF, Word (.docx), TXT.
"""

import asyncio
import logging
import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

# Поддержка UTF-8 в Windows-консоли
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

import httpx
from telegram import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    ReplyKeyboardMarkup,
    Update,
)
from telegram.constants import ChatAction, ParseMode
from telegram.ext import (
    ApplicationBuilder,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

from file_processor import extract_file_content

# Настройка логирования (в консоль и в файл bot.log)
log_path = Path(__file__).resolve().parent / "bot.log"
log_formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

file_handler = logging.FileHandler(log_path, encoding="utf-8")
file_handler.setFormatter(log_formatter)
file_handler.setLevel(logging.INFO)

console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(log_formatter)
console_handler.setLevel(logging.INFO)

logging.basicConfig(
    level=logging.INFO,
    handlers=[file_handler, console_handler]
)
logging.getLogger("httpx").setLevel(logging.WARNING)
logger = logging.getLogger("JarvisBot")

# Константы
NIM_BASE_URL = "https://integrate.api.nvidia.com/v1"
MODEL_NAME = "nvidia/nemotron-3.5-bolt-e4b"
SYSTEM_PROMPT = (
    "Ты — Джарвис (J.A.R.V.I.S.), продвинутый персональный ИИ-ассистент и бизнес-аналитик. "
    "Ты открываешь, читаешь и глубоко анализируешь финансовые файлы и таблицы маркетплейсов (Wildberries, Ozon), "
    "рассчитываешь экономику, комиссии, логистику, рекламные расходы, прибыль и рентабельность. "
    "Ты работаешь напрямую с данными таблиц, поэтому никогда не говори, что 'не можешь открыть файлы' — "
    "ты уже их открыл и полностью изучил. "
    "Отвечай вежливо, уверенно, структурированно, с четкими таблицами, выводами и всегда на русском языке."
)

def get_bot_token() -> str:
    """Получение Telegram Bot токена из переменной окружения или bot_token.txt."""
    token = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
    if token:
        return token
    token_file = Path(__file__).resolve().parent / "bot_token.txt"
    if token_file.exists():
        token = token_file.read_text(encoding="utf-8").strip()
        if token:
            return token
    config_file = Path.home() / ".openjarvis" / "config.toml"
    if config_file.exists():
        try:
            import tomllib
            with open(config_file, "rb") as f:
                data = tomllib.load(f)
                token = data.get("channel", {}).get("telegram", {}).get("bot_token", "").strip()
                if token:
                    return token
        except Exception:
            pass
    return ""

# Папка для загрузки присланных файлов
DOWNLOADS_DIR = Path(__file__).resolve().parent / "downloads"
DOWNLOADS_DIR.mkdir(exist_ok=True)

# Клавиатура по умолчанию (Reply Keyboard)
DEFAULT_KEYBOARD = ReplyKeyboardMarkup(
    [
        ["🔄 Новый диалог", "⚙️ Модель"],
        ["📊 Статус", "ℹ️ Помощь"]
    ],
    resize_keyboard=True,
    is_persistent=True
)

# Быстрые инлайн-кнопки под ответами
INLINE_KEYBOARD = InlineKeyboardMarkup(
    [
        [
            InlineKeyboardButton("🔄 Новый диалог", callback_data="cmd_new"),
            InlineKeyboardButton("⚙️ О модели", callback_data="cmd_model"),
        ]
    ]
)

# Хранилище контекста диалогов: chat_id -> list of message dicts
conversation_history: Dict[int, List[Dict[str, str]]] = {}
MAX_HISTORY_TURNS = 10  # максимум 10 пар сообщений для сохранения контекста

# Очередь пакетного приема файлов: chat_id -> list of file dicts
pending_files: Dict[int, List[Dict[str, Any]]] = {}
# Таймеры пакетов: chat_id -> asyncio.Task
pending_batch_tasks: Dict[int, asyncio.Task] = {}


def get_api_key() -> str:
    """Получение API ключа NVIDIA NIM с гарантированной очисткой от BOM."""
    key = os.environ.get("NIM_API_KEY", "").strip()
    if key:
        return key.replace("\ufeff", "").strip()
    
    # Попробуем прочитать из api_key.txt
    api_file = Path(__file__).resolve().parent / "api_key.txt"
    if api_file.exists():
        key = api_file.read_text(encoding="utf-8-sig").strip()
        if key:
            return key.replace("\ufeff", "").strip()
            
    return ""


async def query_nim(messages: List[Dict[str, str]], api_key: str, max_tokens: int = 2000) -> str:
    """Запрос к NVIDIA NIM API."""
    clean_key = api_key.replace("\ufeff", "").strip()
    headers = {
        "Authorization": f"Bearer {clean_key}",
        "Content-Type": "application/json",
        "Accept": "application/json"
    }
    
    full_messages = []
    if not messages or messages[0].get("role") != "system":
        full_messages.append({"role": "system", "content": SYSTEM_PROMPT})
    full_messages.extend(messages)

    payload = {
        "model": MODEL_NAME,
        "messages": full_messages,
        "temperature": 0.5,
        "max_tokens": max_tokens
    }
    async with httpx.AsyncClient(timeout=120.0) as client:
        resp = await client.post(f"{NIM_BASE_URL}/chat/completions", json=payload, headers=headers)
        resp.raise_for_status()
        data = resp.json()
        choices = data.get("choices", [])
        if choices:
            msg = choices[0].get("message", {})
            content = (msg.get("content") or msg.get("reasoning_content") or "").strip()
            if content:
                return content
            return "Готово. Данные обработаны."
        return "Ошибка: пустой ответ от модели."


async def send_split_message(update: Update, text: str, reply_markup=None):
    """Отправка сообщения с разделением длинного текста (лимит Telegram 4096 символов)."""
    max_len = 4000
    chunks = [text[i:i + max_len] for i in range(0, len(text), max_len)]
    
    for i, chunk in enumerate(chunks):
        markup = reply_markup if i == len(chunks) - 1 else None
        try:
            await update.effective_message.reply_text(
                chunk,
                parse_mode=ParseMode.MARKDOWN,
                reply_markup=markup
            )
        except Exception:
            await update.effective_message.reply_text(
                chunk,
                reply_markup=markup
            )


# ---------------- Команды бота ----------------

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик /start."""
    chat_id = update.effective_chat.id
    conversation_history[chat_id] = []
    
    welcome_msg = (
        "👋 *Привет! Я твой персональный ассистент Джарвис.*\n\n"
        "⚡ Работаю на базе **OpenJarvis** и облачных микросервисов **NVIDIA NIM**.\n"
        f"🧠 Активная модель: `{MODEL_NAME}`\n\n"
        "📊 *Что я умею:*\n"
        "• Отвечать на любые вопросы и помогать в диалоге\n"
        "• **Работать с файлами:** присылайте Excel (.xlsx, .xls), CSV, PDF, Word\n"
        "• **Считать экономику маркетплейсов (Wildberries, Ozon):** продажи, комиссии, логистику, рекламу, чистую прибыль\n"
        "• Сводить пакеты отчетов от бухгалтера или менеджера в единую сводку\n\n"
        "Просто отправьте мне файлы или напишите сообщение 👇"
    )
    await update.effective_message.reply_text(
        welcome_msg,
        parse_mode=ParseMode.MARKDOWN,
        reply_markup=DEFAULT_KEYBOARD
    )


async def cmd_new(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик очистки контекста (/new)."""
    chat_id = update.effective_chat.id
    conversation_history[chat_id] = []
    msg = "🔄 *Контекст диалога очищен!* Память обновлена. Чем могу помочь?"
    await update.effective_message.reply_text(msg, parse_mode=ParseMode.MARKDOWN, reply_markup=DEFAULT_KEYBOARD)


async def cmd_model(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Информация о модели (/model)."""
    info = (
        "⚙️ *Информация о нейросети:*\n\n"
        f"• *Модель:* `{MODEL_NAME}`\n"
        "• *Провайдер:* NVIDIA NIM (Inference Microservice)\n"
        "• *Архитектура:* Fast Conversational Nemotron LLM\n"
        "• *Анализ документов:* Поддерживается (Excel, CSV, PDF, Word, TXT)\n"
        "• *Статус:* Активна и готова к работе 🟢"
    )
    await update.effective_message.reply_text(info, parse_mode=ParseMode.MARKDOWN, reply_markup=DEFAULT_KEYBOARD)


async def cmd_status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Проверка статуса системы (/status)."""
    api_key = get_api_key()
    if not api_key:
        status_text = "❌ *Ошибка:* API ключ NVIDIA NIM не задан."
    else:
        status_text = (
            "📊 *Статус системы Джарвис:*\n\n"
            "• *Платформа:* Windows 10 (Local PC)\n"
            "• *Движок:* NVIDIA NIM Cloud API\n"
            "• *Подключение:* Онлайн 🟢\n"
            f"• *Модель:* `{MODEL_NAME}`\n"
            f"• *Память сессии:* {len(conversation_history.get(update.effective_chat.id, []))} сообщений\n"
            "• *Обработчик файлов:* Активен (openpyxl, pandas, pypdf, docx)"
        )
    await update.effective_message.reply_text(status_text, parse_mode=ParseMode.MARKDOWN, reply_markup=DEFAULT_KEYBOARD)


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Справка (/help)."""
    help_text = (
        "ℹ️ *Возможности Джарвиса:*\n\n"
        "📈 *Анализ файлов и отчетов:*\n"
        "Присылайте или пересылайте таблицы (.xlsx, .xls, .csv), отчеты по рекламе, PDF и документы Word. "
        "Джарвис сведет все файлы, рассчитает выручку, расходы, комиссии WB/Ozon и чистую прибыль.\n\n"
        "💬 *Диалог:* Задавайте любые вопросы по бизнесу, финансам, расчетам или коду.\n\n"
        "*Кнопки меню:*\n"
        "• `🔄 Новый диалог` — сбросить память текущей беседы.\n"
        "• `⚙️ Модель` — параметры нейросети.\n"
        "• `📊 Статус` — проверка связи.\n"
        "• `ℹ️ Помощь` — эта подсказка."
    )
    await update.effective_message.reply_text(help_text, parse_mode=ParseMode.MARKDOWN, reply_markup=DEFAULT_KEYBOARD)


async def handle_callback_query(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка кликов по инлайн-кнопкам."""
    query = update.callback_query
    await query.answer()
    
    if query.data == "cmd_new":
        chat_id = update.effective_chat.id
        conversation_history[chat_id] = []
        await query.message.reply_text("🔄 Контекст диалога очищен! Чем могу помочь?", reply_markup=DEFAULT_KEYBOARD)
    elif query.data == "cmd_model":
        await cmd_model(update, context)


# ---------------- Обработка файлов и документов ----------------

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик входящих файлов и документов."""
    chat_id = update.effective_chat.id
    user_name = update.effective_user.first_name if update.effective_user else "Пользователь"
    doc = update.effective_message.document
    caption = (update.effective_message.caption or "").strip()

    if not doc:
        return

    file_name = doc.file_name or f"file_{doc.file_id[:8]}"
    logger.info(f"📁 Получен документ от {user_name} (ID: {chat_id}): {file_name} ({doc.file_size} байт)")

    # Папка пользователя
    user_dir = DOWNLOADS_DIR / str(chat_id)
    user_dir.mkdir(exist_ok=True)
    save_path = user_dir / file_name

    try:
        tg_file = await context.bot.get_file(doc.file_id)
        await tg_file.download_to_drive(custom_path=save_path)
    except Exception as e:
        logger.exception(f"Ошибка скачивания {file_name}")
        await update.effective_message.reply_text(f"❌ Ошибка скачивания файла {file_name}: {e}")
        return

    # Извлекаем данные из файла
    try:
        parsed_content = extract_file_content(save_path)
    except Exception as e:
        logger.exception(f"Ошибка обработки {file_name}")
        parsed_content = f"Ошибка извлечения данных из файла {file_name}: {e}"

    if chat_id not in pending_files:
        pending_files[chat_id] = []

    pending_files[chat_id].append({
        "name": file_name,
        "path": save_path,
        "content": parsed_content,
        "caption": caption
    })

    # Сбрасываем таймер пакета, если файлы пересылаются подряд
    if chat_id in pending_batch_tasks and not pending_batch_tasks[chat_id].done():
        pending_batch_tasks[chat_id].cancel()

    # Запускаем отложенную обработку пакета (пауза 2.0 сек для сбора всех пересланных файлов)
    pending_batch_tasks[chat_id] = asyncio.create_task(
        process_batch_after_delay(chat_id, update, context, delay=2.0)
    )


async def process_batch_after_delay(chat_id: int, update: Update, context: ContextTypes.DEFAULT_TYPE, delay: float = 2.0):
    """Сбор пакета файлов и отправка общего аналитического запроса в ИИ."""
    try:
        await asyncio.sleep(delay)
    except asyncio.CancelledError:
        return

    files = pending_files.pop(chat_id, [])
    if not files:
        return

    api_key = get_api_key()
    if not api_key:
        await update.effective_message.reply_text(
            "❌ Ключ NVIDIA NIM не найден в `api_key.txt`.\nПожалуйста, добавьте ваш токен!",
            reply_markup=DEFAULT_KEYBOARD
        )
        return

    user_name = update.effective_user.first_name if update.effective_user else "Пользователь"

    # Определяем задание пользователя:
    # 1. Из подписей к файлам
    # 2. Если подписи нет — проверяем последнее сообщение в истории диалога
    captions = [f["caption"] for f in files if f.get("caption")]
    user_task = " ".join(captions).strip() if captions else ""

    if not user_task:
        # Проверяем последний вопрос пользователя из контекста
        user_history = conversation_history.get(chat_id, [])
        for m in reversed(user_history):
            if m.get("role") == "user" and not m.get("content", "").startswith("Пользователь прислал"):
                user_task = m.get("content", "").strip()
                break

    if not user_task:
        user_task = (
            "Сделай детальный сводный расчет экономики магазина на маркетплейсе (выручка, продажи в шт, "
            "комиссия маркетплейса, логистика, расходы на рекламу, чистый остаток к перечислению, рентабельность). "
            "Сведи показатели по всем отчетам в понятные таблицы и напиши выводы и рекомендации."
        )

    if len(files) == 1:
        f = files[0]
        prompt = (
            f"Ты открыл и проанализировал файл '{f['name']}'. Вот извлеченные из него данные и расчеты:\n\n"
            f"{f['content']}\n\n"
            f"Задание пользователя: {user_task}\n\n"
            "Инструкция: Проанализируй файл как профессиональный бизнес-аналитик, выдели ключевые финансовые показатели, рассчитай итоги и представь ответ в виде наглядного отчета с таблицами."
        )
        status_msg = f"📊 Анализирую файл *{f['name']}*..."
    else:
        file_blocks = []
        for idx, f in enumerate(files, 1):
            file_blocks.append(f"=== Отчет {idx}: {f['name']} ===\n{f['content']}")
        
        all_blocks = "\n\n".join(file_blocks)
        prompt = (
            f"Ты открыл и полностью прочитал пакет отчетов ({len(files)} файлов):\n\n"
            f"{all_blocks}\n\n"
            f"Задание пользователя: {user_task}\n\n"
            "Инструкция: Сведи данные из всех отчетов воедино как финансовый директор. "
            "Построй итоговую сводную таблицу (Выручка, Продажи в шт, Комиссия маркетплейса, Логистика, Затраты на рекламу, Итоговые удержания, Чистый доход к перечислению). "
            "Рассчитай ключевые финансовые метрики (долю рекламы ДРР, рентабельность) и дай практические рекомендации владельцу магазина."
        )
        status_msg = f"📂 Принято {len(files)} файлов. Свожу данные и рассчитываю экономику магазина..."

    await update.effective_message.reply_text(status_msg, parse_mode=ParseMode.MARKDOWN)
    await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.TYPING)

    # Ведение истории
    if chat_id not in conversation_history:
        conversation_history[chat_id] = []

    history = conversation_history[chat_id]
    history.append({"role": "user", "content": prompt})

    if len(history) > MAX_HISTORY_TURNS * 2:
        history = history[-MAX_HISTORY_TURNS * 2:]
        conversation_history[chat_id] = history

    try:
        reply = await query_nim(history, api_key, max_tokens=2500)
        history.append({"role": "assistant", "content": reply})
        logger.info(f"🤖 Джарвис отправил готовый анализ {len(files)} файлов для {user_name} ({len(reply)} симв.)")
        await send_split_message(update, reply, reply_markup=INLINE_KEYBOARD)
    except httpx.HTTPStatusError as e:
        logger.error(f"HTTP ошибка NIM при анализе файлов: {e.response.status_code} - {e.response.text}")
        await update.effective_message.reply_text(
            f"⚠️ Ошибка ответа от нейросети ({e.response.status_code}). Попробуйте разбить файлы на меньшие части.",
            reply_markup=DEFAULT_KEYBOARD
        )
    except Exception as e:
        logger.exception("Ошибка при обработке файлов")
        await update.effective_message.reply_text(
            f"⚠️ Ошибка при расчете данных: {e}",
            reply_markup=DEFAULT_KEYBOARD
        )


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка изображений."""
    user_name = update.effective_user.first_name if update.effective_user else "Пользователь"
    logger.info(f"📷 Получено фото от {user_name}")
    await update.effective_message.reply_text(
        "📷 Фото получено. Если вы отправляете отчеты или таблицы, пожалуйста, прикрепляйте их как файлы-документы (.xlsx, .csv, .pdf), чтобы Джарвис мог точно извлечь все формулы и цифры!",
        reply_markup=DEFAULT_KEYBOARD
    )


async def handle_non_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Ответ на неподдерживаемые типы сообщений (голосовые, стикеры)."""
    user_name = update.effective_user.first_name if update.effective_user else "Пользователь"
    logger.info(f"⚠️ Неподдерживаемый тип сообщения от {user_name}")
    await update.effective_message.reply_text(
        "🎙️ Голосовые сообщения пока не поддерживаются. Пожалуйста, напишите запрос текстом или пришлите файлы (.xlsx, .csv, .pdf, .docx)!",
        reply_markup=DEFAULT_KEYBOARD
    )


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка обычных текстовых сообщений пользователя."""
    text = (update.effective_message.text or "").strip()
    chat_id = update.effective_chat.id
    user_name = update.effective_user.first_name if update.effective_user else "Пользователь"

    if not text:
        return

    logger.info(f"📨 Сообщение от {user_name} (ID: {chat_id}): {text}")

    # Проверка нажатия постоянных кнопок
    if text == "🔄 Новый диалог":
        return await cmd_new(update, context)
    if text == "⚙️ Модель":
        return await cmd_model(update, context)
    if text == "📊 Статус":
        return await cmd_status(update, context)
    if text == "ℹ️ Помощь":
        return await cmd_help(update, context)

    api_key = get_api_key()
    if not api_key:
        await update.effective_message.reply_text(
            "❌ Ключ NVIDIA NIM не найден в `api_key.txt`.\nПожалуйста, добавьте ваш токен!",
            reply_markup=DEFAULT_KEYBOARD
        )
        return

    await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.TYPING)

    if chat_id not in conversation_history:
        conversation_history[chat_id] = []

    history = conversation_history[chat_id]
    history.append({"role": "user", "content": text})

    if len(history) > MAX_HISTORY_TURNS * 2:
        history = history[-MAX_HISTORY_TURNS * 2:]
        conversation_history[chat_id] = history

    try:
        reply = await query_nim(history, api_key)
        history.append({"role": "assistant", "content": reply})
        logger.info(f"🤖 Ответ Джарвиса для {user_name} ({len(reply)} симв.): {reply[:60]}...")
        await send_split_message(update, reply, reply_markup=INLINE_KEYBOARD)
    except httpx.HTTPStatusError as e:
        logger.error(f"HTTP ошибка NIM: {e.response.status_code} - {e.response.text}")
        await update.effective_message.reply_text(
            f"⚠️ Ошибка ответа от NVIDIA NIM ({e.response.status_code}). Попробуйте позже или сбросьте диалог.",
            reply_markup=DEFAULT_KEYBOARD
        )
    except Exception as e:
        logger.exception("Ошибка при обработке запроса")
        await update.effective_message.reply_text(
            f"⚠️ Произошла ошибка: {e}",
            reply_markup=DEFAULT_KEYBOARD
        )


def main():
    api_key = get_api_key()
    if not api_key:
        print("[!] ВНИМАНИЕ: NIM_API_KEY не задан. Бот может не отвечать на запросы.")
    else:
        print(f"[OK] NIM API ключ загружен: {api_key[:12]}...")

    bot_token = get_bot_token()
    if not bot_token:
        print("[!] ОШИБКА: Telegram токен не найден в bot_token.txt или TELEGRAM_BOT_TOKEN.")
        return

    print(f"[OK] Запуск Telegram бота (токен: {bot_token[:10]}...)...")
    app = ApplicationBuilder().token(bot_token).build()

    # Регистрация команд
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("new", cmd_new))
    app.add_handler(CommandHandler("model", cmd_model))
    app.add_handler(CommandHandler("status", cmd_status))
    app.add_handler(CommandHandler("help", cmd_help))

    # Обработка инлайн-кнопок
    app.add_handler(CallbackQueryHandler(handle_callback_query))

    # Обработка документов и файлов (.xlsx, .csv, .pdf, .docx и др.)
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))

    # Обработка фотографий
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))

    # Обработка текстовых сообщений
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    # Обработка прочих сообщений (голос, стикеры)
    app.add_handler(MessageHandler(~filters.TEXT & ~filters.COMMAND & ~filters.Document.ALL & ~filters.PHOTO, handle_non_text))

    print("[OK] Бот успешно запущен и слушает входящие сообщения Telegram!")
    app.run_polling(drop_pending_updates=False)


if __name__ == "__main__":
    main()
