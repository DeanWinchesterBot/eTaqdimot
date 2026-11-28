import os
import sys
import re
import json
import time
import threading
import urllib.parse
from collections import defaultdict
from datetime import date
from concurrent.futures import ThreadPoolExecutor

import requests
import telebot
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

BOT_TOKEN = (os.environ.get("BOT_TOKEN") or "").strip()
GEMINI_API_KEY = (os.environ.get("GEMINI_API_KEY") or "").strip()
DAILY_LIMIT = int(os.environ.get("DAILY_LIMIT", "3"))      # bitta foydalanuvchiga kuniga
USE_IMAGES = os.environ.get("USE_IMAGES", "1") == "1"
ALLOWED_USERS = set()  # bo'sh = hamma uchun ochiq. Masalan: {123456789}
ADMIN_IDS = set()      # limitdan ozod foydalanuvchilar

NAVY = RGBColor(18, 38, 82)
ACCENT = RGBColor(37, 99, 235)
TEXT = RGBColor(35, 38, 45)
BG = RGBColor(246, 248, 252)

bot = telebot.TeleBot(BOT_TOKEN, threaded=True) if BOT_TOKEN else None
client = (
    genai.Client(api_key=GEMINI_API_KEY, http_options=types.HttpOptions(timeout=90000))
    if GEMINI_API_KEY else None
)

usage = defaultdict(lambda: {"day": None, "n": 0})
usage_lock = threading.Lock()


# ---------------------------------------------------------------- limitlar
def is_allowed(user_id: int) -> bool:
    return not ALLOWED_USERS or user_id in ALLOWED_USERS


def try_consume(user_id: int) -> bool:
    if user_id in ADMIN_IDS:
        return True
    today = date.today().isoformat()
    with usage_lock:
        u = usage[user_id]
        if u["day"] != today:
            u["day"], u["n"] = today, 0
        if u["n"] >= DAILY_LIMIT:
            return False
        u["n"] += 1
        return True


def refund(user_id: int):
    with usage_lock:
        if usage[user_id]["n"] > 0:
            usage[user_id]["n"] -= 1


# ---------------------------------------------------------------- matnni tahlil
COUNT_RE = re.compile(
    r"(\d+)\s*(?:ta\s+)?(?:betlik|slaydlik|sahifalik|bet|slayd|sahifa|ta)\b",
    re.IGNORECASE,
)
FILLER_RE = re.compile(
    r"\b(mavzusida|mavzuda|haqida|kerak|tayyorlab|tayyorla|qilib|ber|menga|iltimos|taqdimot|slayd)\b",
    re.IGNORECASE,
)


def parse_request(text: str):
    m = COUNT_RE.search(text)
    count = int(m.group(1)) if m else 10
    count = max(3, min(count, 25))
    topic = COUNT_RE.sub(" ", text)
    topic = FILLER_RE.sub(" ", topic)
    topic = re.sub(r"\s+", " ", topic).strip(" .,:;-") or text.strip()
    return topic[:200], count


# ---------------------------------------------------------------- Gemini
def clean_slides(raw, count):
    if isinstance(raw, dict):
        raw = raw.get("slides", [])
    slides = []
    for item in raw if isinstance(raw, list) else []:
        if not isinstance(item, dict):
            continue
        title = str(item.get("title", "")).strip()
        points = [str(p).strip() for p in (item.get("points") or []) if str(p).strip()]
        if not title:
            continue
        slides.append({
            "title": title[:120],
            "points": points[:6],
            "image_prompt": str(item.get("image_prompt") or title),
        })
    if len(slides) < 2:
        raise ValueError("Gemini yaroqli slaydlar qaytarmadi")
    return slides[:count]


def generate_slide_data(topic: str, count: int):
    if not client:
        raise RuntimeError("GEMINI_API_KEY sozlanmagan.")

    prompt = f"""
Mavzu: "{topic}".
Shu mavzuda {count} ta slayddan iborat taqdimot tuzib ber.
1-slayd titul slaydi bo'lsin: "title" - taqdimot nomi, "points" - bitta qisqa kichik sarlavha.
Oxirgi slayd xulosa bo'lsin.
Matnlar o'zbek tilida (lotin), lo'nda va mazmunli bo'lsin. Har bir punkt 8-14 so'z, har slaydda 3-5 punkt.
"image_prompt" - inglizcha, mavzuga mos, minimal uslubdagi illustratsiya tavsifi.

Javobni FAQAT JSON massiv ko'rinishida qaytar:
[{{"title": "...", "points": ["...", "..."], "image_prompt": "..."}}]
"""
    last = None
    for attempt in range(3):
        try:
            resp = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt,
                config=types.GenerateContentConfig(response_mime_type="application/json"),
            )
            return clean_slides(json.loads(resp.text), count)
        except Exception as e:
            last = e
            print(f"Gemini urinish {attempt + 1}/3 xato: {e}")
            time.sleep(2 * (attempt + 1))
    raise last


# ---------------------------------------------------------------- rasm
def download_single_image(task):
    idx, prompt, filename = task
    try:
        url = (
            "https://image.pollinations.ai/prompt/"
            f"{urllib.parse.quote(prompt)}?width=800&height=600&nologo=true"
        )
        res = requests.get(url, timeout=25)
        if res.status_code == 200 and len(res.content) > 1000:
            with open(filename, "wb") as f:
                f.write(res.content)
            return idx, filename
    except Exception as e:
        print(f"Rasm {idx} xatosi:", e)
    return idx, None


# ---------------------------------------------------------------- PPTX
def add_rect(slide, x, y, w, h, color):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    return shp


def build_pptx(slides_data: list, output_file: str):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    image_paths, tasks = {}, []
    stamp = int(time.time() * 1000)
    if USE_IMAGES:
        for idx, item in enumerate(slides_data):
            if idx > 0 and idx % 2 == 1:
                tasks.append((idx, item["image_prompt"], f"img_{idx}_{stamp}.jpg"))
    if tasks:
        with ThreadPoolExecutor(max_workers=4) as ex:
            for idx, path in ex.map(download_single_image, tasks):
                if path:
                    image_paths[idx] = path

    try:
        last = len(slides_data) - 1
        for idx, item in enumerate(slides_data):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = BG

            # titul slayd
            if idx == 0:
                add_rect(slide, 0, 0, Inches(0.5), prs.slide_height, ACCENT)
                tb = slide.shapes.add_textbox(Inches(1.2), Inches(2.2), Inches(10.9), Inches(2.4))
                tf = tb.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = item["title"]
                p.font.size = Pt(42 if len(item["title"]) < 60 else 34)
                p.font.bold = True
                p.font.color.rgb = NAVY
                if item["points"]:
                    p2 = tf.add_paragraph()
                    p2.text = item["points"][0]
                    p2.font.size = Pt(22)
                    p2.font.color.rgb = ACCENT
                    p2.space_before = Pt(18)
                continue

            # sarlavha paneli
            add_rect(slide, 0, 0, prs.slide_width, Inches(1.35), NAVY)
            tb = slide.shapes.add_textbox(Inches(0.7), Inches(0.25), Inches(11.9), Inches(0.95))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = item["title"]
            p.font.size = Pt(28 if len(item["title"]) < 50 else 22)
            p.font.bold = True
            p.font.color.rgb = RGBColor(255, 255, 255)

            has_img = idx in image_paths and idx != last
            width = Inches(6.6) if has_img else Inches(11.9)
            points = item["points"]
            chars = sum(len(x) for x in points)
            size = 22 if chars < 250 else 19 if chars < 350 else 16

            bb = slide.shapes.add_textbox(Inches(0.7), Inches(1.8), width, Inches(5.0))
            tf = bb.text_frame
            tf.word_wrap = True
            for i, point in enumerate(points):
                para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                para.text = f"▪ {point}"
                para.font.size = Pt(size)
                para.font.color.rgb = TEXT
                para.space_after = Pt(12)

            if has_img:
                slide.shapes.add_picture(image_paths[idx], Inches(7.6), Inches(1.85), width=Inches(5.0))

            # slayd raqami
            nb = slide.shapes.add_textbox(Inches(12.2), Inches(6.9), Inches(0.9), Inches(0.4))
            np_ = nb.text_frame.paragraphs[0]
            np_.text = str(idx + 1)
            np_.font.size = Pt(12)
            np_.font.color.rgb = ACCENT

        prs.save(output_file)
    finally:
        for path in image_paths.values():
            if os.path.exists(path):
                os.remove(path)


# ---------------------------------------------------------------- Telegram
if bot:
    @bot.message_handler(commands=["start", "help"])
    def cmd_start(message):
        if not is_allowed(message.from_user.id):
            bot.reply_to(message, "Bu bot hozircha faqat test guruhi uchun ishlaydi.")
            return
        bot.reply_to(
            message,
            "Salom! Mavzu va slaydlar sonini yozing:\n\n"
            "Menejment asoslari mavzusida 12 betlik slayd kerak\n\n"
            f"Kuniga {DAILY_LIMIT} tagacha taqdimot yasash mumkin.",
        )

    @bot.message_handler(content_types=["text"])
    def handle_text(message):
        uid = message.from_user.id
        if not is_allowed(uid):
            bot.reply_to(message, "Bu bot hozircha faqat test guruhi uchun ishlaydi.")
            return

        text = (message.text or "").strip()
        if len(text) < 3:
            bot.reply_to(message, "Mavzuni yozing, masalan: Menejment asoslari 10 ta slayd")
            return
        if not try_consume(uid):
            bot.reply_to(message, f"Bugungi limit ({DAILY_LIMIT} ta) tugadi. Ertaga qayta urinib ko'ring.")
            return

        topic, count = parse_request(text)
        wait_msg = bot.reply_to(message, f"⚙️ {topic} mavzusida {count} betlik slayd tayyorlanmoqda...")
        out_file = f"slayd_{message.chat.id}_{int(time.time())}.pptx"

        try:
            data = generate_slide_data(topic, count)
            build_pptx(data, out_file)
            with open(out_file, "rb") as f:
                bot.send_document(
                    message.chat.id, f,
                    caption=f"✅ Tayyor!\n📌 Mavzu: {topic}\n📄 Slaydlar: {len(data)} bet",
                )
            bot.delete_message(message.chat.id, wait_msg.message_id)
        except Exception as e:
            print("Xato:", repr(e))
            refund(uid)  # xato bo'lsa limit qaytariladi
            bot.edit_message_text(
                "❌ Slayd tayyorlab bo'lmadi. Birozdan keyin qayta urinib ko'ring.",
                message.chat.id, wait_msg.message_id,
            )
        finally:
            if os.path.exists(out_file):
                os.remove(out_file)


def run_bot():
    sys.stdout.reconfigure(line_buffering=True)
    if not bot:
        print("XATO: BOT_TOKEN topilmadi!")
        return
    try:
        bot.remove_webhook()
    except Exception as e:
        print("Webhook o'chirishda xato:", e)
    print("Telegram bot polling boshlandi...")
    while True:
        try:
            bot.infinity_polling(skip_pending=True, timeout=20, long_polling_timeout=20)
        except Exception as e:
            print("Polling xatosi:", e)
            time.sleep(5)


if __name__ == "__main__":
    run_bot()
