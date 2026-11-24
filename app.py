import os
import re
import json
import time
import threading
import urllib.parse
from concurrent.futures import ThreadPoolExecutor
import requests
from flask import Flask
import telebot
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

BOT_TOKEN = os.environ.get("BOT_TOKEN")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")

# Test uchun ruxsat berilgan Telegram ID lar (vergul bilan ajrating)
ALLOWED_USERS = {
    # Masalan: 123456789, 987654321
}

app = Flask(__name__)
bot = telebot.TeleBot(BOT_TOKEN)
client = genai.Client(api_key=GEMINI_API_KEY)

@app.route("/")
def home():
    return "Slayd Bot test rejimida ishlamoqda!"

def is_allowed(user_id: int) -> bool:
    # Agar ro'yxat bo'sh bo'lsa, vaqtincha hamma kirishi mumkin; ID lar qo'shilsa faqat o'shalar
    if not ALLOWED_USERS:
        return True
    return user_id in ALLOWED_USERS

def generate_slide_data(topic: str, count: int):
    prompt = f"""
    Mavzu: "{topic}".
    Shu mavzuda {count} ta slayddan iborat taqdimot tuzib ber.
    1-slayd kirish (titul) slaydi bo'lsin, oxirgi slayd xulosa bo'lsin.
    Matnlar o'zbek tilida, lo'nda, aniq va mazmunli bo'lsin. Har bir punktni qisqa yoz (8-14 so'z).

    Javobni FAQAT quyidagi JSON formatida qaytar:
    [
      {{
        "title": "Slayd sarlavhasi",
        "points": [
          "1-muhim fikr",
          "2-muhim fikr",
          "3-muhim fikr",
          "4-muhim fikr"
        ],
        "image_prompt": "clean modern presentation illustration for {topic}, minimal style, high quality"
      }}
    ]
    """
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
        ),
    )
    return json.loads(response.text)

def download_single_image(task):
    idx, prompt, filename = task
    try:
        encoded_prompt = urllib.parse.quote(prompt)
        url = f"https://image.pollinations.ai/prompt/{encoded_prompt}?width=800&height=600&nologo=true"
        res = requests.get(url, timeout=25)
        if res.status_code == 200 and len(res.content) > 1000:
            with open(filename, "wb") as f:
                f.write(res.content)
            return idx, filename
    except Exception as e:
        print(f"Rasm {idx} xatosi:", e)
    return idx, None

def build_pptx(slides_data: list, output_file: str):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Faqat juft/tanlangan slaydlarga (masalan 1-dan keyin har 2-slaydda) rasm tayyorlaymiz — tez va toza chiqadi
    tasks = []
    image_paths = {}
    stamp = int(time.time())

    for idx, item in enumerate(slides_data):
        if idx > 0 and idx % 2 == 1:
            fname = f"img_{idx}_{stamp}.jpg"
            tasks.append((idx, item.get("image_prompt", item["title"]), fname))

    if tasks:
        with ThreadPoolExecutor(max_workers=4) as executor:
            for idx, saved_path in executor.map(download_single_image, tasks):
                if saved_path:
                    image_paths[idx] = saved_path

    try:
        for idx, item in enumerate(slides_data):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            bg = slide.background.fill
            bg.solid()
            bg.fore_color.rgb = RGBColor(246, 248, 252)

            if idx == 0:
                t_box = slide.shapes.add_textbox(Inches(1.2), Inches(2.3), Inches(10.9), Inches(2.2))
                tf = t_box.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = item["title"]
                p.font.size = Pt(42)
                p.font.bold = True
                p.font.color.rgb = RGBColor(18, 38, 82)
                continue

            has_img = idx in image_paths
            text_width = Inches(6.6) if has_img else Inches(11.6)

            t_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.45), Inches(11.9), Inches(0.95))
            tf_t = t_box.text_frame
            tf_t.word_wrap = True
            pt = tf_t.paragraphs[0]
            pt.text = f"{idx}. {item['title']}"
            pt.font.size = Pt(27)
            pt.font.bold = True
            pt.font.color.rgb = RGBColor(18, 38, 82)

            b_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.6), text_width, Inches(5.3))
            tf_b = b_box.text_frame
            tf_b.word_wrap = True
            for point in item.get("points", []):
                pb = tf_b.add_paragraph()
                pb.text = f"▪ {point}"
                pb.font.size = Pt(19)
                pb.font.color.rgb = RGBColor(35, 38, 45)
                pb.space_after = Pt(12)

            if has_img:
                slide.shapes.add_picture(image_paths[idx], Inches(7.6), Inches(1.65), width=Inches(5.0))

        prs.save(output_file)
    finally:
        for path in image_paths.values():
            if os.path.exists(path):
                os.remove(path)

@bot.message_handler(commands=["start"])
def cmd_start(message):
    if not is_allowed(message.from_user.id):
        bot.reply_to(message, "Bu bot hozircha faqat test guruhi uchun ishlaydi.")
        return

    bot.reply_to(
        message,
        f"Salom! Sizning ID: `{message.from_user.id}`\n\n"
        "Mavzu va slaydlar sonini yozing:\n"
        "*Menejment asoslari mavzusida 12 betlik slayd kerak*",
        parse_mode="Markdown"
    )

@bot.message_handler(func=lambda m: True)
def handle_text(message):
    if not is_allowed(message.from_user.id):
        bot.reply_to(message, "Bu bot hozircha faqat test guruhi uchun ishlaydi.")
        return

    text = message.text.strip()
    nums = re.findall(r"\d+", text)
    slide_count = max(3, min(int(nums[0]) if nums else 10, 25))
    topic = re.sub(r"\d+\s*(betlik|ta|slayd|sahifa).*", "", text, flags=re.IGNORECASE).strip() or text

    wait_msg = bot.reply_to(
        message,
        f"⚙️ *{topic}* mavzusida {slide_count} betlik slayd tayyorlanmoqda...",
        parse_mode="Markdown"
    )
    out_file = f"slayd_{message.chat.id}_{int(time.time())}.pptx"

    try:
        data = generate_slide_data(topic, slide_count)
        build_pptx(data, out_file)
        with open(out_file, "rb") as f:
            bot.send_document(
                message.chat.id,
                f,
                caption=f"✅ Tayyor!\n📌 Mavzu: {topic}\n📄 Slaydlar: {len(data)} bet"
            )
        bot.delete_message(message.chat.id, wait_msg.message_id)
    except Exception as e:
        bot.edit_message_text(f"❌ Xatolik yuz berdi: {e}", message.chat.id, wait_msg.message_id)
    finally:
        if os.path.exists(out_file):
            os.remove(out_file)

def run_bot():
    while True:
        try:
            bot.infinity_polling(skip_pending=True, timeout=20)
        except Exception as e:
            print("Polling xatosi:", e)
            time.sleep(5)

threading.Thread(target=run_bot, daemon=True).start()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
