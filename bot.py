import os
import asyncio
import logging
from PIL import Image
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

SANA, BUYURTMACHI, LOT_NUMER = 0, 1, 2

COMPANY_NAME = "«Dobus Qurilish» МЧЖ"
DIRECTOR_NAME = "Рузиев Э. Б."

# ЖАМИ 13 ТА АЛОҲИДА ҲУЖЖАТ ЛУҒАТИ
DOCUMENTS = {
    # 1. Алоҳида Кафолат хати
    "Kafolat_xati": {
        "title": "Кафолат хати",
        "intro": "{COMPANY_NAME} шуни маълум қиладики, {SANA}даги {LOT_NUMER}-сонли тендерда энг яхши таклифларни танлашда иштирок этиб, товар етказиб бериш учун ёки иш ва хизматларни бажариш учун ишларни бажариш ва хизматларни кўрсатиш бўйича техник таклиф. Тендерда ғолиб деб топилган тақдирда, шартномада кўрсатиладиган ишларни бажариш ва хизматларни кўрсатиш муддатлари юзасидан қуйидаги мажбуриятларни ўз зиммасига олади ва тўлиқ кафолатлайди:",
        "items": [
            "Ишларни бошлаш муддати: Шартнома иккала томонлама имзоланиб, қонуний кучга киргандан ҳамда Давлат буюртмачиси томонидан олдиндан тўлов маблағлари келиб тушган кундан бошлаб 3 (уч) банк кунидан кечиктирмасдан объектда ишларни бошлашни;",
            "Шартнома шартларига амал қилиш: Ишлар ва хизматларни техник топшириқ ҳамда Нуқсон далолатномасига мувофиқ, босқичма-босқич ва белгиланган муддатларда сифатли якунлашни;",
            "Техник база ва сафарбарлик: Ишларни муддатида тугатиш учун муҳандислик-техник ходимларни ва асбоб-ускуналарни объектга биринчи кунданоқ сафарбар этишни;"
        ],
        "outro": "Ушбу кафолат хати шартнома имзоланган кундан бошлаб, барча ишлар тўлиқ якунланиб, Буюртмачи томонидан якуний қабул қилиш далолатномаси имзолангунга қадар ўз кучини сақлаб қолади."
    },
    # 2. Алоҳида Танишув хати
    "Tanishuv_xati": {
        "title": "Танишув хати",
        "intro": "{COMPANY_NAME} шуни маълум қиладики, {SANA}даги {LOT_NUMER}-сонли тендер бўйича ўтказилаётган танлов ҳужжатлари, лойиҳа-смета ҳужжатлари ҳамда тасдиқланган Техник вазифа (ТЗ) шартлари ва талаблари билан тўлиқ танишиб чиққанлигини маълум қилади.",
        "items": [
            "Техник талабларга мувофиқлик: Объектдаги барча ишларни Техник вазифада кўрсатилган параметрлар ва норматив стандартларга риоя қилган ҳолда амалга оширишни;",
            "Сифат кафолати: Ишларни бажаришда фойдаланиладиган барча материалларнинг сифатли ва сертификатланган бўлишини кафолатлаймиз."
        ],
        "outro": "Мазкур танловнинг Техник вазифасида кўрсатилган талаблар юзасидан ҳеч қандай эътирозлар мавжуд эмас."
    },
    # 3. Алоҳида Ускуналар хати
    "Usqunalar_xati": {
        "title": "Асбоб-ускуналар ва жиҳозлар тўғрисида маълумотнома",
        "intro": "{COMPANY_NAME} шуни маълум қиладики, {SANA}даги {LOT_NUMER}-сонли тендер бўйича ишларни бажариш учун қуйидаги мавжуд асбоб-ускуналар ва жиҳозлар ажратилган:",
        "items": [
            "Компрессор воздушный Total 24L / 50L - 2 дона",
            "Болгарка Crown (180мм / 125мм) - 4 дона",
            "Перфоратор Ingco / Crown - 2 дона",
            "Аккумуляторная дрель-шуруповерт - 3 дона",
            "Сварочный аппарат Ураган MIG-350 / АСПТ-2000А - 4 дона"
        ],
        "outro": "Барча асбоб-ускуналар соз ва фойдаланишга тўлиқ тайёр ҳолатда."
    },
    # 4. 1k хати
    "1k_kafolat_xati": {
        "title": "1k Кафолат хати",
        "intro": "{COMPANY_NAME} шуни маълум қиладики, {SANA}даги {LOT_NUMER}-сонли тендерда энг яхши таклифларни танлашда иштирок этиб, ўз мажбуриятларини тўлиқ бажаришни кафолатлайди.",
        "items": [
            "Ишларни бошлаш ва якунлаш муддатларига қатъий риоя қилиш;",
            "Сифатли ва нормативларга мос равишда хизмат кўрсатиш."
        ],
        "outro": ""
    },
    # 5. 2k хати
    "2k_tanishuv": {
        "title": "2k Танишув маълумотномаси",
        "intro": "{COMPANY_NAME} {SANA}даги {LOT_NUMER}-сонли тендернинг барча шартлари, Техник топшириғи ва талаблари билан тўлиқ танишиб чиққанлигини тасдиқлайди.",
        "items": [
            "Тендер ҳужжатларида кўрсатилган сифат ва хавфсизлик нормалари қабул қилинди."
        ],
        "outro": ""
    },
    # 6. 4k хати
    "4k_usqunalar": {
        "title": "4k Ускуналар рўйхати",
        "intro": "{COMPANY_NAME} {SANA}даги {LOT_NUMER}-сонли тендер лойиҳасини амалга ошириш учун тегишли барча маданий-техник ва моддий-техник ускуналарга эга.",
        "items": [
            "Зарур қурилиш ва монтаж ускуналари тўлиқ сафарбар этилади."
        ],
        "outro": ""
    },
    # 7. 5k хати
    "5k_xodimlar": {
        "title": "5k Мутахассислар ва ходимлар ҳақида маълумотнома",
        "intro": "{COMPANY_NAME} {SANA}даги {LOT_NUMER}-сонли тендер бўйича малакали муҳандис-техник ва ишчи ходимлар таркибини тақдим этади:",
        "items": [
            "Бош муҳандис - 1 нафар;",
            "Прораб (Иш юритувчи) - 2 нафар;",
            "Мутахассис ва ишчилар - 16 нафар."
        ],
        "outro": ""
    },
    # 8. 6k хати
    "6k_ish_tajribasi": {
        "title": "6k Иш тажрибаси ҳақида маълумотнома",
        "intro": "{COMPANY_NAME} томонидан {SANA} ва ўтган даврлар мобайнида бажарилган аналог объектлар ва тажриба тўғрисида маълумот.",
        "items": [
            "Бино ва иншоотларни мукаммал таъмирлаш ва реконструкция қилиш ишлари муваффақиятли бажарилган."
        ],
        "outro": ""
    },
    # 9. 7k хати
    "7k_moliyaviy": {
        "title": "7k Молиявий барқарорлик ҳақида маълумотнома",
        "intro": "{COMPANY_NAME} {SANA}даги {LOT_NUMER}-сонли тендер шартларига мувофиқ молиявий жиҳатдан барқарор ва тўлов муддати ўтган қарздорликларга эга эмас.",
        "items": [
            "Банк ва солиқ ташкилотлари олдида қарздорлик мавжуд эмас."
        ],
        "outro": ""
    },
    # 10. 8k хати
    "8k_kafolat_sifat": {
        "title": "8k Сифат ва мувофиқлик кафолати",
        "intro": "{COMPANY_NAME} {SANA}даги {LOT_NUMER}-сонли лот бўйича фойдаланиладиган барча қурилиш материалларининг давлат стандартларига мослигини кафолатлайди.",
        "items": [
            "Бажарилган ишларга кафолат муддати 12 ойни ташкил этади."
        ],
        "outro": ""
    },
    # 11. 9k хати
    "9k_mexnat_muxofazasi": {
        "title": "9k Меҳнат муҳофазаси ва техника хавфсизлиги кафолати",
        "intro": "{COMPANY_NAME} {SANA}даги {LOT_NUMER}-сонли лотда меҳнат муҳофазаси ва техника хавфсизлиги қоидаларига қатъий риоя қилинишини таъминлайди.",
        "items": [
            "Ишчилар махсус кийим ва ҳимоя воситалари билан тўлиқ таъминланган."
        ],
        "outro": ""
    },
    # 12. 10k хати
    "10k_subporyad": {
        "title": "10k Субподрядчиларни жалб этмаслик кафолати",
        "intro": "{COMPANY_NAME} {SANA}даги {LOT_NUMER}-сонли тендер бўйича барча ишларни учинчи шахсларни жалб этмаган ҳолда, ўз кучи билан бажаришни кафолатлайди.",
        "items": [],
        "outro": ""
    },
    # 13. 12k хати
    "12k_tijorat_taklifi": {
        "title": "12k Тижорат таклифи кафолати",
        "intro": "{COMPANY_NAME} {SANA}даги {LOT_NUMER}-сонли лот бўйича тақдим этилаётган нарх таклифи асосли ва якуний эканлигини кафолатлайди.",
        "items": [],
        "outro": ""
    }
}

def make_stamp_transparent(input_path, output_path):
    try:
        img = Image.open(input_path).convert("RGBA")
        datas = img.getdata()
        new_data = []
        for item in datas:
            if item[0] > 200 and item[1] > 200 and item[2] > 200:
                new_data.append((255, 255, 255, 0))
            else:
                new_data.append(item)
        img.putdata(new_data)
        img.save(output_path, "PNG")
    except Exception as e:
        print(f"Печат расмида хатолик: {e}")

def remove_table_borders(table):
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        tblBorders.append(border)
    tblPr.append(tblBorders)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text("1️⃣ **Санани** киритинг (масалан: `2026-йил «14» август`):", parse_mode="Markdown")
    return SANA

async def get_sana(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data['sana'] = update.message.text
    await update.message.reply_text("2️⃣ **Буюртмачи номини** киритинг (масалан: `ОКМК`):", parse_mode="Markdown")
    return BUYURTMACHI

async def get_buyurtmachi(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data['buyurtmachi'] = update.message.text
    await update.message.reply_text("3️⃣ **Лот рақамини** киритинг (масалан: `8213557`):", parse_mode="Markdown")
    return LOT_NUMER

async def generate_and_send(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    lot_number = update.message.text
    sana = context.user_data.get('sana', '')
    buyurtmachi = context.user_data.get('buyurtmachi', '')

    await update.message.reply_text("⏳ Барча 13 та алоҳида ҳужжат 14-шрифтда ва 1 саҳифали форматда тайёрланмоқда...")

    base_dir = os.path.dirname(os.path.abspath(__file__))
    original_stamp = os.path.join(base_dir, "stamp.png")
    clean_stamp = os.path.join(base_dir, "stamp_transparent.png")

    if os.path.exists(original_stamp):
        make_stamp_transparent(original_stamp, clean_stamp)

    temp_dir = os.path.join(base_dir, f"temp_{lot_number}")
    os.makedirs(temp_dir, exist_ok=True)

    generated_files = []

    for key, doc_info in DOCUMENTS.items():
        doc = Document()
        
        for section in doc.sections:
            section.top_margin = Cm(1.0)
            section.bottom_margin = Cm(1.0)
            section.left_margin = Cm(1.8)
            section.right_margin = Cm(1.0)

        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(14)

        # 1. Корхона шапкаси
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_after = Pt(1)
        p_logo.paragraph_format.keep_with_next = True
        r_logo = p_logo.add_run("ООО «DOBUS QURILISH»")
        r_logo.font.size = Pt(20)
        r_logo.bold = True

        p_req = doc.add_paragraph()
        p_req.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_req.paragraph_format.space_after = Pt(0)
        p_req.paragraph_format.keep_with_next = True
        r_req = p_req.add_run("г. Ташкент Мирабадский район массив Куйлюк-2, дом-9, кв-27. Тел: 71 290-93-78\np/с 2020 8000 2048 2684 7001, АТИБ «Ипотека банк» Миробад филиал, МФО: 00 420 ОКЭД: 43310, ИНН: 301 458 084")
        r_req.font.size = Pt(8.5)

        p_line = doc.add_paragraph()
        p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_line.paragraph_format.space_after = Pt(6)
        p_line.paragraph_format.keep_with_next = True
        r_line = p_line.add_run("_________________________________________________________________________________")
        r_line.bold = True
        r_line.font.size = Pt(8)

        # 2. Сана ва Буюртмачи
        table_top = doc.add_table(rows=1, cols=2)
        table_top.alignment = WD_TABLE_ALIGNMENT.CENTER
        remove_table_borders(table_top)

        cell_date = table_top.cell(0, 0)
        cell_date.width = Cm(9.0)
        p_d = cell_date.paragraphs[0]
        p_d.paragraph_format.keep_with_next = True
        r_date = p_d.add_run(sana)
        r_date.font.size = Pt(14)
        r_date.bold = True
        r_date.underline = True

        cell_buy = table_top.cell(0, 1)
        cell_buy.width = Cm(9.0)
        p_buy = cell_buy.paragraphs[0]
        p_buy.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_buy.paragraph_format.keep_with_next = True
        r_buy = p_buy.add_run(f"«{buyurtmachi}»")
        r_buy.font.size = Pt(14)
        r_buy.bold = True

        # Сарлавҳа (14pt)
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_before = Pt(8)
        p_title.paragraph_format.space_after = Pt(8)
        p_title.paragraph_format.keep_with_next = True
        r_title = p_title.add_run(doc_info["title"])
        r_title.font.size = Pt(14)
        r_title.bold = True

        # Кириш матни
        intro_text = doc_info["intro"].format(COMPANY_NAME=COMPANY_NAME, SANA=sana, LOT_NUMER=lot_number)
        p_intro = doc.add_paragraph()
        p_intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_intro.paragraph_format.first_line_indent = Inches(0.4)
        p_intro.paragraph_format.space_after = Pt(4)
        p_intro.paragraph_format.keep_with_next = True
        r_intro = p_intro.add_run(intro_text)
        r_intro.font.size = Pt(14)

        # Бандлар
        for idx, item in enumerate(doc_info["items"], 1):
            p_item = doc.add_paragraph()
            p_item.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_item.paragraph_format.left_indent = Inches(0.2)
            p_item.paragraph_format.space_after = Pt(3)
            p_item.paragraph_format.keep_with_next = True
            r_item = p_item.add_run(f"{idx}. {item}")
            r_item.font.size = Pt(14)

        # Якуний матн
        if doc_info["outro"]:
            p_outro = doc.add_paragraph()
            p_outro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_outro.paragraph_format.first_line_indent = Inches(0.4)
            p_outro.paragraph_format.space_before = Pt(4)
            p_outro.paragraph_format.space_after = Pt(6)
            p_outro.paragraph_format.keep_with_next = True
            r_outro = p_outro.add_run(doc_info["outro"])
            r_outro.font.size = Pt(14)

        # 3. Мўҳр ва Имзо
        p_space = doc.add_paragraph()
        p_space.paragraph_format.space_before = Pt(8)
        p_space.paragraph_format.keep_with_next = True

        table_sign = doc.add_table(rows=1, cols=3)
        table_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
        remove_table_borders(table_sign)

        cell_left = table_sign.cell(0, 0)
        cell_left.width = Cm(6.0)
        cell_left.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_l = cell_left.paragraphs[0]
        p_l.paragraph_format.keep_with_next = True
        r_sig1 = p_l.add_run(f"{COMPANY_NAME} раҳбари")
        r_sig1.font.size = Pt(14)
        r_sig1.bold = True

        cell_middle = table_sign.cell(0, 1)
        cell_middle.width = Cm(6.0)
        cell_middle.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_stamp = cell_middle.paragraphs[0]
        p_stamp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_stamp.paragraph_format.keep_with_next = True
        if os.path.exists(clean_stamp):
            p_stamp.add_run().add_picture(clean_stamp, width=Inches(1.8))

        cell_right = table_sign.cell(0, 2)
        cell_right.width = Cm(5.5)
        cell_right.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_sig_name = cell_right.paragraphs[0]
        p_sig_name.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_sig_name.paragraph_format.keep_with_next = True
        r_sig2 = p_sig_name.add_run(f"________________ {DIRECTOR_NAME}")
        r_sig2.font.size = Pt(14)
        r_sig2.bold = True

        file_path = os.path.join(temp_dir, f"{key}_{lot_number}.docx")
        doc.save(file_path)
        generated_files.append(file_path)

    # Жами 13 та файлни юбориш
    for file_p in generated_files:
        try:
            with open(file_p, 'rb') as doc_file:
                await update.message.reply_document(document=doc_file)
            await asyncio.sleep(0.4)
        except Exception as e:
            print(f"Файл юборишда хатолик: {e}")
        finally:
            if os.path.exists(file_p):
                os.remove(file_p)

    if os.path.exists(temp_dir):
        os.rmdir(temp_dir)

    await update.message.reply_text("✅ Жами 13 та алоҳида хат (Кафолат хати, Танишув хати, Ускуналар хати ва 1k, 2k, 4k, 5k, 6k, 7k, 8k, 9k, 10k, 12k) 14-шрифтда тайёрланди ва юборилди!")
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Амал бекор қилинди.")
    return ConversationHandler.END

def main():
    TOKEN = "8996916069:AAFfxGbWY6YrK4f784ChJneTAg7tyuLoqW4"

    app = ApplicationBuilder().token(TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler('start', start)],
        states={
            SANA: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_sana)],
            BUYURTMACHI: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_buyurtmachi)],
            LOT_NUMER: [MessageHandler(filters.TEXT & ~filters.COMMAND, generate_and_send)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    )

    app.add_handler(conv_handler)
    print("Бот 13 та ҳужжат режимида ишга тушди...")
    app.run_polling()

if __name__ == '__main__':
    main()
