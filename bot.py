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

SANA, BUYURTMACHI, LOT_NUMER = range(3)

# Раҳбар маълумоти
DIRECTOR_NAME = "Рузиев Э. Б."

# Сиз сўраган 3 та алоҳида ҳужжат маълумотлари
DOCUMENTS = {
    "uskunalar_ariza": {
        "title": "Зарур ускуна ва техникалар мавжудлиги тўғрисида\nАРИЗА",
        "text": "«Dobus Qurilish» МЧЖ шуни маълум қиладики, {SANA} даги {LOT_NUMER}-сонли тендер шартлари ва техник топшириқда кўрсатилган барча ишларни ва етказиб беришни ўз вақтида ҳамда сифатли амалга ошириш учун зарур бўлган барча махсус ускуналар, техникалар ва асбоб-анжомлар жамиятимиз ихтиёрида мавжуд ва соз ҳолатдадир."
    },
    "kafolat_xati": {
        "title": "Ишончлилик, сифат ва муддат кафолати тўғрисида\nКАФОЛАТ ХАТИ",
        "text": "«Dobus Qurilish» МЧЖ {SANA} даги {LOT_NUMER}-сонли тендер доирасида бажариладиган барча ишлар ва етказиб беришлар сифатли, техник талабларга тўлиқ мос равишда ва белгиланган муддатларда бажарилишига кафолат беради."
    },
    "tanishdim_ariza": {
        "title": "Тендер ҳужжатлари ва шартлари билан танишиб чиқилганлиги тўғрисида\nАРИЗА",
        "text": "«Dobus Qurilish» МЧЖ {SANA} даги {LOT_NUMER}-сонли тендер харид ҳужжатлари, техник топшириқлар ҳамда шартнома лойиҳасининг сифат ва бошқа барча талаблари билан атрофлича ва тўлиқ танишиб чиққанини ва ушбу шартларга эътирозсиз розилигини маълум қилади."
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
    await update.message.reply_text("Ассалому алайкум! Ариза ва кафолат хатларини шакллантириш учун **Санани** киритинг (масалан: `30-август 2026-йил`):", parse_mode="Markdown")
    return SANA

async def get_sana(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data['sana'] = update.message.text
    await update.message.reply_text("Раҳмат! Енди **Буюртмачи номини** киритинг (масалан: `«ОКМК» АЖ`):", parse_mode="Markdown")
    return BUYURTMACHI

async def get_buyurtmachi(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data['buyurtmachi'] = update.message.text
    await update.message.reply_text("Енди **Лот рақамини** киритинг (масалан: `55555555`):", parse_mode="Markdown")
    return LOT_NUMER

async def generate_and_send(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    lot_number = update.message.text
    sana = context.user_data['sana']
    buyurtmachi = context.user_data['buyurtmachi']

    await update.message.reply_text("⏳ Ҳужжатлар тайёрланмоқда...")

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
            section.left_margin = Cm(2.0)
            section.right_margin = Cm(1.0)

        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(14)

        # Шапка (Корхона номи)
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_after = Pt(2)
        r_logo = p_logo.add_run("ООО «DOBUS QURILISH»")
        r_logo.font.size = Pt(26)
        r_logo.bold = True

        p_req = doc.add_paragraph()
        p_req.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_req.paragraph_format.space_after = Pt(0)
        r_req = p_req.add_run("г. Ташкент Мирабадский район массив Куйлюк-2, дом-9, кв-27. Тел: 71 290-93-78\np/с 2020 8000 2048 2684 7001, АТИБ «Ипотека банк» Миробад филиал, МФО: 00 420 ОКЭД: 43310, ИНН: 301 458 084")
        r_req.font.size = Pt(10)

        p_line = doc.add_paragraph()
        p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_line.paragraph_format.space_after = Pt(14)
        r_line = p_line.add_run("_________________________________________________________________________________")
        r_line.bold = True
        r_line.font.size = Pt(10)

        # Сана ва Буюртмачи
        table_top = doc.add_table(rows=1, cols=2)
        table_top.alignment = WD_TABLE_ALIGNMENT.CENTER
        remove_table_borders(table_top)

        cell_date = table_top.cell(0, 0)
        cell_date.width = Cm(9.0)
        r_date = cell_date.paragraphs[0].add_run(sana)
        r_date.font.size = Pt(14)
        r_date.bold = True
        r_date.underline = True

        cell_buy = table_top.cell(0, 1)
        cell_buy.width = Cm(9.0)
        p_buy = cell_buy.paragraphs[0]
        p_buy.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_buy = p_buy.add_run(f"«{buyurtmachi}»")
        r_buy.font.size = Pt(14)
        r_buy.bold = True

        doc.add_paragraph().paragraph_format.space_before = Pt(16)

        # Сарлавҳа
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_after = Pt(14)
        r_title = p_title.add_run(doc_info["title"])
        r_title.font.size = Pt(14)
        r_title.bold = True

        # Матни
        content = doc_info["text"].format(SANA=sana, LOT_NUMER=lot_number)
        p_body = doc.add_paragraph()
        p_body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_body.paragraph_format.first_line_indent = Inches(0.5)
        p_body.paragraph_format.space_after = Pt(10)
        r_body = p_body.add_run(content)
        r_body.font.size = Pt(14)

        # Имзо ва Печат
        doc.add_paragraph().paragraph_format.space_before = Pt(30)
        table_sign = doc.add_table(rows=1, cols=3)
        table_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
        remove_table_borders(table_sign)

        cell_left = table_sign.cell(0, 0)
        cell_left.width = Cm(6.5)
        cell_left.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        r_sig1 = cell_left.paragraphs[0].add_run("«Dobus Qurilish» МЧЖ раҳбари")
        r_sig1.font.size = Pt(14)
        r_sig1.bold = True

        cell_middle = table_sign.cell(0, 1)
        cell_middle.width = Cm(6.5)
        cell_middle.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_stamp = cell_middle.paragraphs[0]
        p_stamp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if os.path.exists(clean_stamp):
            p_stamp.add_run().add_picture(clean_stamp, width=Inches(2.5))

        cell_right = table_sign.cell(0, 2)
        cell_right.width = Cm(5.0)
        cell_right.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_sig_name = cell_right.paragraphs[0]
        p_sig_name.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_sig2 = p_sig_name.add_run(DIRECTOR_NAME)
        r_sig2.font.size = Pt(14)
        r_sig2.bold = True

        file_path = os.path.join(temp_dir, f"{key}_{lot_number}.docx")
        doc.save(file_path)
        generated_files.append(file_path)

    for file_p in generated_files:
        try:
            with open(file_p, 'rb') as doc_file:
                await update.message.reply_document(document=doc_file)
            await asyncio.sleep(1.0)
        except Exception as e:
            print(f"Файл юборишда хатолик: {e}")
        finally:
            if os.path.exists(file_p):
                os.remove(file_p)

    if os.path.exists(temp_dir):
        os.rmdir(temp_dir)

    await update.message.reply_text("✅ Барча ҳужжатлар муваффақиятли тайёрланди ва юборилди!")
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Амал бекор қилинди.")
    return ConversationHandler.END

def main():
    # Токенингиз шу ерга ўрнатилди
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
    print("Бот ишга тушди...")
    app.run_polling()

if __name__ == '__main__':
    main()
