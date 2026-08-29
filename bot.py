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

DIRECTOR_NAME = "Рузиев Э. Б."

FORMS_DATA = {
    "1k": ("1к-шакли", "Суд қарорлари бўйича бажарилмаган мажбуриятлар мавжуд эмаслиги тўғрисида кафолат хати\n(Ариза).", 
           "«Dobus Qurilish» МЧЖ шуни маълум қиладики, {SANA} даги {LOT_NUMER}-сонли тендерда энг яхши таклифларни танлаб олиш бўйича таклиф тақдим этилган пайтда «Dobus Qurilish» МЧЖнинг суд қарорлари бўйича бажарилмаган мажбуриятлари бўлмаган.\n«Dobus Qurilish» МЧЖ, шунингдек, суд қарори бўйича мажбурият юзага келган тақдирда, бу ҳақда Буюртмачига дарҳол ёзма равишда хабар бериш мажбуриятини ўз зиммасига олади."),
    "2k": ("2к-шакли", "Ўзига нисбатан жорий этилган тўловга қобилиятсизлик (банкротлик, ликвидация) тартиб-таомилларининг мавжуд эмаслиги тўғрисида кафолат хати\n(Ариза).", 
           "«Dobus Qurilish» МЧЖ шуни маълум қиладики, {SANA} даги {LOT_NUMER}-сонли тендерда энг яхши таклифларни танлаб олиш бўйича таклиф тақдим этилган пайтда тўлов қобилиятсизлиги (банкротлик, тугатиш) жараёнлари жорий этилмаган."),
    "3k": ("3к-шакли", "Инсофсиз ижрочилар рўйхатига киритилмаганлиги тўғрисида кафолат хати\n(Ариза).", 
           "«Dobus Qurilish» МЧЖ шуни маълум қиладики, {SANA} даги {LOT_NUMER}-сонли тендерда энг яхши таклифларни танлаб олиш бўйича таклиф тақдим этилган пайтда «Dobus Qurilish» МЧЖ инсофсиз ижрочилар рўйхатига киритилмаган.\n«Dobus Qurilish» МЧЖ, шунингдек, инсофсиз рўйхатига киритилган тақдирда, бу ҳақда Буюртмачига дарҳол ёзма равишда хабар бериш мажбуриятини ўз зиммасига олади."),
    "4k": ("4к-шакли", "Манфаатлар тўқнашуви мавжуд эмаслиги тўғрисида кафолат хати\n(Ариза).", 
           "«Dobus Qurilish» МЧЖ {SANA} даги {LOT_NUMER}-сонли тендер бўйича харид ҳужжатларини ўрганиб чиқиб, манфаатлар тўқнашувининг барча шакллари мавжуд эмаслигини маълум қилади."),
    "5k": ("5к-шакли", "Аффиланганлик мавжуд эмаслиги тўғрисида кафолат хати\n(Ариза).", 
           "«Dobus Qurilish» МЧЖ {SANA} даги {LOT_NUMER}-сонли тендер бўйича аффиланганликнинг барча шакллари мавжуд эмаслигини маълум қилади."),
    "6k": ("6к-шакли", "Коррупция кўринишларига йўл қўймаслик бўйича кафолат хати\n(Ариза).", 
           "«Dobus Qurilish» МЧЖ {SANA} даги {LOT_NUMER}-сонли тендерда иштирок этиб, коррупцияга оид ҳуқуқбузарликларга йўл қўймаслик мажбуриятини олади."),
    "7k": ("7к-шакли", "Техник, молиявий, моддий ва инсон ресурслари мавжудлиги тўғрисида кафолат хати\n(Ариза).", 
           "«Dobus Qurilish» МЧЖ шуни маълум қиладики, {SANA} даги {LOT_NUMER}-сонли тендер бўйича шартномани бажариш учун етарли миқдордаги техник, молиявий, моддий ва инсон ресурсларига эга."),
    "8k": ("8к-шакли", "Етказиб бериш тажрибаси ҳақида маълумот\n(Ариза).", 
           "«Dobus Qurilish» МЧЖ {SANA} даги {LOT_NUMER}-сонли тендер бўйича ўхшаш товарларни етказиб бериш тажрибасига эгалигини маълум қилади."),
    "9k": ("9к-шакли", "Суд томонидан кўриб чиқилаётган ишлар мавжуд эмаслиги ҳақида кафолат хати\n(Ариза).", 
           "«Dobus Qurilish» МЧЖ ва Буюртмачи ўртасида суд томонидан кўриб чиқилаётган низоли ишлар мавжуд эмас."),
    "10k": ("10к-шакли", "Солиқлар ва йиғимлар бўйича қарздорлик мавжуд эмаслиги тўғрисида кафолат хати\n(Ариза).", 
           "«Dobus Qurilish» МЧЖ шуни маълум қиладики, {SANA} даги {LOT_NUMER}-сонли тендерда таклиф тақдим этилган пайтда бюджет ва давлат мақсадли жамғармаларига нисбатан солиқлар ва йиғимлар бўйича муддати ўтган қарздорликлари мавжуд эмас."),
    "11k": ("11к-шакли", "Лицензия ва рухсатномалар мавжудлиги тўғрисида кафолат хати\n(Ариза).", 
           "«Dobus Qurilish» МЧЖ {SANA} даги {LOT_NUMER}-сонли тендер бўйича амалга ошириладиган фаолият турлари учун барча зарур лицензия ва рухсатномаларга эга эканлигини тасдиқлайди."),
    "12k": ("12к-шакли", "Иштирокчининг молиявий аҳволи тўғрисида маълумот\n(Ариза).", 
           "«Dobus Qurilish» МЧЖнинг молиявий кўрсаткичлари ва ликвидлик даражаси барқарор ҳамда тендер талабларига мос келишини маълум қиламиз."),
    "13k": ("13к-шакли", "Таклифнинг амал қилиш муддати тўғрисида кафолат хати\n(Ариза).", 
           "«Dobus Qurilish» МЧЖ {SANA} даги {LOT_NUMER}-сонли тендер бўйича тақдим этилган нарх таклифи ва бошқа шартлар эълон қилинган кундан бошлаб камида 60 календарь кун мобайнида ўз кучида қолишини кафолатлайди."),
    "14k": ("14к-шакли", "Ишончлилик ва сифат кафолати тўғрисида ариза\n(Кафолат хати).", 
           "«Dobus Qurilish» МЧЖ {SANA} даги {LOT_NUMER}-сонли тендер доирасида бажариладиган ишларининг сифати ва белгиланган муддатларда тўлиқ бажарилишини таъминлашга кафолат беради.")
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
        print(f"Rasmda xatolik: {e}")

def remove_table_borders(table):
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        tblBorders.append(border)
    tblPr.append(tblBorders)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Assalomu alaykum! Tender hujjatlarini shakllantirish uchun **Sanani** kiriting (masalan: `26-август 2026-йил`):", parse_mode="Markdown")
    return SANA

async def get_sana(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data['sana'] = update.message.text
    await update.message.reply_text("Rahmat! Endi **Buyurtmachi nomini** kiriting (masalan: `«ОКМК» АЖ`):", parse_mode="Markdown")
    return BUYURTMACHI

async def get_buyurtmachi(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data['buyurtmachi'] = update.message.text
    await update.message.reply_text("Endi **Lot raqamini** kiriting (masalan: `666554488`):", parse_mode="Markdown")
    return LOT_NUMER

async def generate_and_send(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    lot_number = update.message.text
    sana = context.user_data['sana']
    buyurtmachi = context.user_data['buyurtmachi']

    await update.message.reply_text("⏳ Barcha 14 ta hujjat shakllantirilmoqda, iltimos kuting...")

    base_dir = os.path.dirname(os.path.abspath(__file__))
    original_stamp = os.path.join(base_dir, "stamp.png")
    clean_stamp = os.path.join(base_dir, "stamp_transparent.png")

    if os.path.exists(original_stamp):
        make_stamp_transparent(original_stamp, clean_stamp)

    temp_dir = os.path.join(base_dir, f"temp_{lot_number}")
    os.makedirs(temp_dir, exist_ok=True)

    generated_files = []

    for key, (shakl_num, title, text_template) in FORMS_DATA.items():
        doc = Document()
        for section in doc.sections:
            section.top_margin = Cm(1.0)
            section.bottom_margin = Cm(1.0)
            section.left_margin = Cm(2.0)
            section.right_margin = Cm(1.0)

        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(14)

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
        p_line.paragraph_format.space_after = Pt(18)
        r_line = p_line.add_run("_________________________________________________________________________________")
        r_line.bold = True
        r_line.font.size = Pt(10)

        p_shakl = doc.add_paragraph()
        p_shakl.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_shakl.paragraph_format.space_after = Pt(12)
        r_shakl = p_shakl.add_run(f"Харид ҳужжатларига илова\n«{shakl_num}»")
        r_shakl.font.size = Pt(10)

        table_top = doc.add_table(rows=1, cols=2)
        table_top.alignment = WD_TABLE_ALIGNMENT.CENTER
        remove_table_borders(table_top)
        
        cell_date = table_top.cell(0, 0)
        cell_date.width = Cm(9.0)
        r_date = cell_date.paragraphs[0].add_run(f"{sana}")
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

        p_spacer = doc.add_paragraph()
        p_spacer.paragraph_format.space_before = Pt(12)

        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_after = Pt(14)
        r_title = p_title.add_run(title)
        r_title.font.size = Pt(14)
        r_title.bold = True

        content = text_template.format(SANA=sana, LOT_NUMER=lot_number)
        for para in content.split('\n'):
            p_body = doc.add_paragraph()
            p_body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_body.paragraph_format.first_line_indent = Inches(0.5)
            p_body.paragraph_format.space_after = Pt(10)
            r_body = p_body.add_run(para)
            r_body.font.size = Pt(14)

        doc.add_paragraph().paragraph_format.space_before = Pt(20)
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

        file_path = os.path.join(temp_dir, f"{key}_shakli_{lot_number}.docx")
        doc.save(file_path)
        generated_files.append(file_path)

    for file_p in generated_files:
        try:
            with open(file_p, 'rb') as doc_file:
                await update.message.reply_document(document=doc_file)
            await asyncio.sleep(1.2)
        except Exception as e:
            print(f"Fayl yuborishda xatolik: {e}")
        finally:
            if os.path.exists(file_p):
                os.remove(file_p)

    if os.path.exists(temp_dir):
        os.rmdir(temp_dir)

    await update.message.reply_text("✅ Barcha 14 ta shakl muvaffaqiyatli tayyorlandi va yuborildi!")
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
    print("Боғланиш тайёр...")
    app.run_polling()

if __name__ == '__main__':
    main()
