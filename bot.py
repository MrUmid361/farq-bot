"""
Mahsulotlar solishtiruvchi Telegram bot.

Foydalanuvchi ikkita fayl (xlsx yoki xls — jumladan 1C dan chiqadigan,
HTML formatida "yolg'on" .xls fayllar ham) yuboradi. Bot ustunlarni
avtomatik aniqlaydi (nomi, miqdor, narx, summa, kod), so'ngra ikkala
fayl orasidagi haqiqiy raqamli farqlarni topib, tayyor Excel hisobot
qilib qaytaradi.

O'RNATISH:
    pip install python-telegram-bot==21.* pandas openpyxl beautifulsoup4 lxml xlrd

ISHGA TUSHIRISH:
    export BOT_TOKEN="123456:ABC-tokeningiz"
    python bot.py
"""

import io
import os
import logging
import re
from dataclasses import dataclass, field

import pandas as pd
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from telegram import Update, Document
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    ConversationHandler,
    filters,
)

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN", "")

# Conversation holati
WAITING_FILE_1, WAITING_FILE_2 = range(2)

# Ustunlarni aniqlash uchun kalit so'zlar (ruscha/o'zbekcha, kichik harflarda)
NAME_KEYWORDS = ["номенклатура", "маҳсулот", "название", "наименование", "nomi", "товар"]
CODE_KEYWORDS = ["код", "kod"]
QTY_KEYWORDS = ["количество", "миқдор", "кол-во", "miqdor"]
PRICE_KEYWORDS = ["цена", "нарх", "нарҳ", "narx"]
# "сумма"/"қиймат" so'zi bor, lekin НДС/QQS/жами so'zlari yo'q ustun — asosiy summa
SUM_KEYWORDS = ["сумма", "қиймат", "summa", "qiymat"]
EXCLUDE_IN_SUM = ["ндс", "ққс", "жами", "vat", "qqsni"]


@dataclass
class ParsedFile:
    df: pd.DataFrame  # columns: name, code, qty, price, sum
    raw_columns: list = field(default_factory=list)


def _norm(s):
    return str(s).strip().lower()


def _to_number(x):
    """'7 668,00' yoki '7668.00' yoki 7668.0 -> float"""
    if x is None:
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace("\xa0", "").replace(" ", "")
    s = s.replace(",", ".")
    s = re.sub(r"[^0-9.\-]", "", s)
    if s in ("", "-", "."):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _pick_column(columns, keywords, exclude=None):
    exclude = exclude or []
    for col in columns:
        c = _norm(col)
        if any(k in c for k in keywords) and not any(e in c for e in exclude):
            return col
    return None


def read_xlsx(path: str) -> ParsedFile:
    df_raw = pd.read_excel(path)
    cols = list(df_raw.columns)

    name_col = _pick_column(cols, NAME_KEYWORDS)
    code_col = _pick_column(cols, CODE_KEYWORDS)
    qty_col = _pick_column(cols, QTY_KEYWORDS)
    price_col = _pick_column(cols, PRICE_KEYWORDS)
    sum_col = _pick_column(cols, SUM_KEYWORDS, exclude=EXCLUDE_IN_SUM)

    if not all([name_col, qty_col, price_col, sum_col]):
        raise ValueError(
            "Ustunlarni avtomatik aniqlab bo'lmadi. Topilgan ustunlar: " + ", ".join(map(str, cols))
        )

    out = pd.DataFrame()
    out["name"] = df_raw[name_col].astype(str).str.strip()
    out["code"] = df_raw[code_col].astype(str).str.strip().str.lstrip("0") if code_col else ""
    out["qty"] = df_raw[qty_col].apply(_to_number)
    out["price"] = df_raw[price_col].apply(_to_number)
    out["sum"] = df_raw[sum_col].apply(_to_number)
    out = out.dropna(subset=["qty", "price", "sum"])
    return ParsedFile(df=out, raw_columns=cols)


def _find_product_table(soup: BeautifulSoup):
    """1C / hisob-faktura HTML jadvallari orasidan mahsulotlar jadvalini topadi:
    eng ko'p qatorli va birinchi ustuni ketma-ket raqamlardan iborat jadval."""
    best = None
    best_rows = 0
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        numeric_rows = 0
        for r in rows:
            cells = r.find_all(["td", "th"])
            if cells and cells[0].get_text(strip=True).isdigit():
                numeric_rows += 1
        if numeric_rows > best_rows:
            best_rows = numeric_rows
            best = table
    return best, best_rows


def read_html_xls(path: str) -> ParsedFile:
    with open(path, "rb") as f:
        raw = f.read()

    encoding = "utf-8"
    m = re.search(rb"charset=([\w-]+)", raw[:2000], re.IGNORECASE)
    if m:
        encoding = m.group(1).decode("ascii", errors="ignore")

    html = raw.decode(encoding, errors="ignore")
    soup = BeautifulSoup(html, "lxml")

    table, n_rows = _find_product_table(soup)
    if table is None or n_rows < 2:
        raise ValueError("Faylda mahsulotlar jadvali topilmadi.")

    data_rows = []
    for r in table.find_all("tr"):
        cells = [c.get_text(strip=True) for c in r.find_all(["td", "th"])]
        if cells and cells[0].isdigit() and len(cells) >= 7:
            data_rows.append(cells)

    if not data_rows:
        raise ValueError("Faylda mahsulot qatorlari topilmadi.")

    # Odatiy hisob-faktura tartibi: № | Nomi | Kod-nomi | O'lchov | Miqdor | Narx | Summa | ...
    ncols = len(data_rows[0])
    name_idx = 1
    code_idx = 2 if ncols > 2 else None
    qty_idx = 4 if ncols > 4 else 3
    price_idx = 5 if ncols > 5 else 4
    sum_idx = 6 if ncols > 6 else 5

    records = []
    for cells in data_rows:
        try:
            name = cells[name_idx]
            code = ""
            if code_idx is not None and code_idx < len(cells):
                code = cells[code_idx].split("-")[0].strip().lstrip("0")
            qty = _to_number(cells[qty_idx])
            price = _to_number(cells[price_idx])
            total = _to_number(cells[sum_idx])
            if qty is None or price is None or total is None:
                continue
            records.append((name, code, qty, price, total))
        except IndexError:
            continue

    df = pd.DataFrame(records, columns=["name", "code", "qty", "price", "sum"])
    return ParsedFile(df=df, raw_columns=[f"col{i}" for i in range(ncols)])


def read_any(path: str) -> ParsedFile:
    """.xlsx -> pandas/openpyxl. .xls -> avval haqiqiy BIFF xls (xlrd),
    bo'lmasa HTML-asosli xls sifatida o'qiladi."""
    lower = path.lower()
    if lower.endswith(".xlsx") or lower.endswith(".xlsm"):
        return read_xlsx(path)

    if lower.endswith(".xls"):
        try:
            return read_xlsx(path)  # pandas ba'zan xlrd orqali BIFF xls'ni ham o'qiy oladi
        except Exception:
            return read_html_xls(path)

    raise ValueError("Faqat .xlsx yoki .xls fayllar qo'llab-quvvatlanadi.")


def compare_files(pf1: ParsedFile, pf2: ParsedFile) -> tuple[bytes, dict]:
    """Ikki faylni (qty, price) bo'yicha moslashtirib, faqat haqiqiy farqlarni qaytaradi."""
    from collections import Counter

    df1, df2 = pf1.df.copy(), pf2.df.copy()

    def key(row):
        return (round(row["qty"], 3), round(row["price"], 2))

    c1 = Counter(key(r) for _, r in df1.iterrows())
    c2 = Counter(key(r) for _, r in df2.iterrows())

    only1 = c1 - c2
    only2 = c2 - c1

    rows_out = []
    for k, v in only1.items():
        matches = df1[(df1["qty"].round(3) == k[0]) & (df1["price"].round(2) == k[1])]
        for _, m in matches.iterrows():
            rows_out.append(("1-faylda bor, 2-faylda yo'q", m["name"], m["code"], m["qty"], m["price"], m["sum"]))
    for k, v in only2.items():
        matches = df2[(df2["qty"].round(3) == k[0]) & (df2["price"].round(2) == k[1])]
        for _, m in matches.iterrows():
            rows_out.append(("2-faylda bor, 1-faylda yo'q", m["name"], m["code"], m["qty"], m["price"], m["sum"]))

    diff_df = pd.DataFrame(rows_out, columns=["Manba", "Nomi", "Kod", "Miqdor", "Narx", "Summa"])
    diff_df = diff_df.sort_values(["Manba", "Nomi"]).reset_index(drop=True)

    total1 = df1["sum"].sum()
    total2 = df2["sum"].sum()

    # ---- Excel hisobot yasash ----
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Xulosa"
    bold = Font(name="Arial", bold=True)
    title_font = Font(name="Arial", bold=True, size=14)

    ws["A1"] = "Fayllar orasidagi farq hisoboti"
    ws["A1"].font = title_font
    ws.merge_cells("A1:D1")

    ws["A3"] = "1-fayl qatorlar soni"
    ws["B3"] = len(df1)
    ws["A4"] = "2-fayl qatorlar soni"
    ws["B4"] = len(df2)
    ws["A5"] = "1-fayl jami summasi"
    ws["B5"] = round(total1, 2)
    ws["A6"] = "2-fayl jami summasi"
    ws["B6"] = round(total2, 2)
    ws["A7"] = "Umumiy farq (2-fayl − 1-fayl)"
    ws["B7"] = "=B6-B5"
    for r in (3, 4, 5, 6, 7):
        ws[f"A{r}"].font = Font(name="Arial")
    ws["A7"].font = bold
    ws["B7"].font = bold
    for r in (5, 6, 7):
        ws[f"B{r}"].number_format = "#,##0.00"
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 18

    ws2 = wb.create_sheet("Farqlar")
    headers = ["№", "Manba", "Nomi", "Kod", "Miqdor", "Narx", "Summa"]
    ws2.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws2.cell(row=1, column=c)
        cell.font = bold
        cell.fill = PatternFill("solid", fgColor="D9E1F2")

    for i, row in diff_df.iterrows():
        r = i + 2
        ws2.cell(row=r, column=1, value=i + 1)
        ws2.cell(row=r, column=2, value=row["Manba"])
        ws2.cell(row=r, column=3, value=row["Nomi"])
        ws2.cell(row=r, column=4, value=row["Kod"])
        ws2.cell(row=r, column=5, value=row["Miqdor"])
        ws2.cell(row=r, column=6, value=row["Narx"])
        ws2.cell(row=r, column=7, value=row["Summa"])
        ws2.cell(row=r, column=5).number_format = "#,##0.00"
        ws2.cell(row=r, column=6).number_format = "#,##0.00"
        ws2.cell(row=r, column=7).number_format = "#,##0.00"

    widths = [5, 28, 40, 20, 12, 14, 15]
    for i, w in enumerate(widths, start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    summary = {
        "total1": total1,
        "total2": total2,
        "diff": total2 - total1,
        "n_diff_rows": len(diff_df),
    }
    return buf.getvalue(), summary


# ---------------- Telegram handlerlar ----------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "Salom! Men ikkita Excel (.xlsx yoki .xls) faylni solishtirib, "
        "ular orasidagi haqiqiy farqlarni topib beraman.\n\n"
        "Birinchi faylni yuboring."
    )
    return WAITING_FILE_1


async def receive_file_1(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc: Document = update.message.document
    if not doc or not doc.file_name.lower().endswith((".xlsx", ".xls", ".xlsm")):
        await update.message.reply_text("Iltimos, .xlsx yoki .xls fayl yuboring.")
        return WAITING_FILE_1

    path = f"/tmp/{update.effective_chat.id}_1_{doc.file_name}"
    file = await doc.get_file()
    await file.download_to_drive(path)
    context.user_data["file1_path"] = path

    await update.message.reply_text("1-fayl qabul qilindi ✅\nEndi ikkinchi faylni yuboring.")
    return WAITING_FILE_2


async def receive_file_2(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc: Document = update.message.document
    if not doc or not doc.file_name.lower().endswith((".xlsx", ".xls", ".xlsm")):
        await update.message.reply_text("Iltimos, .xlsx yoki .xls fayl yuboring.")
        return WAITING_FILE_2

    path = f"/tmp/{update.effective_chat.id}_2_{doc.file_name}"
    file = await doc.get_file()
    await file.download_to_drive(path)

    await update.message.reply_text("2-fayl qabul qilindi ✅. Solishtirilmoqda...")

    try:
        pf1 = read_any(context.user_data["file1_path"])
        pf2 = read_any(path)
        report_bytes, summary = compare_files(pf1, pf2)
    except Exception as e:
        logger.exception("Solishtirishda xatolik")
        await update.message.reply_text(f"❌ Xatolik yuz berdi: {e}")
        return ConversationHandler.END

    text = (
        f"✅ Solishtirish tugadi.\n\n"
        f"1-fayl jami: {summary['total1']:,.2f}\n"
        f"2-fayl jami: {summary['total2']:,.2f}\n"
        f"Umumiy farq: {summary['diff']:,.2f}\n"
        f"Farqli qatorlar soni: {summary['n_diff_rows']}\n\n"
        f"To'liq hisobot faylini yuboryapman 👇"
    )
    await update.message.reply_text(text)
    await update.message.reply_document(
        document=io.BytesIO(report_bytes), filename="farqlar_hisoboti.xlsx"
    )

    context.user_data.clear()
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("Bekor qilindi. Qayta boshlash uchun /start yuboring.")
    return ConversationHandler.END


def main():
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN muhit o'zgaruvchisi topilmadi. `export BOT_TOKEN=...` qiling.")

    app = Application.builder().token(BOT_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            WAITING_FILE_1: [MessageHandler(filters.Document.ALL, receive_file_1)],
            WAITING_FILE_2: [MessageHandler(filters.Document.ALL, receive_file_2)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    app.add_handler(conv)

    logger.info("Bot ishga tushdi...")
    app.run_polling()


if __name__ == "__main__":
    main()
