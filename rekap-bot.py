from telegram.ext import Updater, MessageHandler, Filters, CommandHandler
from collections import defaultdict
import pandas as pd
import os
import re
import pytesseract
from PIL import Image
from datetime import datetime

# ================== CONFIG ==================
TOKEN = os.getenv("BOT_TOKEN")

# ================== UTIL ==================
def normalisasi_paket(paket):
    paket = str(paket).upper()
    jam = re.search(r'([1-5])\s*JAM', paket)
    if not jam:
        return None
    jam = jam.group(1)
    return f"b2g3 {jam} jam" if "B2G3" in paket else f"{jam} jam"


def rekap_data(rows):
    data = defaultdict(lambda: defaultdict(int))

    for tanggal, lokasi, paket, jumlah in rows:
        paket_norm = normalisasi_paket(paket)
        if not paket_norm:
            continue

        try:
            jumlah = int(float(jumlah))
        except:
            continue

        try:
            if isinstance(tanggal, datetime):
                day = tanggal.day
                month = tanggal.strftime("%b").lower()
                year = tanggal.year
            else:
                t = str(tanggal).split(" ")[0]
                dt = datetime.fromisoformat(t)
                day = dt.day
                month = dt.strftime("%b").lower()
                year = dt.year
        except:
            continue

        data[(year, month, day)][paket_norm] += jumlah

    if not data:
        return "⚠️ Tidak ada data yang bisa direkap."

    hasil = ""
    for (year, month, day) in sorted(data):
        hasil += f"{day} {month}\n"
        for p, j in data[(year, month, day)].items():
            hasil += f"- {p} : {j}\n"
        hasil += "\n"

    return hasil


# ================== HANDLER ==================
def tampilkan_menu_bulan(update, context):
    sheets = context.user_data["sheets"]

    pesan = "📅 *Pilih bulan laporan:*\n\n"
    for i, s in enumerate(sheets, start=1):
        pesan += f"{i}. {s}\n"
    pesan += "\nKetik nomor bulan (contoh: 1)"

    update.message.reply_text(pesan, parse_mode="Markdown")
    context.user_data["menunggu_pilih_bulan"] = True


def handle_excel(update, context):
    file = update.message.document.get_file()
    file.download("data.xlsx")

    xls = pd.ExcelFile("data.xlsx")
    sheets = xls.sheet_names
    xls.close()

    context.user_data.clear()
    context.user_data["excel_path"] = "data.xlsx"
    context.user_data["sheets"] = sheets

    tampilkan_menu_bulan(update, context)


def handle_pilih_bulan(update, context):
    if not context.user_data.get("menunggu_pilih_bulan"):
        return

    try:
        pilihan = int(update.message.text)
        sheets = context.user_data["sheets"]
        sheet_name = sheets[pilihan - 1]
    except:
        update.message.reply_text("❌ Pilihan bulan tidak valid.")
        return

    df_raw = pd.read_excel("data.xlsx", sheet_name=sheet_name, header=None)

    header_row = None
    for i in range(5):
        row = df_raw.iloc[i].astype(str).str.lower()
        if (
            row.str.contains("tanggal").any()
            and row.str.contains("paket").any()
            and row.str.contains("jumlah").any()
        ):
            header_row = i
            break

    if header_row is None:
        update.message.reply_text("❌ Header tabel tidak ditemukan.")
        return

    df = pd.read_excel("data.xlsx", sheet_name=sheet_name, header=header_row)
    df.columns = df.columns.astype(str).str.strip().str.lower()

    col_map = {}
    for col in df.columns:
        if "tanggal" in col:
            col_map["tanggal"] = col
        elif "lokasi" in col:
            col_map["lokasi"] = col
        elif "paket" in col:
            col_map["paket"] = col
        elif "jumlah" in col:
            col_map["jumlah"] = col

    if len(col_map) < 4:
        update.message.reply_text(
            f"❌ Kolom tidak lengkap di sheet {sheet_name}.\n"
            f"Kolom terbaca: {list(df.columns)}"
        )
        return

    df = df.dropna(subset=[col_map["tanggal"], col_map["paket"], col_map["jumlah"]])

    rows = [
        (
            r[col_map["tanggal"]],
            r[col_map["lokasi"]],
            r[col_map["paket"]],
            r[col_map["jumlah"]],
        )
        for _, r in df.iterrows()
    ]

    hasil = rekap_data(rows)

    update.message.reply_text(
        f"📊 *Rekap {sheet_name}*\n\n{hasil}",
        parse_mode="Markdown"
    )

    context.user_data["menunggu_pilih_bulan"] = False
    context.user_data["menunggu_lanjut"] = True

    update.message.reply_text(
        "📅 Mau rekap bulan lain?\nKetik: *ya* / *tidak*",
        parse_mode="Markdown"
    )


def handle_text(update, context):
    text = update.message.text.lower().strip()

    if context.user_data.get("menunggu_lanjut"):
        if text == "ya":
            context.user_data["menunggu_lanjut"] = False
            tampilkan_menu_bulan(update, context)
        elif text == "tidak":
            context.user_data.clear()
            update.message.reply_text("✅ Selesai. Terima kasih.")
        else:
            update.message.reply_text("Ketik *ya* atau *tidak*.", parse_mode="Markdown")


# ================== MAIN ==================
updater = Updater(
    TOKEN,
    use_context=True,
    request_kwargs={
        "connect_timeout": 15,
        "read_timeout": 15
    }
)

dp = updater.dispatcher
dp.add_handler(MessageHandler(
    Filters.document.mime_type(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ),
    handle_excel
))
dp.add_handler(MessageHandler(Filters.text & Filters.regex(r"^\d+$"), handle_pilih_bulan))
dp.add_handler(MessageHandler(Filters.text & ~Filters.command, handle_text))

try:
    updater.start_polling(poll_interval=1.0, timeout=30)
    updater.idle()
except Exception as e:
    print("⚠️ Bot berhenti karena error jaringan:", e)
