import os
import re
import threading
from flask import Flask
from collections import defaultdict
from datetime import datetime

import pandas as pd
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    MessageHandler,
    CommandHandler,
    ContextTypes,
    filters,
)

# ================== CONFIG ==================

TOKEN = os.getenv("TOKEN")
if not TOKEN:
    raise ValueError("TOKEN belum diset di Environment Variables Render")

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
        for p, j in sorted(data[(year, month, day)].items()):
            hasil += f"- {p} : {j}\n"
        hasil += "\n"

    return hasil


# ================== HANDLERS ==================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Halo!\n\nKirim file Excel (.xlsx) untuk mulai rekap."
    )


async def tampilkan_menu_bulan(update: Update, context: ContextTypes.DEFAULT_TYPE):
    sheets = context.user_data["sheets"]

    pesan = "📅 *Pilih bulan laporan:*\n\n"
    for i, s in enumerate(sheets, start=1):
        pesan += f"{i}. {s}\n"
    pesan += "\nKetik nomor bulan (contoh: 1)"

    context.user_data["menunggu_pilih_bulan"] = True
    await update.message.reply_text(pesan, parse_mode="Markdown")


async def handle_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        file = await update.message.document.get_file()
        await file.download_to_drive("data.xlsx")

        xls = pd.ExcelFile("data.xlsx")
        sheets = xls.sheet_names
        xls.close()

        context.user_data.clear()
        context.user_data["sheets"] = sheets

        await tampilkan_menu_bulan(update, context)

    except Exception as e:
        await update.message.reply_text(f"❌ Gagal membaca file.\n{e}")


async def handle_pilih_bulan(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.user_data.get("menunggu_pilih_bulan"):
        return

    try:
        pilihan = int(update.message.text)
        sheets = context.user_data["sheets"]
        sheet_name = sheets[pilihan - 1]
    except:
        await update.message.reply_text("❌ Pilihan bulan tidak valid.")
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
        await update.message.reply_text("❌ Header tabel tidak ditemukan.")
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
        await update.message.reply_text(
            f"❌ Kolom tidak lengkap.\nKolom terbaca: {list(df.columns)}"
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

    await update.message.reply_text(
        f"📊 *Rekap {sheet_name}*\n\n{hasil}",
        parse_mode="Markdown"
    )

    context.user_data["menunggu_pilih_bulan"] = False
    context.user_data["menunggu_lanjut"] = True

    await update.message.reply_text(
        "📅 Mau rekap bulan lain?\nKetik: *ya* / *tidak*",
        parse_mode="Markdown"
    )


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.lower().strip()

    if context.user_data.get("menunggu_lanjut"):
        if text == "ya":
            context.user_data["menunggu_lanjut"] = False
            await tampilkan_menu_bulan(update, context)
        elif text == "tidak":
            context.user_data.clear()
            await update.message.reply_text("✅ Selesai. Terima kasih.")
        else:
            await update.message.reply_text("Ketik *ya* atau *tidak*.", parse_mode="Markdown")


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    print(f"Terjadi error: {context.error}")


# ================== FLASK ==================

app = Flask(__name__)

@app.route("/")
def home():
    return "Bot is running!"

def run_web():
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)

threading.Thread(target=run_web).start()


# ================== MAIN ==================

app_bot = ApplicationBuilder().token(TOKEN).build()

app_bot.add_handler(CommandHandler("start", start))

app_bot.add_handler(
    MessageHandler(
        filters.Document.MimeType(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        handle_excel,
    )
)

app_bot.add_handler(
    MessageHandler(filters.TEXT & filters.Regex(r"^\d+$"), handle_pilih_bulan)
)

app_bot.add_handler(
    MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text)
)

app_bot.add_error_handler(error_handler)

app_bot.run_polling()
