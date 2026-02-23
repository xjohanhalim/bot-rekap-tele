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

TOKEN = os.getenv("TOKEN")

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


# ================== HANDLERS ==================

async def handle_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file = await update.message.document.get_file()
    await file.download_to_drive("data.xlsx")

    xls = pd.ExcelFile("data.xlsx")
    sheets = xls.sheet_names
    xls.close()

    context.user_data.clear()
    context.user_data["sheets"] = sheets
    context.user_data["menunggu_pilih_bulan"] = True

    pesan = "📅 *Pilih bulan laporan:*\n\n"
    for i, s in enumerate(sheets, start=1):
        pesan += f"{i}. {s}\n"
    pesan += "\nKetik nomor bulan (contoh: 1)"

    await update.message.reply_text(pesan, parse_mode="Markdown")


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

    rows = [
        (r["tanggal"], r["lokasi"], r["paket"], r["jumlah"])
        for _, r in df.iterrows()
        if not pd.isna(r["tanggal"]) and not pd.isna(r["paket"])
    ]

    hasil = rekap_data(rows)

    await update.message.reply_text(
        f"📊 *Rekap {sheet_name}*\n\n{hasil}",
        parse_mode="Markdown"
    )

    context.user_data.clear()


# ================== FLASK (RENDER FREE) ==================

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

app_bot.add_handler(
    MessageHandler(
        filters.Document.MimeType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        handle_excel
    )
)

app_bot.add_handler(
    MessageHandler(filters.TEXT & filters.Regex(r"^\d+$"), handle_pilih_bulan)
)

app_bot.run_polling()
