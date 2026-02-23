import os
import re
import threading
from flask import Flask
from collections import defaultdict
from datetime import datetime

import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    MessageHandler,
    CommandHandler,
    CallbackQueryHandler,
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


async def handle_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file = await update.message.document.get_file()
    await file.download_to_drive("data.xlsx")

    xls = pd.ExcelFile("data.xlsx")
    sheets = xls.sheet_names
    xls.close()

    context.user_data["sheets"] = sheets

    keyboard = [
        [InlineKeyboardButton(s, callback_data=f"sheet_{i}")]
        for i, s in enumerate(sheets)
    ]

    await update.message.reply_text(
        "📅 Pilih bulan laporan:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )


async def handle_sheet(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    index = int(query.data.split("_")[1])
    sheet_name = context.user_data["sheets"][index]

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
        await query.message.reply_text("❌ Header tabel tidak ditemukan.")
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
        await query.message.reply_text(
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

    keyboard = [
        [
            InlineKeyboardButton("🔁 Rekap Lagi", callback_data="again"),
            InlineKeyboardButton("❌ Selesai", callback_data="done"),
        ]
    ]

    await query.message.reply_text(
        f"📊 *Rekap {sheet_name}*\n\n{hasil}",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )


async def handle_again_done(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "again":

        # Kalau session sudah habis
        if "sheets" not in context.user_data:
            await query.message.reply_text(
                "⚠️ Session sudah berakhir.\nSilakan kirim ulang file Excel untuk mulai lagi."
            )
            return

        sheets = context.user_data["sheets"]

        keyboard = [
            [InlineKeyboardButton(s, callback_data=f"sheet_{i}")]
            for i, s in enumerate(sheets)
        ]

        await query.message.reply_text(
            "📅 Pilih bulan laporan:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif query.data == "done":
        context.user_data.clear()
        await query.message.reply_text("✅ Selesai. Terima kasih.")


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

app_bot.add_handler(CallbackQueryHandler(handle_sheet, pattern=r"^sheet_"))
app_bot.add_handler(CallbackQueryHandler(handle_again_done, pattern="^(again|done)$"))

app_bot.run_polling()
