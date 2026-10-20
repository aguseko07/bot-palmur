import telebot
import openpyxl
import os

TOKEN = '8037929349:AAHHYxyZRC-Gpz7pTND2y55NMl5_-DR4wW8'
bot = telebot.TeleBot(TOKEN)

EXCEL_PALMUR = 'Laporan_Palmur.xlsx'

# DAFTAR ITEM DAN TARGET DALAM SATUAN PCS
DAFTAR_ITEM = [
    {"id": 0, "nama_item": "Minyak Goreng 2L", "target": 100},
    {"id": 1, "nama_item": "Beras Premium 5kg", "target": 150},
    {"id": 2, "nama_item": "Gula Pasir 1kg", "target": 80},
    # Tambahkan item lainnya di sini jika perlu
]

# Menyimpan sesi input aktif pengguna
user_sessions = {}

def init_excel():
    if not os.path.exists(EXCEL_PALMUR):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Palmur"
        ws.append(["No", "Nama Item", "Target (Pcs)", "Sales (Pcs)", "Ach (%)"])
        wb.save(EXCEL_PALMUR)

init_excel()

def format_pcs(angka):
    if angka is None or angka == "":
        return "-"
    if isinstance(angka, (int, float)):
        return f"{angka:,.0f} pcs".replace(',', '.')
    return str(angka)

def format_val(val):
    if val is None or val == "":
        return "-"
    if isinstance(val, float):
        return f"{val:.1f}%"
    return str(val)

@bot.message_handler(commands=['start'])
def send_welcome(message):
    bantuan = (
        "🌴 *Bot Laporan Palmur Interaktif V4.0*\n\n"
        "Gunakan perintah `/input` untuk mulai mengisi laporan penjualan item demi item secara berurutan.\n\n"
        "📋 `/laporan` - Lihat rekap seluruh item Palmur\n"
        "🗑️ `/hapus` - Hapus data terakhir"
    )
    bot.reply_to(message, bantuan, parse_mode='Markdown')

# ==========================================
# MEMULAI SESI INPUT INTERAKTIF
# ==========================================
@bot.message_handler(commands=['input', 'palmur'])
def start_interactive_input(message):
    chat_id = message.chat.id
    # Mulai dari item pertama (index 0)
    user_sessions[chat_id] = {
        "step": 0,
        "results": []
    }
    tanya_item berikutnya(chat_id)

def tanya_item_berikutnya(chat_id):
    session = user_sessions.get(chat_id)
    if not session:
        return
        
    step = session["step"]
    if step < len(DAFTAR_ITEM):
        item = DAFTAR_ITEM[step]
        pesan = (
            f"📝 *Input Penjualan Palmur ({step + 1}/{len(DAFTAR_ITEM)})\n\n"
            f"🏷️ Item: *{item['nama_item']}*\n"
            f"🎯 Target: {item['target']} pcs\n\n"
            f"👇 *Berapa jumlah sales (pcs) untuk item ini?*\n"
            f"_(Ketik langsung angkanya di chat)_"
        )
        bot.send_message(chat_id, pesan, parse_mode='Markdown')
    else:
        # Selesai, simpan semua ke Excel
        simpan_semua_hasil(chat_id)

@bot.message_handler(func=lambda message: message.chat.id in user_sessions)
def handle_text_input(message):
    chat_id = message.chat.id
    session = user_sessions[chat_id]
    
    try:
        sales_val = float(message.text.replace('.', '').replace(',', ''))
        step = session["step"]
        item = DAFTAR_ITEM[step]
        
        target = item['target']
        ach = (sales_val / target * 100) if target > 0 else 0
        
        # Simpan hasil sementara ke sesi
        session["results"].append({
            "nama_item": item['nama_item'],
            "target": target,
            "sales": sales_val,
            "ach": ach
        })
        
        # Lanjut ke item berikutnya
        session["step"] += 1
        
        if session["step"] < len(DAFTAR_ITEM):
            tanya_item_berikutnya(chat_id)
        else:
            simpan_semua_hasil(chat_id)
            
    except ValueError:
        bot.reply_to(message, "❌ Masukkan angka sales (pcs) yang valid.")

def simpan_semua_hasil(chat_id):
    session = user_sessions.pop(chat_id, None)
    if not session or not session["results"]:
        return
        
    wb = openpyxl.load_workbook(EXCEL_PALMUR)
    ws = wb.active
    
    ringkasan = "✅ *Semua Penjualan Palmur Berhasil Disimpan!*\n\n"
    
    for res in session["results"]:
        next_row = ws.max_row + 1
        no = next_row - 1
        ws.append([no, res["nama_item"], res["target"], res["sales"], res["ach"]])
        
        ringkasan += f"• *{res['nama_item']}* ➔ Sales: {format_pcs(res['sales'])} (Ach: {format_val(res['ach'])})\n"
        
    wb.save(EXCEL_PALMUR)
    
    bot.send_message(chat_id, ringkasan, parse_mode='Markdown')
    
    with open(EXCEL_PALMUR, 'rb') as doc:
        bot.send_document(chat_id, doc, caption="📂 File Excel Palmur Terbaru")

# ==========================================
# COMMAND: /laporan
# ==========================================
@bot.message_handler(commands=['laporan', 'rekap'])
def handle_laporan_palmur(message):
    try:
        wb = openpyxl.load_workbook(EXCEL_PALMUR, data_only=True)
        ws = wb.active
        
        if ws.max_row <= 1:
            bot.reply_to(message, "❌ Belum ada data item Palmur yang tersimpan.")
            return

        pesan = "📋 *REKAP LAPORAN PALMUR (PCS)*\n"
        pesan += "--------------------------------------\n"
        
        total_target = 0
        total_sales = 0
        
        for r in range(2, ws.max_row + 1):
            no = ws.cell(row=r, column=1).value
            item = ws.cell(row=r, column=2).value
            tgt = ws.cell(row=r, column=3).value or 0
            sls = ws.cell(row=r, column=4).value or 0
            ach = ws.cell(row=r, column=5).value or 0
            
            total_target += tgt
            total_sales += sls
            
            pesan += f"{no}. *{item}*\n"
            pesan += f"   Tgt: {format_pcs(tgt)} | Sls: {format_pcs(sls)} | Ach: *{format_val(ach)}*\n"
            
        total_ach = (total_sales / total_target * 100) if total_target > 0 else 0
        
        pesan += "--------------------------------------\n"
        pesan += f"🎯 *Total Target:* {format_pcs(total_target)}\n"
        pesan += f"📦 *Total Sales:* {format_pcs(total_sales)}\n"
        pesan += f"📈 *Ach Keseluruhan:* *{format_val(total_ach)}*"
        
        bot.reply_to(message, pesan, parse_mode='Markdown')
        
        with open(EXCEL_PALMUR, 'rb') as doc:
            bot.send_document(message.chat.id, doc, caption="📂 File Excel Palmur Terbaru")
            
    except Exception as e:
        bot.reply_to(message, f"❌ Gagal mengambil laporan: {e}")

# ==========================================
# COMMAND: /hapus
# ==========================================
@bot.message_handler(commands=['hapus'])
def handle_hapus_palmur(message):
    try:
        wb = openpyxl.load_workbook(EXCEL_PALMUR)
        ws = wb.active
        
        if ws.max_row <= 1:
            bot.reply_to(message, "❌ Tidak ada data yang bisa dihapus.")
            return
            
        deleted_item = ws.cell(row=ws.max_row, column=2).value
        ws.delete_rows(ws.max_row)
        wb.save(EXCEL_PALMUR)
        
        bot.reply_to(message, f"🗑️ Data item terakhir (*{deleted_item}*) berhasil dihapus.")
    except Exception as e:
        bot.reply_to(message, f"❌ Gagal menghapus: {e}")

if __name__ == '__main__':
    print("🤖 Bot Palmur Interaktif V4.0 Berhasil Dijalankan...")
    bot.polling(none_stop=True, interval=0)
