import telebot
import os
import shutil
from datetime import datetime

pwd = os.getcwd()
token = open(pwd+'/cleaning/token.txt', 'r').read() #read bot token in token.txt
bot = telebot.TeleBot(token)
dead_line = datetime(2020, 11, 9)
local_dt = datetime.now()
folder = "11W2"

@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.send_chat_action(message.chat.id, 'typing')
    bot.reply_to(message, f"Semangat Pageee!!! {message.chat.first_name} {message.chat.last_name}. \nJika ini pertama kali, perkenalkan aku adalah Bot Treg 5 Order Cleaner. \n"+
        "\nAku ada untuk menampung order-order IndiHome dari Witel di Treg5 (usia > 7hari) yang akan diajukan proses cleaning\n"+
        "\nSebelum kirim file .xlsx nya, pastikan sudah mengikuti format di KPRO.\n"
        "\nUntuk melihat contoh format filenya ketik command:\n"+
        "/format\n"+
        "\nUntuk bantuan lain ketik command:\n"+
        "/help\n")

@bot.message_handler(commands=['chat_id'])
def chat_id(message):
    updates = bot.get_updates(1234, 100, 20)
    bot.reply_to(message,updates[0].message.chat_id)

@bot.message_handler(commands=['format'])
def format(message):
    bot.reply_to(message, "berikut format yang diterima:")
    # sendPhoto
    bot.send_chat_action(message.chat.id, 'upload_photo')
    photo = open(pwd + '/cleaning/files/contoh_cleaning.png', 'rb')
    bot.send_photo(message.chat.id, photo)
    # sendDocument
    bot.send_chat_action(message.chat.id, 'upload_document')
    doc = open(pwd + '/cleaning/files/contoh_cleaning.xlsx', 'rb')
    bot.send_document(message.chat.id, doc)

@bot.message_handler(commands=['help'])
def help(message):
    bot.reply_to(message, "/format (untuk melihat contoh format file )\n"+
                 "/help (untuk bantuan)\n"+
                 "/deadline (untuk info deadline pengumpulan)\n"+
                 "/dev (untuk info developer)")

@bot.message_handler(commands=['dev'])
def dev(message):
    bot.send_chat_action(message.chat.id, 'typing')
    bot.reply_to(message, "kritik dan saran bisa hubungi @basnugroho 👨🏻‍💻")

@bot.message_handler(commands=['deadline'])
def dev(message):
    bot.send_chat_action(message.chat.id, 'typing')
    bot.reply_to(message, f'Untuk periode {folder} harap kirim dokumen .xlsx sebelum {str(dead_line)}')

@bot.message_handler(content_types=['document', 'photo'])
def handle_excel(message):
    if datetime.now() > dead_line:
        bot.reply_to(message, f"maaf waktu pengumpulan untuk periode {folder} dengan deadline {str(dead_line)} telah terlewati! Hubungi @basnugroho untuk info lebih lanjut.")
    else:
        if message.document.file_name.split(sep=".")[1] != "xlsx" or message.document.mime_type != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
            bot.reply_to(message, "maaf tipe file belum sesuai (bukan .xlsx)")
        else:
            file_info = bot.get_file(message.document.file_id)
            downloaded_file = bot.download_file(file_info.file_path)
            with open(message.document.file_name, 'wb') as new_file:
                new_file.write(downloaded_file)
                shutil.move(pwd+"/"+message.document.file_name, pwd+"/cleaning/"+folder+"/"+message.document.file_name)
            bot.reply_to(message, f"file {message.document.file_name} diterima ")

@bot.message_handler(commands=['set'])
def set_deadline(message):
    if message.text.split("/set")[1] == "^date":
        year = message.text.split(" ")[2]
        month = message.text.split(" ")[3]
        day = message.text.split(" ")[4]
        dead_line = datetime(year, month, day)
        bot.reply_to(message, f"deadline telah diset ke: {str(dead_line)}")
    else:
        bot.reply_to(message, "format salah! Contoh format untuk set deadline: '#set date 2020, 11, 15'")

@bot.message_handler(func=lambda m: True)
def echo_all(message):
    bot.reply_to(message, f"maaf command '{message.text}' belu dikenali. ketik /help untuk list command yang dikenali")

# @bot.message_handler(func=lambda m: True)
# def echo_all(message):
#     if message.text.split("#")[1]=="^set":
#         if message.text.split("#set")[1]=="^date":
#             year = message.text.split(" ")[2]
#             month = message.text.split(" ")[3]
#             day = message.text.split(" ")[4]
#             dead_line = datetime(year, month, day)
#             bot.reply_to(message, f"oke, deadline telah diset ke: {str(dead_line)}")
#         else:
#             bot.reply_to(message, "format salah! Contoh format untuk set deadline: '#set date 2020, 11, 15'")
#     else:
#         bot.reply_to(message, "command " + 'message.text' + " tidak dikenali")
bot.polling()