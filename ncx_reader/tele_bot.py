import telebot
import pandas as pd
import os

TOKEN = "1462926418:AAFcGO-Ab3jyq9VbnyeGgcHD3JHgujKxt3w"
tb = telebot.TeleBot(TOKEN) #create a new Telegram Bot object

#sendMessage
pwd = os.getcwd()
file_name = "report_provi_NCX_FAILED_2020_11_02_to_ncx_check_kawal_ncx.xlsx"
dataFrame = pd.read_excel(pwd + "/ncx_reader/files/" + file_name)
chat_id = -498725127 #coba lazy ncx
print(dataFrame.head())
for df in dataFrame['cmd']:
    tb.send_message(chat_id, df)
tb.remove_webhook()

# # forwardMessage
# tb.forward_message(to_chat_id, from_chat_id, message_id)
