import win32com.client as client
import datefinder
import datetime


address = 'thomasboyz@live.com'
outlook = client.Dispatch('Outlook.Application')
namespace = outlook.GetNameSpace('MAPI')
account = namespace.Folders[address]
inbox = account.Folders('Inbox')

keywords = ['sick', 'appointment', 'vacation', 'abscent', 'Out of Office']
emails = []
dates = []
Calendar = []


class Employee:
    email = ""
    name = ""
    date = ""





# for word in keywords:
#    emails += [message for message in inbox.Items if word in message.Body.lower()]

test = [
    message for message in inbox.Items if message.SenderEmailAddress.endswith('hotmail.com')]


for message in test:
    dates += datefinder.find_dates(message.Body)

for date in dates:
    date = date.strftime("%b/%d/%Y")
    print(date)


'''
with open("Out of Office.txt", "w") as file
    for message in test_mail:
        file.write(message.SenderEmailAddress + " --- ")
        for date in dates:
            file.write(date.strftime("%d-%b-%Y") + '\n')
'''

print("END")
