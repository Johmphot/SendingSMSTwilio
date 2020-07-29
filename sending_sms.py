# Download the helper library from https://www.twilio.com/docs/python/install
import os, csv, time
from datetime import datetime
from email.utils import parsedate_tz
import pandas as pd
import subprocess
from twilio.rest import Client

#Point to virtualenv containing account_sid & auth token

#subprocess.run([source, './twilio.env'])

# Your Account Sid and Auth Token from twilio.com/console
# Testing credentials
account_sid = 'AC8081570defa86830b974745378543f2a'
auth_token = '8f3d26f0bb5cf9d19c8e90be882e1b1b'

path = os.path.dirname(os.path.abspath(__file__))
excel_path = path+"/phonenumber.xlsx"

phone_number_list = []
code_list = []
client_list = []

def addCountryCode(phone_list): #add +66 to Phone Number
    new_phone_list = []
    for pn in phone_list:
        if int(pn[0]) == 0:
            pn = "+66"+pn[1:]
        else:
            pn = "+66"+pn 

        new_phone_list.append(pn) 

    return new_phone_list

def readList(filepath): # read phone number & voucher code from Excel file
    try:
        df = pd.read_excel(filepath, converters={'phone_number':str,'code':str})
        phone_number_list = addCountryCode((df['phone_number']))
        code_list = (df['code'])
        for pn, c in zip(phone_number_list,code_list):
            if len(pn) == 12 and len(c) == 10: #phone with country code = 12 char, code = 10 char
                client_list.append([pn, c])
            else:
                print ("Code / Phone Number Incompleted")

        print (client_list)
    except:
        print("File Error")

def getTextBody():
    filepath = path + '/body.txt'
    try:
       f = open(filepath, 'r')
       text = f.read()
       print (text)
       return text
    except:
        print("File Error")
    
def mergeText(body, code): #add code to the end of text
    return body+code

def testSMS(phone_number, body):
    print ("Send to %s", phone_number)
    print (body)

def sendSMS(phone_number, body, code):
    client = Client(account_sid, auth_token)
    
    message = client.messages.create(
                     from_='AOTAPP',
                     to= phone_number,
                     body = text
                 )

    print(message.sid)

def sendMultipleSMS(body, client_list): # send multiple SMS from Excel file
    client = Client(account_sid, auth_token)

    with open (os.path.dirname(os.path.abspath(__file__))+"/report"+str(datetime.now().date())+".csv", 'w') as f:
        fieldnames = (['Phone Number', 'Code','Message SID', 'Status', 'Timestamp', 'Full Message'] )
        output = csv.writer(f, delimiter=",")
        output.writerow(fieldnames)
            
        for usr in client_list:
            if len(usr) == 2:
                msg_body = mergeText(body, usr[1])
                message = client.messages.create(
                        from_='AOTAPP',
                        to= usr[0],
                        body = msg_body,
                    )

                msg_status = message.fetch().status
                now = datetime.now()
                output.writerow([usr[0], usr[1], message.sid, msg_status, now, msg_body])
                time.sleep(0.2) #delay 0.2 sec per 1 message sent
            else:
                print ("Line Error")
                continue
        f.close()



#### MAIN CODE HERE ####

#initialize using production credentials
account_sid = os.environ['TWILIO_ACCOUNT_SID']
auth_token = os.environ['TWILIO_AUTH_TOKEN']

text = getTextBody()
readList(excel_path)

if input('Proceed to Send SMS? (Y/n)') != 'Y':
    exit()
else :   
    sendMultipleSMS(text, client_list)