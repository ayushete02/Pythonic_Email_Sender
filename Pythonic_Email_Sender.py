''' Before starting program please Allow less secure apps: ON link:  https://myaccount.google.com/lesssecureapps?pli=1&rapt=AEjHL4OgLOVId9DxSpj6GTOkPKSElDDzpCBJuwKknnxCIBvzyXLNvrztWXWUeoykd-BWFyLvFuqXMk92ERdA1FK7Sw5gejXWkg
    Give permission to your Google to send Email By Simply clicking on the new arrived Google mail''' # ==> "Yes,this is me" 
# lets start
# Made by : Ayush Shete

import pandas as pd
import smtplib
import datetime
from plyer import notification


# Enter your personal Login detail
GMAIL_ID = '<Enter Senders Email Address>'   #Enter your Email Addrress
GMAIL_PSWD = '<Enter Your Password>'         #Enter Your Password 


def Send_Email(To, Subject, Message):
    try:
        print("Please wait Email is Sending...")
        Email = smtplib.SMTP('smtp.gmail.com', 587)
        Email.starttls()
        Email.login(GMAIL_ID, GMAIL_PSWD)
        Email.sendmail(GMAIL_ID, To, f"Subject: {Subject}\n\n{Message}")
        print("\n**********************************************\n")
        print("\tmessage send successfully")
        print("\n**********************************************\n")
        notification.notify(
            title="Message Send Successfully",
            message=f"Message Send to {To}",
            app_icon="",
            timeout=3
        )
        History = open("History.txt", "a")
        Name_in_Excel = item['Name']
        CurrentTime = datetime.datetime.now()
        History.write(f"Name: {Name_in_Excel}\t\tTime:{CurrentTime}\nEmail Send Successfully to: {To}\tfrom: {GMAIL_ID}\nSubject: {Subject}\n\n\n")
        History.close()
        Email.quit()
    except:
        print("\n\t*****************************************\n")
        print("\tXXX Email not Sent Invalid Email Id XXX\n\t  Please Check Your Internet Conection\n\t      please Check Entered Details")
        print("\n\t*****************************************\n")
        notification.notify(
            title="Error Message not Send",
            message=f"Please Check Your Presonal Data",
            app_icon="",
            timeout=3
        )
        History = open("History.txt", "a")
        Name_in_Excel = item['Name']
        CurrentTime = datetime.datetime.now()
        History.write(f"Name: {Name_in_Excel}\t\tTime:{CurrentTime}\nEmail not Send to: {To}\tfrom: {GMAIL_ID}\nSubject: {Subject}\n\n\n")
        History.close()


df = pd.read_excel('Details.xlsx')

# Today / Current year initilizing
TodayDate = datetime.datetime.now().strftime("%d-%m")
CurrentYear = datetime.datetime.now().strftime("%Y")
print("Data Processing...")


for index, item in df.iterrows():  # reading line by line in Excel
    Date_in_Excel = item['Date'].strftime("%d-%m")
    Year_in_Excel = str(item['Wishing_Year'])

    if Date_in_Excel == TodayDate and CurrentYear == Year_in_Excel:
        To_in_Excel = item['Receiver_Email_Id']
        Subject_in_Excel = item['Subject']
        Message_in_Excel = item['Message']
        Send_Email(To_in_Excel, Subject_in_Excel, Message_in_Excel)
        
    # Made by : Ayush Shete