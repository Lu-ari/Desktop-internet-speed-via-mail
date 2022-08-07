
import webbrowser
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from threading import Timer
import time
from PIL import ImageGrab
import win32com.client as win32
import socket

timer_interval =1
def send_mail():
    import win32com.client as win32
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = '@x.com'# receiver mail adress
    mail.Subject = socket.gethostname()+' computer host speed test'
    mail.Body = '…:speed test'
    mail.HTMLBody = ' <h2>speed test</h2>' #this field is optional

    # To attach a file to the email (optional):
    attachment  = ""# write here directory
    mail.Attachments.Add(attachment)

    mail.Send()

def browser_open():
    webbrowser.open(https://fast.com/', new=2)

def browser_close():
    os.system("taskkill /im chrome.exe /f")

def getDesktopimg():
    im = ImageGrab.grab()
    im.save("sample.png")# write here directory to save png file
def delayrun():
    print("test  test")
    t = Timer(timer_interval,delayrun())
    t.start()

while True:
    browser_open()
    time.sleep(10)
    getDesktopimg()
    browser_close()
    send_mail()
    


