import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
from selenium.common.exceptions import *
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common import exceptions
import random
import datetime
from datetime import datetime
from selenium import webdriver
import pandas as pd
import os.path
import sys
import pandas as pd
import time
import glob
import os
import shutil
import getpass
global xuser
import ntpath
import datetime
from openpyxl import *
import sys
import time
from tkinter import *
from PIL import ImageTk, Image
import gzip
import shutil
import tkinter as tk
import random
import win32com.client,datetime
import win32com.client as win32
global xuser
import getpass
import time
from zipfile import ZipFile
from datetime import datetime
import os
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as By
from selenium.common.exceptions import NoAlertPresentException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import *
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import datetime as date
from datetime import timedelta
import sys
import pandas as pd
import logging
import pandas as pd
import getpass
import numpy as np
import tarfile
import sys
import time
from win32com.client import Dispatch
from tkinter import *
from PIL import ImageTk, Image
import gzip
import shutil
import tkinter as tk
import random
import win32com.client,datetime
import win32com.client as win32
global xuser
import getpass
import time
from zipfile import ZipFile
from datetime import datetime
import os
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as By
from selenium.common.exceptions import NoAlertPresentException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import *
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import datetime as date
from datetime import timedelta
import sys
import pandas as pd
import logging
import pandas as pd
import getpass
import numpy as np
from selenium import webdriver
from selenium.webdriver.support.ui import Select

root = tk.Tk()

logging.basicConfig(filename="C:\\Users\\"+getpass.getuser()+"\\Documents\\ATO.log",format='%(asctime)s: %(levelname)s: %(message)s',
                    datefmt='%m/%d/%y %I:%M:%S %p',level=logging.DEBUG)

username = getpass.getuser()
logging.debug("BUlk movie cancellation invoked")
root.title("Movie cancellation tool")

try:
    os.mkdir("C:\\Users\\" +getpass.getuser()+ "\\Desktop\\Movie-cancel-automation")
except:
    pass

def extraction():
    outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
    mapi = Dispatch("Outlook.Application").GetNamespace("MAPI")

    #inbox = mapi.Folders("Archives").Folders("Inbox").Folders("Movies")
    inbox = mapi.Folders("Inbox")
    messages = inbox.Items
    mess = []
    rctime = []
    for item in messages:
        received_time = str(item.ReceivedTime)
        print(item.Subject)
        mess.append(item.Subject)
        rctime.append(received_time)

        mainframe = pd.DataFrame(list(zip(mess,rctime)))
        import datetime
        mainframe.columns = ["Extraction Subject","Received Time"]
        print(mainframe)
        val_date = datetime.date.today()
        username = getpass.getuser()
        mainframe.to_excel("C:\\Users\\" + username + "\\Desktop\\Movie-cancel-automation\\Subject_Extract" + str(val_date) + ".xlsx",
            index=False)
    print("done")

def scrapee():
    path = "C:\\Users\\" + getpass.getuser() + "\\Desktop\\Movie-cancel-automation"
    for files in os.listdir(path):
        if files.startswith('Email_extract'):
            final = os.path.join(path, files)
    df = pd.read_excel(final)
    print(df)

    selenium = webdriver.Chrome('C:\Program Files (x86)\SeleniumWrapper\chromedriver.exe')
    wait = WebDriverWait(selenium, 60, poll_frequency=1,
                         ignored_exceptions=[NoSuchElementException, ElementNotSelectableException,
                                             ElementNotVisibleException])
    selenium.get("https://midway-auth.amazon.com/login?next=%2FSSO%2Fredirect%3Fresponse_type%3Did_token%26client_id%3Dgrass-eu.aka.amazon.com%253A443%26redirect_uri%3Dhttps%253A%252F%252Fgrass-eu.aka.amazon.com%253A443%252F%253Fprod%253Don%2526realm%253DDETECT%2526id%253D408-9465686-7914769%26scope%3Dopenid%26nonce%3Deee7de1178fdba384ae251a2a55616b509f5b978d57a7d03d8922df49c39d0b7%26sentry_handler_version%3DTomcatSSOFilter-1.1-1%26uid%3D63dd80d6-52f2-47cd-94ef-df35a2175240%26for_sentry%3Dtrue%26sentry_verification%3DeyJ4NXUiOiJodHRwOlwvXC9zZW50cnktcGtpLmFtYXpvbi5jb21cL3B1YmxpY2tleVwvNTA5NDI4MiIsInR5cCI6IkpXUyIsImFsZyI6IlBTMjU2In0.eyJ0eXBlIjoiU2VudHJ5VmVyaWZpY2F0aW9uIiwiZXhwIjoxNTY5NTAwNzc4fQ.Z4x6_YTCT-sPeX13R0GddaZ248n8aXhRIohD1dndB_lkdX7R_GLXY-62t9Mfj7a16TdG25YVhJBIB7ruK9rLiy3f94k2ZTtBpgTd9M7uS670r0CuD3jpmiBulJC-oFY7Oh8Zz2w1WgdGBeHbVil7a7DHw0uQpaIhHBwDEvFSjJEdqM5rpcMT6TVCQ48X5WGDIpG_Lku8a_i1mwsY-wOIaEEub33q0CSF35Xc4KDh123G4h2KY1H0H3XzJjiB6BCI35zv4RWU8lRyKki3AFbWB58hc5WoxXT66LYUIFK9PX1fOo12iUNIaegrIyf1MZmmTlu4PSU_L2v95wwy8J8Mtg%26kerberos%3DeyJ4NXUiOiJodHRwOlwvXC9zZW50cnktcGtpLmFtYXpvbi5jb21cL3B1YmxpY2tleVwvNTA5NDI4MiIsInR5cCI6IkpXUyIsImFsZyI6IlBTMjU2In0.eyJzdWJqZWN0IjoiZ29rdWtAQU5ULkFNQVpPTi5DT00iLCJleHAiOjE1Njk1MDAxOTR9.kFDxFPVClCBudcSC8Qcaia876xWnfPqLuhMQ33-t7hWbA-JcTlLBukC4akAiQ6ERYgdhimfMPLTbWBks-eeG97ILwqpGFyBVJwthanWupdanBxcV7Hy07pxjHk_ZZSKKbLWjHMlnZJ0KocxTBcKQq60yR1hGHwjVTVJ525PsQ2b0kzakCQqIRSDYAeQo2AqDs1HOH0pk9vjhzusxgzGnH3t6JgyObBLZMJTdchSme2Ffy-I9lCiEqUli-8iuJjQSqaRiSbxaLkr92Ra_uAXQZ87HudTJQ5D2zsjomPGhX8OnhGh8TZuTLGzgZSNN38pCHhxP6C1pmv62FEz3fbYoIQ&noauth=false&require_digital_identity=false")


    x = entry.get()
    y = out.get()
    time.sleep(5)
    selenium.find_element_by_id("user_name_field").send_keys(x)
    time.sleep(1)
    selenium.find_element_by_id("user_name_btn").click()
    time.sleep(2)
    selenium.find_element_by_id("password_field").send_keys(y)
    time.sleep(1)
    selenium.find_element_by_id("password_btn").click()
    time.sleep(15)

    for i in range(0, len(df)):

        c = df.iloc[i, 5]
        d = df.iloc[i, 6]
        print(df.iloc[i, 5])
        print(df.iloc[i, 6])
        selenium.get("https://dashboard.beta.hfctl.payments.a2z.com/hfcuser/cancel")

        wait.until(EC.visibility_of_element_located((By.XPATH, "//*[@id='venue_id']")))
        selenium.find_element_by_xpath("//*[@id='venue_id']").send_keys(str(c))

        wait.until(EC.visibility_of_element_located((By.XPATH, "//*[@id='session_id']")))
        selenium.find_element_by_xpath("//*[@id='session_id']").send_keys(str(d))
        try:
            wait.until(EC.visibility_of_element_located((By.XPATH, "/html/body/form/div[3]/div/input")))
            selenium.find_element_by_xpath("/html/body/form/div[3]/div/input").click()
            time.sleep(5)
            wait.until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[3]/button")))
            selenium.find_element_by_xpath("/html/body/div[3]/button").click()
            obj = selenium.switch_to.alert
            msg = obj.text
            print(msg)
            obj.accept()
        except:
            obj = selenium.switch_to.alert
            msg = obj.text
            print(msg)
            obj.accept()

def sendmail():
    mapi = Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = mapi.Folders("Archives").Folders("Inbox").Folders("Movies")
    print("Inbox name is:", inbox.Name)
    messages = inbox.Items
    for item in messages:
        reply = item.Reply()
        reply.Body = "The Shows have been Cancelled"
        reply.Send()
        print("Shows have been cancelled")




ct = [random.randrange(256) for x in range(3)]
canv = Canvas(root, width=800, height=500, bg='orange')
canv.grid()

#brightness = int(round(0.299 * ct[0] + 0.587 * ct[1] + 0.114 * ct[2]))
#ct_hex = "%02x%02x%02x" % tuple(ct)
#bg_colour = '#' + "".join(ct_hex)
o = tk.Button(root,text='Email Extraction',fg='White',bg='navyblue',command = extraction)
o.place(x=20, y=30 + 2 * 30, width=180, height=25)
p = tk.Button(root,text='Cancel Orders',fg='White',bg='navyblue',command = scrapee)
p.place(x=220, y=30 + 2 * 30, width=180, height=25)
lnnn = tk.Button(root,text='Email Reply',fg='White',bg='navyblue',command = sendmail)
lnnn.place(x=420, y=30 + 2 * 30, width=180, height=25)

Label(root, text="Midway username").place(x=280, y=30 + 5 * 30, width=180, height=25)
entry = Entry(root)
entry.place(x=420, y=30 + 5 * 30, width=180, height=25)
Label(root, text="Midway Password").place(x=280, y=30 + 6 * 30, width=180, height=25)
out = Entry(root, show="*")
out.place(x=420, y=30 + 6 * 30, width=180, height=25)
root.mainloop()