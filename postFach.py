import win32com.client
import os
from datetime import datetime, timedelta
import xlwings
import pandas as pd
import tkinter as tk
import re


intSpaltePO = 1
intSpaltePath = 2
intSpalteSubject = 3
intSpalteSent = 4
intSpalteReceived = 5
intSpalteCategory = 6
intSpalteBody = 10
intZeile = 5

def Inbox_Importieren() :
    
    outlook = win32com.client.Dispatch('outlook.application')
    nameSpace  = outlook.GetNamespace("MAPI")


    wb = xlwings.Book("Path")
    
    ws = wb.sheets[0]
    
    account = nameSpace.CreateRecipient("yourMail")
    olFolderInbox = 6
    ordner = nameSpace.GetSharedDefaultFolder(account, olFolderInbox)
    #inbox = account.Folders.Item("Inbox")
    intZeile = 5
    count = 0
    zeilenhöhe = ws.range("A1").row_height
    for item in ordner.Items:
        if item.Class == 43: 
            count += 1
            ws.cells(intZeile, intSpaltePath).value = "\\Inbox\\"
            ws.cells(intZeile, intSpalteSubject).value = item.Subject
            ws.cells(intZeile, intSpalteSent).value = item.SentOn
            ws.cells(intZeile, intSpalteReceived).value = item.ReceivedTime
            ws.cells(intZeile, intSpalteCategory).value = item.Categories
            body_text = item.Body
            body_text = re.sub(r'[^\w\s]', '', body_text)
            ws.cells(intZeile, intSpalteBody).value = body_text
       
            strSubject = ws.cells(intZeile, intSpalteSubject).value
            strBody = ws.cells(intZeile, intSpalteBody).value
        
            GetNum = re.findall(r'\d+', strSubject)
            if not GetNum:
                GetNum = re.findall(r'\d+', strBody)
            if GetNum:
                for num in GetNum:

                    if num.startswith("30") or num.startswith("0030"):
                        if num.startswith("00"):
                            ws.cells(intZeile, intSpaltePO).value = num[:10]
                        else: 
                            ws.cells(intZeile, intSpaltePO).value = num[:8]
                        break

                    elif num.startswith("33") or num.startswith("0033"):
                        if num.startswith("00"):
                            ws.cells(intZeile, intSpaltePO).value = num[:12]
                        else: 
                            ws.cells(intZeile, intSpaltePO).value = num[:10]
                        break

                    elif num.startswith("45") or num.startswith("0045"):
                        if num.startswith("00"):
                            ws.cells(intZeile, intSpaltePO).value = num[:12]
                        else:
                            ws.cells(intZeile, intSpaltePO).value = num[:10]
        
            intZeile += 1
        
    ws.range("F2").value = count
    
    ws.range("A5").expand().row_height = zeilenhöhe
   

Inbox_Importieren()


root = tk.Tk()
label = tk.Label(root, text="Script execution complete.")
label.pack()
root.mainloop()
   
