from googleapiclient.discovery import build
from google.oauth2 import service_account
from datetime import datetime, timedelta
import cv2
from PyQt5.QtWidgets import QApplication, QLabel, QMainWindow, QWidget, QPushButton, QLineEdit, QFormLayout, QComboBox
from googleapiclient.http import MediaFileUpload
import gspread
import os

creds = service_account.Credentials.from_service_account_file('./credential.json',
                                                              scopes=['https://www.googleapis.com/auth/drive','https://www.googleapis.com/auth/spreadsheets'])
drive_service = build('drive', 'v3', credentials=creds)

folderdests = [""]
folderdest = ""
imagedests = [""]
imagedest = ""
today = datetime.now().date()
datestart = today - timedelta(days = today.weekday())

class FirstWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.w = None
        window = QWidget()
        layout = QFormLayout()
        self.box = QComboBox()
        self.box.addItem("Office 1")
        self.box.addItem("Office 2")
        self.button = QPushButton("Accept")
        label=QLabel("Choose the office that you work at")
        layout.addRow(label)
        self.button.clicked.connect(self.updateOffice)
        layout.addRow("Office : ", self.box)
        layout.addRow(self.button)
        window.setLayout(layout)
        self.setCentralWidget(window)

    def updateOffice(self):
        global folderdest
        global folderdests
        global imagedests
        global imagedest
        office = str(self.box.currentText())
        match office:
            case "Office 1":
                folderdest = folderdests[0]
                imagedest = imagedests[0]
            case "Office 2":
                folderdest = folderdests[1]
                imagedest = imagedests[1]
        self.w = SecondWindow()
        self.hide()
        self.w.show()
        
class SecondWindow(QWidget):
    def __init__(self):
        super().__init__()
        layout = QFormLayout()
        self.linefirst = QLineEdit()
        self.linelast = QLineEdit()
        self.button = QPushButton("Take Picture")
        self.button.clicked.connect(self.getImage)        
        layout.addRow("First Name : ", self.linefirst)
        layout.addRow("Last Name : ", self.linelast)
        layout.addRow(self.button)
        
        gc = gspread.authorize(creds)
        global folderdest
        
        try:
            sh = gc.open(datestart.strftime("%m/%d/%Y"),folderdest)           
        except:
            sh = gc.create(datestart.strftime("%m/%d/%Y"),folderdest)
            worksheet = sh.get_worksheet(0)
            worksheet.update_title(datestart.strftime("%m/%d/%Y"))
            sh.add_worksheet((datestart + timedelta(1)).strftime("%m/%d/%Y"),100,10,1)
            sh.add_worksheet((datestart + timedelta(2)).strftime("%m/%d/%Y"),100,10,2)
            sh.add_worksheet((datestart + timedelta(3)).strftime("%m/%d/%Y"),100,10,3)
            sh.add_worksheet((datestart + timedelta(4)).strftime("%m/%d/%Y"),100,10,4)
            sh.add_worksheet((datestart + timedelta(5)).strftime("%m/%d/%Y"),100,10,5)
            sh.add_worksheet((datestart + timedelta(6)).strftime("%m/%d/%Y"),100,10,6)
            
        self.worksheet = sh.get_worksheet(today.weekday())

        self.setLayout(layout)

          
    def getImage(self):
        
        global imagedest
        cam = cv2.VideoCapture(0,cv2.CAP_DSHOW)
        result, image = cam.read()
        cam.release()

        image = cv2.cvtColor(image,cv2.COLOR_BGR2GRAY)

        firstname = self.linefirst.text()
        lastname = self.linelast.text()
        imagename = firstname+" "+lastname+" "+datetime.now().strftime("%m-%d-%Y %H-%M-%S")+".jpg"
        cv2.imwrite(imagename,image)
        
        allfirst = [item for item in self.worksheet.col_values(1) if item]
        alllast = [item for item in self.worksheet.col_values(2) if item]
        included = False
        x = 1
        for i in range(0,len(allfirst)):
            if allfirst[i] == firstname and alllast[i] == lastname:
                included = True
                x = i+1
        if included:
            times = [item for item in self.worksheet.row_values(x) if item]
            self.worksheet.update_cell(x,len(times)+1,datetime.now().strftime("%H:%M:%S"))
        else:
            self.worksheet.update_cell(len(allfirst)+1,1,firstname)
            self.worksheet.update_cell(len(allfirst)+1,2,lastname)
            self.worksheet.update_cell(len(allfirst)+1,3,datetime.now().strftime("%H:%M:%S"))
        
        file_metadata = {
                "name" : imagename,
                "mimeType" : "image/jpeg",
                "parents" : [imagedest]}
        media = MediaFileUpload("./"+imagename,mimetype="image/jpeg")
        f = drive_service.files().create(body = file_metadata, media_body = media,
                                             supportsAllDrives=True).execute()
        del media
        try:
            os.remove(imagename)
        except Exception as e:
            print(e)
         
        self.linefirst.clear()
        self.linelast.clear()


app = QApplication([])
app.setStyle("Fusion")
w = FirstWindow()
w.show()
app.exec()
