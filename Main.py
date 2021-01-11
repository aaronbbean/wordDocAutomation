import docx
import docx2txt
import os
import time
import datetime
from pathlib import Path
from tkinter import messagebox
import tkinter as tk



class filedocs:
    def __init__(self):
        self.content=[]
        self.date=''
        self.stoplist = ["Main.exe","Records.txt","Main.py"]
        self.reportnumber = ''
        self.incedentlist = ["Accident, Personal Injury","Aggressive Behavior","Alarm","Alcohol/Drugs","Appliance left on","Employee/Public assist","Escort","Found Item","Loss or Disappearance(item)","Elevator","Door Malfunction","Turnstile Malfunction","Power Loss","Medical Emergency","Other","Property Damage","Safety Hazard","Suspicious Activity","Suspicious Person","Trespassing","Unauthorized Entry","Unauthorized Photography","Unlocked Doors/Gates","Vehicle Accident","Water leaks"]
        self.good = True
        self.incedenttype = ''
        self.pathcurrent = ""
        self.pathback = ""

    def errormessage(self,filename):

        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(title=filename, message="{} will not be be moved until an incedent type is given".format(filename))

        pass
    def twotxt(self,filename):

        self.pathcurrent = os.path.dirname(os.path.realpath("__file__"))

        self.pathback = Path(self.pathcurrent).parent

        clearText = docx2txt.process(filename)



        for line in clearText.splitlines():

          if line != '':

            self.content.append(line)


        self.IncedentType(filename)




    def GettingDate(self,filename):
        FILE_IN =  self.pathcurrent+"\\"+filename
        file_date = time.ctime(os.path.getmtime(FILE_IN))
        file_date = datetime.datetime.strptime(file_date, "%a %b %d %H:%M:%S %Y")


        self.date = str(file_date.strftime('%m%d%y'))

    def IncedentType(self,filename):
        i = self.content[8]
        ii = len(i)
        incedent = i[3:ii]

        if incedent == "Choose an item.":
            self.errormessage(filename)
            self.stop(filename)
        else:

            for checking in self.incedentlist:
                if incedent == checking:


                    self.incedenttype = incedent

            self.GettingDate(filename)
            self.addCaseNumber()
            self.filenameing(filename)
            self.end(filename)





    def addCaseNumber(self):
        try:
            caseNum = open("Records.txt","r")

            for i in caseNum.readlines():
                self.reportnumber = str(i)


            caseNum.close()
            NewNumber = open("Records.txt","w")
            self.reportnumber = int(self.reportnumber) + 1
            self.reportnumber = str(self.reportnumber)

            NewNumber.write(self.reportnumber)
            NewNumber.close()
        except:
            self.reportnumber = 101

    def end(self,filename):
        del self.content[:]
        self.reportnumber = ""
        self.incedenttype = ""
        os.remove(filename)
    def stop(self,filename):
        del self.content[:]
        self.reportnumber = ""
        self.incedenttype = ""
        self.stoplist.append(filename)
    def filenameing(self,filename):
        #this will get the date and the report number annd name the folder and file
        doc = docx.Document(filename)
        for p in doc.paragraphs:

            if 'Case#' in p.text:
                inline = p.runs

                for i in range(len(inline)):
                    if 'Case#' in inline[i].text:
                        """this will make the case number"""
                        caseNumber = self.date + self.reportnumber
                        text = inline[i].text.replace('Case#', 'Case#%s' %(caseNumber))
                        inline[i].text = text

        slash = "\\"
        filepath=str(self.pathback)+slash+self.incedenttype+slash+caseNumber
        checkingForFolder = str(self.pathback)+slash+self.incedenttype

        if os.path.exists(checkingForFolder) == True:
            doc.save(filepath+".docx")
        else:
            os.mkdir(checkingForFolder)
            doc.save(filepath+".docx")
    def startfile(self):
        files= os.listdir()


        for file in files:
            if file in self.stoplist:
                pass
            else:
                self.twotxt(file)

if __name__=="__main__":

    start = filedocs()
    start.startfile()
