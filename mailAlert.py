import smtplib as smtp
import xlrd
import time
import datetime


class readExcel():
    """read excel document containing email and reminder data"""

    def __init__(self, excel):

        """create global class variable"""

        #read excel sheet
        reminder = xlrd.open_workbook(excel)
        self.emailSheet = reminder.sheet_by_index(0)
        self.reminderSheet = reminder.sheet_by_index(1)
        self.dateModeReminder = reminder.datemode

        pass

    def emailList(self):

        #(x,y)

        emailList = []

        for cell in range(self.emailSheet.nrows):
            #loop get each cell value within columns

            #get all values required
            division = self.emailSheet.cell_value(cell,1)
            email = self.emailSheet.cell_value(cell,0)

            #append row
            emailData = [division,email]

            #create row array
            emailList.append(emailData)

            #put mailer here

            pass
        print(emailList)

        pass

    def reminderFinder(self):

        reminderList = []

        for cell in range(self.reminderSheet.nrows):
            #loop get each cell value within columns

            #get date value
            dateValue = self.reminderSheet.cell_value(cell,0)
            #get data as datetime
            date = xlrd.xldate_as_tuple(dateValue, self.dateModeReminder)

            #get all values required
            date = datetime.datetime(date[0],date[1],date[2],10)
            reminder = self.reminderSheet.cell_value(cell,1)
            division = self.reminderSheet.cell_value(cell,2)

            #append row
            data = [date, reminder, division]

            #create row array
            reminderList.append(data)

            pass
        print(reminderList)

        pass

class reminderFinder(object):
    """docstring for reminderFinder."""

    def __init__(self):

        pass

    def loopReminder(self):
        #loop trough excel data

        pass

    def fname(self):


        pass

class mailer(object):
    """docstring for mailer."""

    def __init__(self):

        pass

    def sendMail(self, smtpAddress, smtpPort, emailAdress, password, emailTarget):


        #send access SMTP
        server = smtp.SMTP(smtpAddress, smtpPort)
        #send hello to the server
        server.ehlo()
        #encrypt data
        server.starttls()
        #send hello to the server
        server.ehlo()

        #login email
        server.login(emailAdress, password)

        subject = "this is thest subject"
        body = "this is test body email"

        msg = f'Subject: {subject}\n\n{body}'

        server.sendmail(emailAdress, emailTarget, msg)
        print('sent')

        pass

#===================[run]===================#

#emailAdress = input("email address: ")
#emailTarget = input("email Target: ")
#password = input("password: ")
#
#smtpAddress = "smtp.gmail.com"
#smtpPort = 587
#
#
#mail = mailer().sendMail(smtpAddress, smtpPort, emailAdress, password, emailTarget)

read = readExcel("data.xlsx")
email = read.emailList()
date = read.reminderFinder()
