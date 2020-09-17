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
        self.divisionSheet = reminder.sheet_by_index(2)
        #read excel files as datemode format
        self.dateModeReminder = reminder.datemode

        pass

    def emailFinder(self):

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

            pass
        print(emailList)

        return(emailList)

    def reminderFinder(self):

        reminderList = []

        for cell in range(self.reminderSheet.nrows):
            #loop get each cell value within columns

            #get date value
            dateValue = self.reminderSheet.cell_value(cell,0)
            #get data as datetime
            date = xlrd.xldate_as_tuple(dateValue, self.dateModeReminder)

            #get all values required
            date = datetime.date(date[0],date[1],date[2])
            reminder = self.reminderSheet.cell_value(cell,1)
            division = self.reminderSheet.cell_value(cell,2)

            #append row
            data = [date, reminder, division]

            #create row array
            reminderList.append(data)

            pass
        print(reminderList)

        return(reminderList)

    def listDivision(self):

        divisionList = []

        for cell in range(self.divisionSheet.nrows):
            #loop get each cell value within columns

            #get all values required
            division = [self.divisionSheet.cell_value(cell,0)]

            #create row array
            divisionList.append(division)

            pass

        print(divisionList)
        return(divisionList)

class reminderLoop(object):
    """docstring for reminderFinder."""

    def __init__(self, reminderList, emailList):

        #create gloal variable
        #read list
        self.reminderList = reminderList
        self.emailList = emailList

        pass

    def loopReminder(self, dayNotification, wichDivision):
        #loop through reminder list

        for reminder in self.reminderList:

            #broadcast acording to teh division
            if reminder[2] == wichDivision:

                #broadcast if near due date
                if datetime.date.today() + datetime.timedelta(days = dayNotification) >= reminder[0]:

                    #after get due date now what
                    #broadcast acording to division

                    print(reminder[2],reminder[1:3])

                    pass

                pass

            pass

        pass

    def LoopDivision(self, wichDivision):

        #get mail acording to division
        for mail in self.emailList:

            #get email from list
            if mail[0] == wichDivision:

                print(mail[1])

                pass

            pass

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

class emailBlasterBuilder():
    """Build email Blaster Bot"""

    def __init__(self):

        pass

    def divisionTarget(self, excel):

        #get excel data
        getExcelData = readExcel(excel)
        #get email sheet
        getEmailData = getExcelData.emailFinder()
        #get reminder sheet
        getReminderData = getExcelData.reminderFinder()

        #find reminder and loop division
        getDivision = reminderLoop(getEmailData, getReminderData)

        for division in variable:
            pass

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
email = read.emailFinder()
date = read.reminderFinder()
division = read.listDivision()
find = reminderLoop(date, email)
reminder = find.loopReminder(3,"division 2")
mail = find.LoopDivision("division 3")
