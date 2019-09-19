import os
import openpyxl as o
import random
import datetime
import psutil

# main functions

# to search for name by giving his national ID


def searchForNID(NID):
    global rowNumber

    if NIDsColumn == "":
        saveLogs("===> National IDs column wasn't initiated at " +
                 str(datetime.datetime.now()))
        return "wrong with initiation"

    NID = str(NID)

    rowNumber = '2'
    saveLogs("==> Search for National ID "+str(NID))

    while sheet[codesColumn + rowNumber].value != None:

        cellValue = sheet[NIDsColumn + rowNumber].value

        if cellValue == None:
            rowNumber = int(rowNumber)
            rowNumber += 1
            rowNumber = str(rowNumber)

        else:
            cellValue = str(cellValue)

            if cellValue != NID:
                rowNumber = int(rowNumber)
                rowNumber += 1
                rowNumber = str(rowNumber)

            elif cellValue == NID:
                return returningData(rowNumber)

    saveLogs("===> Searching for the National ID didn't get to a result at " +
             str(datetime.datetime.now()))
    return None


# to search for name by giving his phone


def searchForPhone(phone):
    global rowNumber

    if phonesColumn == "":
        saveLogs("===> Phones column wasn't initiated at " +
                 str(datetime.datetime.now()))
        return "wrong with initiation"

    phone = str(phone)

    rowNumber = '2'
    saveLogs("==> Search for PHONE "+str(phone))

    while sheet[codesColumn + rowNumber].value != None:

        cellValue = sheet[phonesColumn + rowNumber].value

        if cellValue == None:
            rowNumber = int(rowNumber)
            rowNumber += 1
            rowNumber = str(rowNumber)

        else:
            cellValue = str(cellValue)

            if cellValue != phone:
                rowNumber = int(rowNumber)
                rowNumber += 1
                rowNumber = str(rowNumber)

            elif cellValue == phone:
                return returningData(rowNumber)

    saveLogs("===> Searching for the phone didn't get to a result at " +
             str(datetime.datetime.now()))
    return None


# to search for name by giving his code

def searchForCode(code):
    global rowNumber

    if codesColumn == "":
        saveLogs("===> Codes column wasn't initiated at " +
                 str(datetime.datetime.now()))
        return "wrong with initiation"

    code = str(code)

    rowNumber = '2'
    saveLogs("==> Search for ID "+code)

    while sheet[codesColumn + rowNumber].value != None:

        cellValue = sheet[codesColumn + rowNumber].value

        if cellValue == None:
            rowNumber = int(rowNumber)
            rowNumber += 1
            rowNumber = str(rowNumber)

        else:
            cellValue = str(cellValue)

            if cellValue != code:
                rowNumber = int(rowNumber)
                rowNumber += 1
                rowNumber = str(rowNumber)

            elif cellValue == code:
                return returningData(rowNumber)

    saveLogs("===> Searching for the code didn't get to a result at " +
             str(datetime.datetime.now()))
    return None


# to add that the guy attended today

def add_by_day():
    global labtopDeviceNumber

    placeOfAttendeeToday = todayColumn + str(rowNumber)

    sheet[placeOfAttendeeToday] = "1"

    saveLogs("===> Saved today attendance for row number " +
             str(rowNumber) + " at " + str(datetime.datetime.now()))


# helping function

# important variabless
labtopDeviceNumber = None
today = None
wb = None
sheet = None
workshopsColumns = []
namesColumn = ""
codesColumn = ""
phonesColumn = ""
NIDsColumn = ""
todayColumn = ""
rowNumber = None
fileName = ""


# function that return data
def returningData(rowNumber):

    name = sheet[namesColumn + rowNumber].value
    code = sheet[codesColumn + rowNumber].value
    workshop1 = sheet[workshopsColumns[0] + rowNumber].value
    workshop2 = sheet[workshopsColumns[1] + rowNumber].value
    workshop3 = sheet[workshopsColumns[2] + rowNumber].value

    returns = [name, code, workshop1, workshop2, workshop3]
    for i in range(len(returns)):
        if returns[i] == None:
            returns[i] = "Cell Has No Value"

    return returns


# function that saves variables of setting
def changeSettings(labtopNumber, filePath, attendeesInformation, workshops, isOldFile):
    global labtopDeviceNumber, sheet, wb, workshopsColumns, namesColumn, codesColumn, NIDsColumn, todayColumn, fileName, phonesColumn

    saveLogs(
        "==> Changing Settings at " + str(datetime.datetime.now()))

    labtopDeviceNumber = int(labtopNumber)
    filePathList = filePath.split('/')
    fileName = filePathList[-1].lower()
    fileName = fileName.replace('.xlsx', '')

    try:
        # if file was open in the program then save it
        excelFilePath = os.getcwd() + "\\" + fileName + " - " + \
            str(labtopDeviceNumber) + '.xlsx'
        if checkingIfFileOpen(excelFilePath):
            saveLogs("==> File can't be saved as it's open by Microsoft EXCEl at " +
                     str(datetime.datetime.now()))
            return False
        else:
            wb.save(fileName + " - " + str(labtopDeviceNumber) + '.xlsx')
            saveLogs(
                "==> Saving file worked on at " + str(datetime.datetime.now()))
    except:
        pass

    try:
        # if file was already made in same folder of the app then open it
        # checkingIfFileOpen of that file was checked in the try before

        if not os.path.exists(fileName + " - " + str(labtopDeviceNumber) + '.xlsx'):
            raise ValueError('A very specific bad thing happened')

        if isOldFile == False:
            saveLogs(
                "==> Tried to open file " + fileName + " - " +
                str(labtopDeviceNumber) + '.xlsx' + " and asked for permission at " + str(datetime.datetime.now()))
            return ("workWithOldFile")

        elif isOldFile == "overwrite":
            raise ValueError('A very specific bad thing happened')

        else:
            wb = o.load_workbook(fileName + " - " +
                                 str(labtopDeviceNumber) + '.xlsx')
            sheet = wb.active
            saveLogs(
                "==> Opened file from the same folder at " + str(datetime.datetime.now()))

    except:
        try:
            # else if there was no file in the same folder work on a new one
            if checkingIfFileOpen(filePath):
                saveLogs("==> File can't be apened as it's open by Microsoft EXCEl at " +
                         str(datetime.datetime.now()))
                return False
            else:
                wb = o.load_workbook(filePath)
                sheet = wb.active
                saveLogs(
                    "==> Opened file from a path at " + str(datetime.datetime.now()))
                saveLogs(
                    "=> File path is  " + filePath)

        except FileNotFoundError:
            saveLogs("==> Wrong file path was entered at " +
                     str(datetime.datetime.now()))
            saveLogs(
                "=> File path is  " + filePath)
            return False  # means that file is not available

    workshopsColumns = workshops

    namesColumn = attendeesInformation[0]
    codesColumn = attendeesInformation[1]
    phonesColumn = attendeesInformation[2]
    NIDsColumn = attendeesInformation[3]
    todayColumn = attendeesInformation[4]

    savingConfigurations(labtopDeviceNumber, filePath,
                         attendeesInformation, workshopsColumns)

    return True  # means that file available


# attendeesInformation[0] == attendees names column
# attendeesInformation[1] == attendees IDs column
# attendeesInformation[2] == attendees phones column
# attendeesInformation[3] == attendees national IDs column
# attendeesInformation[4] == today column


# function to save settings

def savingConfigurations(labtopDeviceNumber, filePath, attendeesInformation, workshopsColumns):
    file = open(r'app.config', 'w')
    file.write(str(labtopDeviceNumber) + '\n' + filePath + '\n' + attendeesInformation[0]+'\n' + attendeesInformation[1]+'\n' + attendeesInformation[2] + '\n' + attendeesInformation[3] + "\n" + attendeesInformation[4] + "\n" +
               workshopsColumns[0] + '\n' + workshopsColumns[1] + '\n' +
               workshopsColumns[2])
    file.close()
    saveLogs(
        "==> Setting was saved")
    saveLogs("labtop Number= " + str(labtopDeviceNumber) + '\n' + "Names Col= " + attendeesInformation[0]+'\n' + "Codes Col= " + attendeesInformation[1]+'\n' + "Phones Col= "+attendeesInformation[2] + "\n" + "National IDs Col= "+attendeesInformation[3] + "Today's Col= "+attendeesInformation[4] + "\n" +
             "Workshops Cols = " + workshopsColumns[0] + ' - ' + workshopsColumns[1] + ' - ' +
             workshopsColumns[2])


# function to get settings when needed


def readConfiguration():
    global labtopDeviceNumber, today, workshopsColumns

    file = open(r'app.config', 'r')
    setting = file.read()
    file.close()
    setting = setting.split("\n")
    saveLogs("==> Reading old Configuration at " +
             str(datetime.datetime.now()))
    return setting


# function that save new codes in codes file


def saveLogs(logLine):

    file = open(r'logFile.log', 'a')
    file.write(logLine + '\n')
    file.close()


def closingExcelFile():

    if fileName != "":

        excelFilePath = os.getcwd() + "\\" + fileName + " - " + \
            str(labtopDeviceNumber) + '.xlsx'

        if checkingIfFileOpen(excelFilePath):
            saveLogs("==> File can't be saved as it's open by Microsoft EXCEl at " +
                     str(datetime.datetime.now()))

            return False

        else:
            wb.save(fileName + " - " + str(labtopDeviceNumber) + '.xlsx')
            wb.close()
            saveLogs("======> Sheet was saved and closed at " +
                     str(datetime.datetime.now()))
            return True

    else:

        saveLogs("======> There is no sheet to be saved or closed at " +
                 str(datetime.datetime.now()))
        return True


# psutil.pids() ==> all Processes IDs
# p = psutil.Process(PID) ==> to get data of a specific process "PID is the process ID"
# p.name() ==> to get name of process
# p.open_files() ==> to get files opened by that process


def checkingIfFileOpen(filePath):
    excelFilePath = filePath.replace("/", "\\")
    for pid in psutil.pids():
        try:
            if psutil.Process(pid).name() == "EXCEL.EXE":
                for file in psutil.Process(pid).open_files():
                    if excelFilePath == file.path:
                        return True
        except:
            pass

    return False


# at emergency states we will make sure to save the file was working on

def forceSavingFile():
    try:
        wb.save(fileName + " - " + str(labtopDeviceNumber) + '.xlsx')
        wb.close()
    except:
        pass
