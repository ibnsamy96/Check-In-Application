from appjar import gui
import backend
import datetime


app=gui("Check-in Application")

backend.saveLogs("<====================================<>====================================>")
backend.saveLogs("======> This session starts at " + str(datetime.datetime.now()))


###################################################################################################################### these go in the sub windows


###################### function of all buttons of sub windows ( launch , launch setting , launch exit , launch adding )

def launch(win):
    
    if win == 'exit' :
        app.showSubWindow('exit')
    
    elif win == 'Setting' :
        app.showSubWindow('Setting')


setting_saved = False

def launch_setting (win) :
    global setting_saved
    
    if win == 'importConfigurations' and setting_saved == False :
        try:
            setting = backend.readConfiguration()
            app.setEntry("labtopDeviceNumber", str(setting[0]), callFunction=False)
            app.setEntry("filePath", setting[1], callFunction=False)
            app.setEntry("namesColumn", setting[2], callFunction=False)
            app.setEntry("codesColumn", setting[3], callFunction=False)
            app.setEntry("phonesColumn", setting[4], callFunction=False)
            app.setEntry("NIDsColumn", setting[5], callFunction=False)
            app.setEntry("todayColumn", setting[6], callFunction=False)
            app.setEntry("workshopColumn1", setting[7], callFunction=False)
            app.setEntry("workshopColumn2", setting[8], callFunction=False)
            app.setEntry("workshopColumn3", setting[9], callFunction=False)
        except:
            app.setLabel('warning',"No configuration in 'app.config' file")
        

    elif win == 'save_setting' and setting_saved == False :
        if app.getEntry("labtopDeviceNumber") != 0 and app.getEntry("filePath") != "" :
            fileExtension=app.getEntry("filePath") [-4:-1]+app.getEntry("filePath")[-1].lower()
            if fileExtension== "xlsx":
                workshopsColumns=[app.getEntry("workshopColumn1").upper(),app.getEntry("workshopColumn2").upper(),app.getEntry("workshopColumn3").upper()]
                if workshopsColumns[0] != "" and workshopsColumns[1] != "" and workshopsColumns[2] != "" :
                    if workshopsColumns[0].isalpha() and workshopsColumns[1].isalpha() and workshopsColumns[2].isalpha() :
                        if len(workshopsColumns[0]) == 1 and len(workshopsColumns[1]) == 1 and len(workshopsColumns[2]) == 1 :
                            attendeesInformation=[app.getEntry("namesColumn").upper(),app.getEntry("codesColumn").upper(),app.getEntry("phonesColumn").upper(),app.getEntry("NIDsColumn").upper(),app.getEntry("todayColumn").upper()]
                            if attendeesInformation[0] != "" and attendeesInformation[1] != "" and attendeesInformation[2] != "" and attendeesInformation[3] != "" and attendeesInformation[4] != "":
                                if attendeesInformation[0].isalpha() and attendeesInformation[1].isalpha() and attendeesInformation[2].isalpha() and attendeesInformation[3].isalpha() and attendeesInformation[4].isalpha() :
                                    if len(attendeesInformation[0]) == 1 and len(attendeesInformation[1]) == 1 and len(attendeesInformation[2]) == 1 and len(attendeesInformation[3]) == 1 and len(attendeesInformation[4]) == 1 :
                                        fileAvalability = backend.changeSettings(app.getEntry("labtopDeviceNumber"),app.getEntry("filePath"),attendeesInformation,workshopsColumns)
                                        if not fileAvalability:
                                            app.setLabel('warning','Check the file path and make sure to close it')
                                        else:                                        
                                        
                                            app.setButtonImage('save_setting', 'images/switch-7.gif')
                                            setting_saved = True
                                            app.disableEntry("labtopDeviceNumber")
                                            app.disableEntry("filePath")
                                            app.disableEntry("workshopColumn1")
                                            app.disableEntry("workshopColumn2")
                                            app.disableEntry("workshopColumn3")
                                            app.disableButton("importConfigurations")
                                            app.disableEntry("namesColumn")
                                            app.disableEntry("codesColumn")
                                            app.disableEntry("phonesColumn")
                                            app.disableEntry("NIDsColumn")
                                            app.disableEntry("todayColumn")
                                        
                                            app.setTransparency(100)
                                            app.setLabel('warning','')
                                    else:
                                        app.setLabel('warning',"Names,IDs,Phones,NIDs and Today fields must have only one character")
                                else:
                                    app.setLabel('warning',"Names,IDs,Phones,NIDs and Today fields must have alphabets")
                            else:
                                app.setLabel('warning',"Names,IDs,Phones,NIDs and Today fields mustn't be empty") 
                        else:
                            app.setLabel('warning',"Workshops fields must have only one character")
                    else:
                        app.setLabel('warning',"Workshops fields must have alphabets")
                else:
                    app.setLabel('warning',"Workshops fields mustn't be empty") 
            else:
               app.setLabel('warning',"File extension isn't 'xlsx'") 
            
        else:
            app.setLabel('warning',"Don't leave empty fields")
            
        
    elif win == 'save_setting' and setting_saved == True :
        app.setButtonImage('save_setting', 'images/switch-6.gif')
        setting_saved = False
        app.enableEntry("labtopDeviceNumber")
        app.enableEntry("filePath")
        app.enableEntry("workshopColumn1")
        app.enableEntry("workshopColumn2")
        app.enableEntry("workshopColumn3")
        app.enableButton("importConfigurations")
        app.enableEntry("namesColumn")
        app.enableEntry("codesColumn")
        app.enableEntry("phonesColumn")
        app.enableEntry("NIDsColumn")
        app.enableEntry("todayColumn")


yesCounter = 0 #if yes counter == 2 app will show the 'force closing' button instead of 'yes' button
def launch_exit (win) :
    global yesCounter
    
    if win == 'no' :
        app.hideSubWindow('exit')
        app.hideButton('use hard exit')
        app.showButton('yes')

    elif win == 'yes' :
        yesCounter += 1

        if not backend.closingExcelFile() :
            backend.saveLogs("======> Closing the application failed at " + str(datetime.datetime.now()))
            app.hideSubWindow('exit')
            app.setLabel("name", "Can't save your excel file!")
            app.setLabel("about", "Make sure that excel file isn't opened by any app and try again")
            if yesCounter >= 2 :
                app.hideButton('yes')
                app.showButton('use hard exit')

        else:
            backend.saveLogs("======> This session ended by closing the application at " + str(datetime.datetime.now()))
            backend.saveLogs("<====================================<>====================================>")
            app.stop()
    
    elif win == 'use hard exit':
        backend.forceSavingFile()
        backend.saveLogs("===\//===> Sheet was forced to be saved and closed " +
                 str(datetime.datetime.now()))
        backend.saveLogs("======> This session ended by closing the application at " + str(datetime.datetime.now()))
        backend.saveLogs("<====================================<>====================================>")
        app.stop()


########################################################################### all sub windows
        
################################################### this is a pop-up - exit
            
app.startSubWindow("exit", modal=True,blocking=False)
app.setBg('seashell')

app.setSticky("we")

app.addLabel('exit','exit?',0,0,3)
app.setLabelFg('exit','MediumSeaGreen')
app.getLabelWidget("exit").config(font="Arial  20")

app.setSticky("news")

app.addButton('yes',launch_exit,1,0)
app.addButton('no',launch_exit,1,2)
app.setButtonRelief("yes", "flat")
app.setButtonRelief("no", "flat")
app.setButtonBg('yes','MediumSeaGreen')
app.setButtonBg('no','MediumSeaGreen')
app.setButtonFg('yes','seashell')
app.setButtonFg('no','seashell')

app.addButton('use hard exit',launch_exit,1,0)
app.setButtonRelief("use hard exit", "flat")
app.setButtonBg('use hard exit','MediumSeaGreen')
app.setButtonFg('use hard exit','seashell')
app.hideButton('use hard exit')


app.setResizable(canResize=False)
app.setGeometry("300x200")
app.stopSubWindow()

################################################### this is a pop-up - setting

app.startSubWindow("Setting", modal=True)
app.setBg('seashell')
app.setSticky("we")

app.addLabel('space','',0,0,3)
app.getLabelWidget("space").config(font="Corbel  2")


### labtop number area

app.addLabel('lab',' What is your labtop number?',1,0,3)
app.setLabelFg('lab','MediumSeaGreen')
app.getLabelWidget("lab").config(font="Corbel  18")

app.addNumericEntry("labtopDeviceNumber",2,0,3)
app.setEntryDefault("labtopDeviceNumber", "0")
app.setEntryBg('labtopDeviceNumber','white')
app.setFileEntryRelief("labtopDeviceNumber", "flat")

### file path area

app.addLabel('filePath','Which excel file to work on?',4,0,3)
app.setLabelFg('filePath','MediumSeaGreen')
app.getLabelWidget("filePath").config(font="Corbel  18")

app.addFileEntry("filePath",5,0,3)
app.setFileEntryRelief("filePath", "flat")

### information area

app.addLabel('importantColumns','Specify the important columns',7,0,3)
app.setLabelFg('importantColumns','MediumSeaGreen')
app.getLabelWidget("importantColumns").config(font="Corbel  18")


###

app.addLabel('namesColumn',"What is names column?",8,0,3)
app.setLabelFg('namesColumn','MediumSeaGreen')
app.getLabelWidget("namesColumn").config(font="Corbel  12")

app.addEntry("namesColumn",9,0,3)
app.setEntryMaxLength("namesColumn", 1)

### 

app.addLabel('todayColumn',"What is today's column?",10,0,3)
app.setLabelFg('todayColumn','MediumSeaGreen')
app.getLabelWidget("todayColumn").config(font="Corbel  12")

app.addEntry("todayColumn",11,0,3)
app.setEntryMaxLength("todayColumn", 1)

###

app.addLabel('codesColumn','IDs',12,0)
app.setLabelFg('codesColumn','MediumSeaGreen')
app.getLabelWidget("codesColumn").config(font="Corbel  12")

app.addLabel('phonesColumn','Phones',12,1)
app.setLabelFg('phonesColumn','MediumSeaGreen')
app.getLabelWidget("phonesColumn").config(font="Corbel  12")

app.addLabel('NIDsColumn','National IDs',12,2)
app.setLabelFg('NIDsColumn','MediumSeaGreen')
app.getLabelWidget("NIDsColumn").config(font="Corbel  12")

app.addEntry("codesColumn",13,0)
app.addEntry("phonesColumn",13,1)
app.addEntry("NIDsColumn",13,2)
app.setEntryMaxLength("codesColumn", 1)
app.setEntryMaxLength("phonesColumn", 1)
app.setEntryMaxLength("NIDsColumn", 1)


## workshops area

app.addLabel('workshopColumn1','Workshop 1',14,0)
app.setLabelFg('workshopColumn1','MediumSeaGreen')
app.getLabelWidget("workshopColumn1").config(font="Corbel  12")

app.addLabel('workshopColumn2','Workshop 2',14,1)
app.setLabelFg('workshopColumn2','MediumSeaGreen')
app.getLabelWidget("workshopColumn2").config(font="Corbel  12")

app.addLabel('workshopColumn3','Workshop 3',14,2)
app.setLabelFg('workshopColumn3','MediumSeaGreen')
app.getLabelWidget("workshopColumn3").config(font="Corbel  12")

app.addEntry("workshopColumn1",15,0)
app.addEntry("workshopColumn2",15,1)
app.addEntry("workshopColumn3",15,2)
app.setEntryMaxLength("workshopColumn1", 1)
app.setEntryMaxLength("workshopColumn2", 1)
app.setEntryMaxLength("workshopColumn3", 1)


##

app.setSticky('ews')

app.addLabel('warning','',16,0,3)

app.addImageButton('save_setting', launch_setting, 'images/switch-6.gif',17,0,2)
app.setButtonBg('save_setting',"white")
app.setButtonRelief("save_setting", "flat")
app.setButtonActiveBg('save_setting', 'white')

app.addImageButton('importConfigurations', launch_setting, 'images/importConfiguration.gif',17,2)
app.setButtonBg('importConfigurations',"white")
app.setButtonRelief("importConfigurations", "flat")
app.setButtonActiveBg('importConfigurations', 'white')

app.setResizable(canResize=False)
app.setGeometry("600x700")
app.stopSubWindow()

##################################################################################################################################################

###################################################################################################################### these go in the main window


###################### function of all buttons of main windows


#############
# main function for searching

# takeActionUsingKeys = 0  # variable to let the binding keys work

def searchingForName(btn):
    # global takeActionUsingKeys

    # takeActionUsingKeys = 1

    app.disableEntry('givenNumber')

    givenNumber = str(int(app.getEntry("givenNumber")))
    tokensCounter = len(givenNumber)

    if tokensCounter < 7:
        searchForCode(int(givenNumber))
    elif tokensCounter < 14:
        searchForPhone(int(givenNumber))
    else:
        searchForNID(int(givenNumber))


#############
# function for searching using code

def searchForCode (givenNumber):

    return_data = backend.searchForCode(givenNumber)
    

    if return_data :
        returnedName = return_data[0]
        returnedCode = return_data[1]
        workshop1=return_data[2]
        workshop2=return_data[3]
        workshop3=return_data[4]
        
        app.setLabel("name", returnedName + ' - ' + str(returnedCode))
        app.setLabel("about", workshop1 + " | " + workshop2 + " | " + workshop3)

        app.showButton('right')
        app.showButton('wrong')
        # app.disableEntry('code')
        # app.disableEntry('phone')

    else :
        
        app.setLabel("name", "This ID isn't in database")
        app.setLabel("about", "Try writing it again carefully")


#############
# function for searching using phone number

def searchForPhone (givenNumber):
    
    return_data = backend.searchForPhone(givenNumber)

    if return_data :
        returnedName = return_data[0]
        returnedCode = return_data[1]
        workshop1=return_data[2]
        workshop2=return_data[3]
        workshop3=return_data[4]
        
        app.setLabel("name", returnedName + ' - ' + str(returnedCode))
        app.setLabel("about", workshop1 + " | " + workshop2 + " | " + workshop3)

        app.showButton('right')
        app.showButton('wrong')
        # app.disableEntry('code')
        # app.disableEntry('phone')

    else :
        app.setLabel("name", "This phone isn't in database")
        app.setLabel("about", "Try writing it again carefully")


#############
# function for searching using National ID

def searchForNID (givenNumber):
    
    return_data = backend.searchForNID(givenNumber)

    if return_data :
        returnedName = return_data[0]
        returnedCode = return_data[1]
        workshop1=return_data[2]
        workshop2=return_data[3]
        workshop3=return_data[4]
        
        app.setLabel("name", returnedName + ' - ' + str(returnedCode))
        app.setLabel("about", workshop1 + " | " + workshop2 + " | " + workshop3)

        app.showButton('right')
        app.showButton('wrong')
        # app.disableEntry('code')
        # app.disableEntry('phone')

    else :
        app.setLabel("name", "This national ID isn't in database")
        app.setLabel("about", "Try writing it again carefully")
        


def confirm_func_attend (btn):

    backend.add_by_day()
    app.setLabel("name", 'Done')  
    app.hideButton('right')
    app.hideButton('wrong')
    # app.enableEntry('code')
    # app.enableEntry('phone')
    app.enableEntry('givenNumber')
    app.setEntryFocus('givenNumber')

    
    

def confirm_func_not_attend (btn):
    
    app.hideButton('right')
    app.hideButton('wrong')
    # app.enableEntry('code')
    # app.enableEntry('phone')
    app.enableEntry('givenNumber')
    app.setEntryFocus('givenNumber')



########################################################################### all window elements


### settings & exit buttons ###

app.setSticky("w")
app.addImageButton('exit',launch, 'images/exit.gif',0,5)
app.setButtonBg('exit','LimeGreen')
app.setButtonActiveBg('exit', 'LimeGreen')
app.setButtonRelief("exit", "flat")

app.setSticky("e")
app.addImageButton('Setting',launch, 'images/settings.gif',0,4)
app.setButtonBg('Setting','LimeGreen')
app.setButtonActiveBg('Setting', 'LimeGreen')
app.setButtonRelief("Setting", "flat")


#####################

app.addLabel('vertical_space','',1,0,1,4)


app.setBg('seashell')

# app.setSticky('we')
# app.addNumericEntry('phone',1,1)
# app.getEntryWidget("phone").config(font="Arial 28")
# app.setEntryBg('phone','Gainsboro')


# app.setSticky('we')
# app.addImageButton('phone',searchForPhone, 'images/phone.gif',1,2)
# app.setButtonBg('phone','silver')
# app.setButtonRelief("phone", "groove")
# app.setButtonActiveBg('phone', 'silver')
# app.setEntrySubmitFunction("phone", searchForPhone)


# app.setSticky('we')
# app.addNumericEntry('code',1,3)
# app.getEntryWidget("code").config(font="Arial 28")
# app.setEntryBg('code','Gainsboro')


# app.setSticky('we')
# app.addImageButton('code',searchForCode, 'images/code.gif',1,4)
# app.setButtonBg('code','silver')
# app.setButtonRelief("code", "groove")
# app.setButtonActiveBg('code', 'silver')
# app.setEntrySubmitFunction("code", searchForCode)


app.setSticky('we')
app.addNumericEntry('givenNumber', 1, 1, 3)
app.getEntryWidget("givenNumber").config(font="Arial 28")
app.setEntryBg('givenNumber', 'Gainsboro')


app.setSticky('we')
app.addImageButton('givenNumber', searchingForName,
                   'images/code.gif', 1, 4, 1)
app.setButtonBg('givenNumber', 'silver')
app.setButtonRelief("givenNumber", "groove")
app.setButtonActiveBg('givenNumber', 'silver')
app.setEntrySubmitFunction("givenNumber", searchingForName)


app.setSticky('news')

app.addLabel('name','name will be shown here',2,1,4)
app.setLabelFg('name','white')
app.setLabelBg('name','LimeGreen')
app.getLabelWidget("name").config(font="Corbel 30")

app.addLabel('about','data is here',3,1,4)
app.setLabelFg('about','LimeGreen')
app.setLabelBg('about','white')
app.getLabelWidget("about").config(font="Corbel 24")

app.addImageButton('right',confirm_func_attend, 'images/like.gif',4,1,2)
app.setButtonBg('right','white')
app.setButtonActiveBg('right', 'white')
app.setButtonRelief("right", "flat")


app.addImageButton('wrong',confirm_func_not_attend, 'images/dislike.gif',4,3,2)
app.setButtonBg('wrong','white')
app.setButtonActiveBg('wrong', 'white')
app.setButtonRelief("wrong", "flat")

app.hideButton('right')
app.hideButton('wrong')


app.addLabel('horizontal_space','',5,1,4)


#function that work to make setting appear after opening app
def openSettingWindow() :
    app.showSubWindow('Setting')
app.after(100, openSettingWindow)

app.setTitle("Check-in Application")
app.setIcon("images/appLogo.ico")
app.setTransparency(70)
app.setGeometry("fullscreen")
app.go()
