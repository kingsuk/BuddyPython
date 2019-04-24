import GoogleSpeechToText as GSTT
import winspeech
import Helper.luis as luis
import config
import WriteToFile as WF
import os
import WriteFileHelper as WFH

from datetime import datetime, time, date

def CreateForwardLoop(previousText,queryText):
    print(queryText)
    winspeech.say_wait(queryText)

    if config.voiceEnable == True:
        reminderAdditionalInput = GSTT.getSpeechToText()
    else:
        reminderAdditionalInput = input(" : ")

    luisOutput = luis.AnalyseIntent(previousText + reminderAdditionalInput)
    return CreateReminder(luisOutput)

def CreateReminder(output):
    try:
        entities = output['entities']
        reminderText = ""
        reminderDateTime = ""
        reminderDate = ""
        reminderTime = ""
        previousText = ""

        if len(entities) > 0:
            for entity in entities:
                print(entity["type"])
                if entity["type"] == "Reminder.Text":
                    reminderText = entity["entity"]

                elif entity["type"] == "builtin.datetimeV2.datetime":
                    resolutionValues = entity["resolution"]["values"]
                    for values in resolutionValues:
                        reminderDateTime = values["value"]

                elif entity["type"] == "builtin.datetimeV2.date":
                    resolutionValues = entity["resolution"]["values"]
                    for values in resolutionValues:
                        reminderDate = values["value"]

                elif entity["type"] == "builtin.datetimeV2.time":
                    resolutionValues = entity["resolution"]["values"]
                    for values in resolutionValues:
                        reminderTime = values["value"]
            
            if reminderText != "" and reminderDateTime != "":
                
                returnObject = {}
                returnObject["reminderText"] = reminderText
                returnObject["reminderDateTime"] = reminderDateTime

                return returnObject

        else:
            print("No Entity")


        if reminderText == "":
            previousText = output["query"] + " to "
            return CreateForwardLoop(previousText,"Please provide the reminder text")

        elif reminderDateTime == "":
            if reminderDate == "" and reminderTime == "":
                previousText = output["query"] + " "
                return CreateForwardLoop(previousText,"Please provide the date and time")

            elif reminderDate == "":
                previousText = output["query"] + " "
                return CreateForwardLoop(previousText,"Please provide the date")

            elif reminderTime == "":
                previousText = output["query"] + " "
                return CreateForwardLoop(previousText,"Please provide the time")
        
    except Exception as e:
        print("Error in CreateReminder "+str(e))

def CreateReminderDialog(output):
    reminder = CreateReminder(output)
    print(reminder)
    winspeech.say_wait("your reminder text is "+reminder["reminderText"]+", reminder is set on "+reminder["reminderDateTime"]+ " .")
    print("please say confirm to save it and say cancel to dismiss.")
    winspeech.say_wait("please say confirm to save it and say cancel to dismiss.")
    if config.voiceEnable == True:
        userInput = GSTT.getSpeechToText()
    else:
        userInput = input(" : ")

    if userInput == "confirm":
        WF.AppendToFile(reminder)
    else:
        print("Ok, dismissing it")
        winspeech.say_wait("ok, dismissing it")


#======================================================================================

reminderNativeAppPath = r"C:\Users\accenture.robotics\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\buddy.lnk"

def ReadOutReminders(reminders):

    reminderHeader = ""
    if len(reminders) == 0:
        reminderHeader = "You don't have any reminder"
    elif len(reminders) == 1:
        reminderHeader = "You have one reminder"
    else:
        reminderHeader = "You have "+str(len(reminders))+" reminders"

    WFH.WriteToCurrentJson(reminderHeader,"reminder",reminders,reminderNativeAppPath)


    # totalJsonData = {}
    # headerList = []
    # dataType = []
    # headerList.append(reminderHeader)
    # dataType.append("reminder")
    # totalJsonData["dataType"] = dataType
    # totalJsonData["currentHeader"] = headerList
    # totalJsonData["currentData"] = reminders

    # WF.WriteToFile(totalJsonData,"currentdata.json")

    # os.startfile(reminderNativeAppPath)

    print(reminderHeader)
    winspeech.say_wait(reminderHeader)


    for reminder in reminders:
        currentDateTimeObject = datetime.strptime(reminder["reminderDateTime"] , '%Y-%m-%d %H:%M:%S')
        changedDateTimeFormat = currentDateTimeObject.strftime('%I:%M %p')
        outputString = reminder["reminderText"] +" at "+ changedDateTimeFormat
        print(outputString)
        winspeech.say_wait(outputString)


def FindReminders(output):
    reminderDate = ""
    try:
        entities = output['entities']
        #print(len(entities))
        if len(entities) > 0:
            for entity in entities:
                if entity["type"] == "builtin.datetimeV2.date":
                    resolutionValues = entity["resolution"]["values"]
                    for values in resolutionValues:
                        reminderDate = datetime.strptime(values["value"], '%Y-%m-%d').date()
                
        else:
            print("Show reminders of today")
            reminderDate = date.today()

        

        print(reminderDate)

        reminders = WF.ReadFromFile("data.json")["reminders"]
        remindersToShow = []

        for reminder in reminders:
            currentDate = datetime.strptime(reminder['reminderDateTime'], '%Y-%m-%d %H:%M:%S').date()

            if currentDate == reminderDate:
                remindersToShow.append(reminder)

        ReadOutReminders(remindersToShow)
    
    except Exception as e:
        print("Error in Find Reminder "+str(e))