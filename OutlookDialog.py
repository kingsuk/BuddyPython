from datetime import datetime, time, timedelta,date
# from datetime import datetime, time, date
import winspeech
import OutlookMeetings as OM
import WriteToFile as WF
import os
import WriteFileHelper as WFH

reminderNativeAppPath = r"C:\Users\accenture.robotics\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\buddy.lnk"


def ReadOutMeetings(meetings):
    print(meetings)
    headerText = ""
    if len(meetings) == 0:
        headerText = "You don't have any meeting."
    elif len(meetings) == 1:
        headerText = "You have "+str(len(meetings))+" meeting"
    else:
        headerText = "You have "+str(len(meetings))+" meetings"


    
    WFH.WriteToCurrentJson(headerText,"meetings",meetings,reminderNativeAppPath)
    # totalJsonData = {}
    # headerList = []
    # dataType = []
    # headerList.append(headerText)
    # dataType.append("meetings")
    # totalJsonData["dataType"] = dataType
    # totalJsonData["currentHeader"] = headerList
    # totalJsonData["currentData"] = meetings

    # WF.WriteToFile(totalJsonData,"currentdata.json")
    # os.startfile(reminderNativeAppPath)

    print(headerText)
    winspeech.say_wait(headerText)
    for meeting in meetings:
            outputString = "Meeting Subject is: "+meeting["Subject"] + ", Meeting Organizer is: "+meeting["Organizer"] + ", Meeting is from "+datetime.strptime(meeting['Start'], '%m/%d/%Y %I:%M %p').strftime('%I:%M %p')+" to "+datetime.strptime(meeting['End'], '%m/%d/%Y %I:%M %p').strftime('%I:%M %p')+" at "+meeting["Location"]+"."
            print(outputString)
            winspeech.say_wait(outputString)


def ReadOutFreeSlots(freeSlots,startDate):
    if startDate == None:
        startDateSpeak = str(date.today())
    else:
        startDateSpeak = startDate
    
    headerText = "You have the following free slots on "+startDateSpeak

    WFH.WriteToCurrentJson(headerText,"free-slots",freeSlots,reminderNativeAppPath)

    # totalJsonData = {}
    # headerList = []
    # dataType = []
    # headerList.append(headerText)
    # dataType.append("free-slots")
    # totalJsonData["dataType"] = dataType
    # totalJsonData["currentHeader"] = headerList
    # totalJsonData["currentData"] = freeSlots

    # WF.WriteToFile(totalJsonData,"currentdata.json")

    # os.startfile(reminderNativeAppPath)

    print(headerText)
    winspeech.say_wait(headerText)
    for freeSlot in freeSlots:
        outputString = "from "+freeSlot["Start"] + " to "+freeSlot["End"]
        print(outputString)
        winspeech.say_wait(outputString)
   

def ShowMeetings(output):

    try:
        entities = output['entities']
        #print(len(entities))
        if len(entities) > 0:
            if entities[0]["type"] == "builtin.datetimeV2.date":
                startDate = entities[0]["resolution"]["values"][0]["value"]
                print(startDate)
                meetings = OM.getMeetingsByDays(startDate,1,None)
                ReadOutMeetings(meetings)
            else:
                print("Different date type")
        else:
            meetings = OM.getMeetingsByDays(None,1,None)
            ReadOutMeetings(meetings)
    
    except Exception as e:
        print("Error in show meetings "+str(e))

def FreeSlots(output):

    try:
        entities = output['entities']
        #print(len(entities))
        if len(entities) > 0:
            if entities[0]["type"] == "builtin.datetimeV2.date":
                startDate = entities[0]["resolution"]["values"][0]["value"]
                print(startDate)
                meetings = OM.getMeetingsByDays(startDate,1,None,0)
                if len(meetings)!=0:
                    freeSlots = OM.GetFreeTimes(OM.GetTimeSlots('shift'),meetings)
                    ReadOutFreeSlots(freeSlots,startDate)
                else:
                    print("You don't have any meetings")
                    winspeech.say_wait("you don't have any meetings on "+startDate)
            else:
                print("Different date type")
        else:
            meetings = OM.getMeetingsByDays(None,1,None,0)
            if len(meetings)!=0:
                freeSlots = OM.GetFreeTimes(OM.GetTimeSlots('shift'),meetings)
                ReadOutFreeSlots(freeSlots,None)
            else:
                print("You don't have any meetings")
                winspeech.say_wait("you don't have any meetings")
    
    except Exception as e:
        print("Error in show meetings "+str(e))
    
#=======================================================================================================

def CreateEmail(output):
    try:
        entities = output['entities']
        #print(len(entities))
        if len(entities) > 0:
            print("Entitites")
        else:
            print("Please provide the recipient of the email")
            
    
    except Exception as e:
        print("Error in show create mail "+str(e))

