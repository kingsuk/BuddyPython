from win32com.client import Dispatch
from tabulate import tabulate
from datetime import datetime, time, timedelta
import pdb

def getMeetingsByDays(startDay,numberOfDays,endDay,format=1):
     if format==1:
          OUTLOOK_FORMAT = '%m/%d/%Y %I:%M %p'
     else:
          OUTLOOK_FORMAT = '%m/%d/%Y %H:%M' 

     outlook = Dispatch("Outlook.Application")
     ns = outlook.GetNamespace("MAPI")

     appointments = ns.GetDefaultFolder(9).Items 

     # Restrict to items in the next 30 days (using Python 3.3 - might be slightly different for 2.7)
     if startDay == None:
          begin = datetime.today()
     else:
          begin = datetime.strptime(startDay, '%Y-%m-%d')

     #begin = datetime.today()
     #begin = datetime.strptime(startDay, '%Y-%m-%d')
     end = begin + timedelta(days = numberOfDays)
     restriction = "[Start] >= '" + begin.strftime("%m/%d/%Y") + "' AND [End] <= '" +end.strftime("%m/%d/%Y") + "'"
     appointments = appointments.Restrict(restriction)

     appointments.Sort("[Start]")
     appointments.IncludeRecurrences = "True"

     # Iterate through restricted AppointmentItems and print them
     calcTableHeader = ['Title', 'Organizer', 'Start', 'Duration(Minutes)']
     calcTableBody = []
     returnDataset = {}
     #pdb.set_trace()
     for appointmentItem in appointments:
          returnDataset = {}
          returnDataset["Subject"] = appointmentItem.Subject
          returnDataset["Body"] = appointmentItem.Body
          returnDataset["Organizer"] = appointmentItem.Organizer
          returnDataset["Start"] = appointmentItem.Start.Format(OUTLOOK_FORMAT)
          returnDataset["Duration"] = appointmentItem.Duration
          returnDataset["Location"] = appointmentItem.location
          returnDataset["End"] = appointmentItem.end.Format(OUTLOOK_FORMAT)
          calcTableBody.append(returnDataset)

          
     return (calcTableBody)

def is_time_between(begin_time, end_time, check_time=None):
    # If check time is not given, default to current UTC time
    check_time = check_time or datetime.utcnow().time()
    if begin_time < end_time:
        return check_time >= begin_time and check_time <= end_time
    else: # crosses midnight
        return check_time >= begin_time or check_time <= end_time


def GetFreeTimes(SlotTiming,meetings):
     #print(meetings)
     OUTLOOK_FORMAT = "%I:%M %p"
     freeSlots = []
     #meetingEndDateTime = datetime.now()
     loopIndexForEndTimeOut = 0
     for meeting in meetings:
          global lastMeetingEndDateTime
          meetingStartDateTime = datetime.strptime(meeting['Start'], '%m/%d/%Y %H:%M')
          meetingStartTime = datetime.strptime(meeting['Start'], '%m/%d/%Y %H:%M').time()
          meetingEndTime = datetime.strptime(meeting['End'], '%m/%d/%Y %H:%M').time()
          is_start_time_between = is_time_between(datetime.strptime(SlotTiming[0], '%H:%M').time(),datetime.strptime(SlotTiming[1], '%H:%M').time(),meetingStartTime)
          is_end_time_between = is_time_between(datetime.strptime(SlotTiming[0], '%H:%M').time(),datetime.strptime(SlotTiming[1], '%H:%M').time(),meetingEndTime)

          #checking if start time is between shift time
          if is_start_time_between:
               #check if slot start time and meeting start time is not same
               if meetingStartTime != datetime.strptime(SlotTiming[0], '%H:%M').time():
                    if len(freeSlots) == 0:
                         #it is the first slot
                         currentFreeTime = [datetime.strptime(SlotTiming[0], '%H:%M').time(), (meetingStartDateTime - timedelta(minutes=1)).time()]
                         freeSlots.append(currentFreeTime)
                    else:
                         #getting the last meeting end time
                         lastMeetingEndTime = lastMeetingEndDateTime
                         nextFreeSlotStartTime = lastMeetingEndTime + timedelta(minutes=1)
                         currentFreeTime = [nextFreeSlotStartTime.time(), (meetingStartDateTime - timedelta(minutes=1)).time()]
                         freeSlots.append(currentFreeTime)

          if is_start_time_between != True and is_end_time_between:
               if loopIndexForEndTimeOut ==0:
                    print("In first")
               else:
                    lastMeetingEndTime = lastMeetingEndDateTime
                    nextFreeSlotStartTime = lastMeetingEndTime + timedelta(minutes=1)
                    currentFreeTime = [nextFreeSlotStartTime.time(), (meetingStartDateTime - timedelta(minutes=1)).time()]
                    freeSlots.append(currentFreeTime)

               print("Here")
               loopIndexForEndTimeOut += 1
          lastMeetingEndDateTime = datetime.strptime(meeting['End'], '%m/%d/%Y %H:%M')
          
     #return freeSlots


     #check if last meeting end time is between slot
     is_last_time_between_slot = is_time_between(datetime.strptime(SlotTiming[0], '%H:%M').time(),datetime.strptime(SlotTiming[1], '%H:%M').time(),lastMeetingEndDateTime.time())
     
     if is_last_time_between_slot:
          currentFreeTime = [(lastMeetingEndDateTime + timedelta(minutes=1)).time(),datetime.strptime(SlotTiming[1], '%H:%M').time()]
          freeSlots.append(currentFreeTime)

     freeSlotsFormatted = []
     for i in freeSlots:
          freetime = {}
          index = 0
          for j in i:
               index += 1
               #print(datetime.strptime(j, '%H:%M'))
               if index ==1:
                    freetime["Start"] = j.strftime(OUTLOOK_FORMAT)
               else:
                    freetime["End"] = j.strftime(OUTLOOK_FORMAT)
          freeSlotsFormatted.append(freetime)

     return freeSlotsFormatted



def GetTimeSlots(SlotName):
     if SlotName == 'shift':
          return ['07:00','22:00']
     elif SlotName == 'morning':
          return ['00:01','12:00']
     elif SlotName == 'afternoon':
          return ['12:01','17:00']
     elif SlotName == 'evening':
          return ['17:01','20:00']
     elif SlotName == 'night':
          return ['20:01','00:00']

#GetFreeTimes(GetTimeSlots('afternoon'))
#GetFreeTimes(GetTimeSlots('shift'),getMeetingsByDays("2019-04-08"))
#meetings = getMeetingsByDays("2019-04-09",1,None,0)
#print(GetFreeTimes(GetTimeSlots('shift'),meetings))