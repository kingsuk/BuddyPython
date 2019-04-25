import win32com.client
import WriteToFile as WF
from difflib import SequenceMatcher

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

def RemoveDuplicates(duplicate): 
    final_list = [] 
    for num in duplicate: 
        if num not in final_list: 
            final_list.append(num) 
    return final_list

def CreateEmailDB():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                        # the inbox. You can change that number to reference
                                        # any other folder
    messages = inbox.Items

    EmailList = []
    for msg in messages:
        if msg.Class==43:
                EmailNamePair = {}
                EmailNamePair["name"] = ""
                if msg.SenderEmailType=='EX':
                    EmailNamePair["email"] = msg.Sender.GetExchangeUser().PrimarySmtpAddress
                else:
                    EmailNamePair["email"] = msg.SenderEmailAddress

                EmailList.append(EmailNamePair)

    WF.WriteToFile(RemoveDuplicates(EmailList),"outlookemail.json")

def CheckEmailMatch(userInput):
    data = WF.ReadFromFile("outlookemail.json")
    topMatch = {}
    topMatchPercent = 0
    for email in data:
        matchingScore = similar(userInput, email["email"])
        
        if matchingScore > topMatchPercent:
            topMatchPercent = matchingScore
            email["score"] = matchingScore
            topMatch = email

    return topMatch
