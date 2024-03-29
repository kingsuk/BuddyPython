import win32com.client
import WriteToFile as WF
from difflib import SequenceMatcher
import win32com.client as comclt
import win32com.client as win32
import time
import speech_recognition as sr
from pynput.keyboard import Key, Controller

keyboard = Controller()
MailWrite = True

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
                
                if msg.SenderEmailType=='EX':
                    EmailNamePair["email"] = msg.Sender.GetExchangeUser().PrimarySmtpAddress
                    EmailNamePair["name"] = msg.Sender.GetExchangeUser().PrimarySmtpAddress.split('@')[0]
                else:
                    EmailNamePair["email"] = msg.SenderEmailAddress
                    EmailNamePair["name"] = msg.SenderEmailAddress.split('@')[0]

                EmailList.append(EmailNamePair)

    WF.WriteToFile(RemoveDuplicates(EmailList),"outlookemail.json")
#CreateEmailDB()
def CheckEmailMatch(userInput):
    data = WF.ReadFromFile("outlookemail.json")
    topMatch = {}
    topMatchPercent = 0
    for email in data:
        matchingScore = similar(userInput, email["name"])
        
        if matchingScore > topMatchPercent:
            topMatchPercent = matchingScore
            email["score"] = matchingScore
            topMatch = email

    return topMatch

def StartDictation(mail):
    wsh= comclt.Dispatch("WScript.Shell")
    wsh.AppActivate("outlook.application") # select another application

    
    def callback(recognizer, audio):
        # received audio data, now we'll recognize it using Google Speech Recognition
        try:
            # for testing purposes, we're just using the default API key
            # to use another API key, use `r.recognize_google(audio, key="GOOGLE_SPEECH_RECOGNITION_API_KEY")`
            # instead of `r.recognize_google(audio)`
            # print("Google Speech Recognition thinks you said " + recognizer.recognize_google(audio))
            textResponse = recognizer.recognize_google(audio)
            print(textResponse)
            global MailWrite
            print(MailWrite) 
            if textResponse == "send" or textResponse == "send the mail" or textResponse == "send the email" or textResponse == "send it":
                #mail.Send()
                keyboard.press(Key.alt)
                keyboard.press("s")
                keyboard.release(Key.alt) 
                keyboard.release("s")
                MailWrite = False
            elif textResponse == "next line" or textResponse == "paragraph" or textResponse == "go to the next line":
                keyboard.press(Key.enter)
                keyboard.release(Key.enter)
            elif textResponse == "period":
                wsh.SendKeys(".")
            elif textResponse == "end dictation":
                MailWrite = False
            else:
                wsh.SendKeys(textResponse+" ")
                #MailWrite = False

        except sr.UnknownValueError:
             print("Google Speech Recognition could not understand audio")
        
        except sr.RequestError as e:
             print("Could not request results from Google Speech Recognition service; {0}".format(e))


    r = sr.Recognizer()
    m = sr.Microphone()
    with m as source:
        r.adjust_for_ambient_noise(source)  # we only need to calibrate once, before we start listening

    # start listening in the background (note that we don't have to do this inside a `with` statement)
    stop_listening = r.listen_in_background(m, callback)
    # `stop_listening` is now a function that, when called, stops background listening

    # do some unrelated computations for 5 seconds
    while MailWrite: time.sleep(0.1)  # we're still listening even though the main thread is doing other things

    # calling this function requests that the background listener stop listening
    stop_listening(wait_for_stop=False)

    # do some more unrelated things
    #while True: time.sleep(0.1)


def CreateEmail(toAddress,subject):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = toAddress
    mail.Subject = subject
    #mail.Body = 'Message body'
    #mail.HTMLBody = '' #this field is optional
    mail.Display()
    StartDictation(mail)
    
