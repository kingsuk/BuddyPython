import win32com.client
import win32com.client as comclt
import win32com.client as win32
import time
import speech_recognition as sr
from pynput.keyboard import Key, Controller
import os

keyboard = Controller()
MailWrite = True

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
                MailWrite = False
            
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
    while MailWrite: 
        time.sleep(0.1)  # we're still listening even though the main thread is doing other things
        #print("loop")

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
    print(mail)
    mail.Display()
    print(os.getpid())
    StartDictation(mail)

CreateEmail("kingsuk.majumder@accenture.com","Demo Subject")