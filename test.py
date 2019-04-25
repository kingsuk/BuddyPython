import win32com.client as comclt
import win32com.client as win32
import GoogleSpeechToText as GSTT
import time
import speech_recognition as sr

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'To address'
mail.Subject = 'Message subject'
mail.Body = 'Message body'
mail.HTMLBody = '' #this field is optional
mail.Display()

wsh= comclt.Dispatch("WScript.Shell")
wsh.AppActivate("outlook.application") # select another application

MailWrite = True

# this is called from the background thread
def callback(recognizer, audio):
    # received audio data, now we'll recognize it using Google Speech Recognition
    try:
        # for testing purposes, we're just using the default API key
        # to use another API key, use `r.recognize_google(audio, key="GOOGLE_SPEECH_RECOGNITION_API_KEY")`
        # instead of `r.recognize_google(audio)`
        # print("Google Speech Recognition thinks you said " + recognizer.recognize_google(audio))
        textResponse = recognizer.recognize_google(audio)
        print(textResponse)
        wsh.SendKeys(textResponse+" ")

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
while True: time.sleep(0.1)  # we're not listening anymore, even though the background thread might still be running for a second or two while cleaning up and stopping
    