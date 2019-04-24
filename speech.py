import speech_recognition as sr
import requests
import json
import winspeech
import sys
import os
from pathlib import Path

headers = {
    # Request headers
    'Ocp-Apim-Subscription-Key': '08e66a5c2e024e28963d0b23e1702b15',
}

userPath = r"C:\Users\Accenture.Robotics\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
osPath = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"

# Start an in-process recognizer. Don't want the shared one with built-in windows commands.
winspeech.initialize_recognizer(winspeech.INPROC_RECOGNIZER)


def OpenApplication(applicationDisplayName,allApplicationList):
    try:
        os.startfile(allApplicationList[applicationDisplayName])
        print("Okay, Opening "+applicationDisplayName)
        winspeech.say("Okay, Opening "+applicationDisplayName)
    except:
        print("Could not find the application you are asking for.")
        winspeech.say("Could not find the application you are asking for.")

def findIntent(headers,params):
    try:
        r = requests.get('https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/7c51f1a4-4c38-4c6a-8929-337c6e5c7e6f',headers=headers, params=params)
        result = r.json()
        print(result)
        topScoringIntent = result['topScoringIntent']['intent']
        
        if topScoringIntent == "openApplication":
            entities = result['entities']
            if len(entities) > 0:
                entity = entities[0]["entity"].lower()
                allApplicationList = CreateFileMaps()
                OpenApplication(entity,allApplicationList)
                
            else:
                print("Can't understand the application name, or it might not be present")
                winspeech.say("Can't understand the application name, or it might not be present")

    except Exception as e:
        print("[Errno {0}] {1}".format(e.errno, e.strerror))



def getSpeechToText():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Please wait. Calibrating microphone...")
        # listen for 1 second and create the ambient noise energy level
        r.adjust_for_ambient_noise(source, duration=1)
        print("Say something!")
        audio = r.listen(source,phrase_time_limit=3)
 
    try:
        # for testing purposes, we're just using the default API key
        # to use another API key, use `r.recognize_google(audio, key="GOOGLE_SPEECH_RECOGNITION_API_KEY")`
        # instead of `r.recognize_google(audio)`
        result = r.recognize_google(audio)
        print(result)
        return result
        
    except sr.UnknownValueError:
        print("Google Speech Recognition could not understand audio")
        return ""
    except sr.RequestError as e:
        print("Could not request results from Google Speech Recognition service; {0}".format(e))
        return ""

def CreateFileMaps():
    fileMaps = {}
    userPathFiles = list(Path(userPath).rglob("*.[lL][nN][kK]"))
    osPathFiles = list(Path(osPath).rglob("*.[lL][nN][kK]"))

    totalPathFiles = userPathFiles + osPathFiles

    for fileFullPath in totalPathFiles:
        #print(fileFullPath)
        base=os.path.basename(fileFullPath)
        displayName = os.path.splitext(base)[0].lower()
        fileMaps[displayName] = fileFullPath
        print(displayName+",")
        

    return fileMaps

while True:
    #print(CreateFileMaps())
    # winspeech.say("Hello")
    # params = {
    #         # Query parameter
    #         'q': "open notepad",
    #         # Optional request parameters, set to default values
    #         'timezoneOffset': '0',
    #         'verbose': 'false',
    #         'spellCheck': 'false',
    #         'staging': 'false',
    #     }
    # findIntent(headers,params)

    returnText = getSpeechToText()
    
    if returnText != "":
        params = {
            # Query parameter
            'q': returnText,
            # Optional request parameters, set to default values
            'timezoneOffset': '0',
            'verbose': 'false',
            'spellCheck': 'false',
            'staging': 'false',
        }
        findIntent(headers,params)