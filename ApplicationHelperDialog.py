import sys
import os
from pathlib import Path
import winspeech
import GoogleSpeechToText as GSTT
import config

userPath = r"C:\Users\Accenture.Robotics\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
osPath = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"

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
        
    return fileMaps

def OpenApplication(applicationDisplayName,allApplicationList):
    try:
        os.startfile(allApplicationList[applicationDisplayName])
        print("Okay, Opening "+applicationDisplayName)
        winspeech.say_wait("Okay, Opening "+applicationDisplayName)
    except:
        print("Could not find the application you are asking for.")
        winspeech.say_wait("Could not find the application you are asking for.")

installedApplicationMaps = CreateFileMaps()

def OpenApplicationParser(output):
    entities = output['entities']
    applicattionNameInput = ""
    if len(entities) > 0:
        entity = entities[0]["entity"].lower()
        applicattionNameInput = entity
    else:
        print("Please provide the application name: ")
        winspeech.say_wait("Please provide the application name: ")
        
        if config.voiceEnable == True:
            speechResult = GSTT.getSpeechToText()
        else:
            speechResult = input(" : ")
        
        applicattionNameInput = speechResult

    OpenApplication(applicattionNameInput,installedApplicationMaps)