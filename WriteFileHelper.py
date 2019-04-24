import WriteToFile as WF
import os
import psutil    


def WriteToCurrentJson(headerText,dataTypeText,data,filePath):
    totalJsonData = {}
    headerList = []
    dataType = []
    headerList.append(headerText)
    dataType.append(dataTypeText)
    totalJsonData["dataType"] = dataType
    totalJsonData["currentHeader"] = headerList
    totalJsonData["currentData"] = data

    WF.WriteToFile(totalJsonData,"currentdata.json")

    if "Buddy.exe" in (p.name() for p in psutil.process_iter()):
        print("Already running closing it")
        os.system("TASKKILL /F /IM Buddy.exe")
        
    os.startfile(filePath)