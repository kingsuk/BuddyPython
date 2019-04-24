import json
import winspeech

def AppendToFile(currentData):
    with open('data.json') as json_file:  
        data = json.load(json_file)
        data['reminders'].append(currentData)
        with open('data.json', 'w') as outfile:  
            json.dump(data, outfile)
            print("Ok, your reminder is saved")
            winspeech.say_wait("Ok, your reminder is saved")


def ReadFromFile(fileName):
    with open(fileName) as json_file:  
        data = json.load(json_file)
        return data

def WriteToFile(jsonData,fileName):
        with open(fileName, 'w') as outfile:  
                json.dump(jsonData, outfile)
                print("Ok, data is saved")
                