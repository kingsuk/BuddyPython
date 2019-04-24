import win32com.client as win32
import winspeech
import speech_recognition as sr


# Start an in-process recognizer. Don't want the shared one with built-in windows commands.
winspeech.initialize_recognizer(winspeech.INPROC_RECOGNIZER)


def CreateMail(text, subject, recipientNameInput):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = GetNameEmailMap()[recipientNameInput]
    winspeech.say("Please provide the subject line")
    mail.Subject = input("Please provide the subject line : ")
    winspeech.say("Please provide body of the mail, Please say Buddy Stop in an interval to stop dictating")
    mail.HtmlBody = input("Please provide body of the mail : ")
    mail.Display(True)
    return

def GetNameEmailMap():
    NameEmailMap = {}
    NameEmailMap["kingsuk"] = "kingsuk.majumder@accenture.com"
    NameEmailMap["arindam"] = "arindam.f.ghosh@accenture.com"
    NameEmailMap["moushom"] = "moushom.borah@accenture.com"
    return NameEmailMap