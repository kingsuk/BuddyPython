import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                    # the inbox. You can change that number to reference
                                    # any other folder
messages = inbox.Items

for msg in messages:
       if msg.Class==43:
           if msg.SenderEmailType=='EX':
               print(msg.Sender.GetExchangeUser().PrimarySmtpAddress)
           else:
               print(msg.SenderEmailAddress)