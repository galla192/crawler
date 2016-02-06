import win32com.client              #corresponing package file must be installed

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6) #6 is the number for the inbox
                                    #google it for more info
messages = inbox.Items
message = messages.GetFirst()
body_content = message.body
print "current # of messages: "+str(len(messages))
print "Deleting message "+str(len(messages))+"\n"+"From: "+str(message.Sender)+"\n"+"Body: "+body_content
message.Delete()

