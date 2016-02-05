import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                    # the inbox. You can change that number to reference
                                    # any other folder
messages = inbox.Items
message = messages.GetFirst()
body_content = message.body
#print body_content
print "current # of messages: "+str(len(messages))
print "Deleting message "+str(len(messages))+"\n"+"From: "+str(message.Sender)+"\n"+"Body: "+body_content
message.Delete()

