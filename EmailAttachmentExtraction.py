import win32com.client
import os

application = win32com.client.Dispatch('Outlook.Application')
namespace = application.GetNamespace('MAPI')

inboxID = 6
inboxFolder = namespace.GetDefaultFolder(inboxID)
moveToFolder = inboxFolder.Folders.Item('DataAudit')
subject = 'Parent_Changes - Prod : Audit Details'
fileExt = 'xlsx'
filePath = 'C:/test/'

for counter in range(inboxFolder.Items.Count, 0, -1):
    email = inboxFolder.Items.Item(counter)

    if email.Subject == subject:
        attachments = []
        
        for attachment in email.Attachments:
            aName = email.SentOn.strftime("%m.%d.%Y") + ' - ' + attachment.FileName
            if not attachment.FileName.endswith(fileExt):
                continue
               
            fileSaveLocation = os.path.join(filePath, aName)
            attachment.SaveAsFile(fileSaveLocation)
            attachments.append(fileSaveLocation)
            email.Move(moveToFolder)
