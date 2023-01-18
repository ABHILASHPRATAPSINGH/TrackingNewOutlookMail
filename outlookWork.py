import os
import win32com.client
import pythoncom
import re
import win32com.client
from utils import getSrcFolderPath
import win32com.client as win32


class OutlookMailTask:
    def __init__(self, mail):
        self.mail=mail

    def validateAPI(self):
        subjectLine=self.mail.Subject
        senderUsername=self.mail.SenderName
        # getSenderEmailAddress=self.mail.SenderEmailAddress
        splitSubject=subjectLine.split(" ")

        subject=splitSubject[0]
        username=splitSubject[1]
        pwd=splitSubject[2]

        isValidAPI=False
        isValidUsernamePassword=True
        if subject.lower()=='runapi':
            isValidAPI=True
            if username.lower()!='abhilash':
                msg='Hi ' + senderUsername+',\n\nYou are sending Invalid Username'
            elif pwd.lower()!='password':
                msg='Hi ' + senderUsername+',\n\nYou are sending Invalid Password'
            else:
                isValidUsernamePassword=True
                msg='Hi '+senderUsername+',\n\nWe are eceiving your request.\n\nBot is running.....\nWe will send confirmation message after processing.'

        return(isValidAPI,isValidUsernamePassword,msg)


    def performMailTask(self):
        try:
            senderUsername=self.mail.SenderName
            getSenderEmailAddress=self.mail.SenderEmailAddress

            respValidAPI=self.validateAPI()
            isValidAPI=respValidAPI[0]
            isValidUsernamePassword=respValidAPI[1]
            msg=respValidAPI[2]

            if isValidAPI==True:
                outlook1=win32.Dispatch('outlook.application')
                mail1=outlook1.CreateItem(0)
                mail1.Subject='Calling stock API'
                mail1.To=getSenderEmailAddress
                mail1.Body=msg
                mail1.Send()
                print("Message 1 Sent!!")

            if isValidUsernamePassword==True:
                filepath=os.path.join(getSrcFolderPath(),'get5DaysData.xlsx')
                print(filepath)
                outlook2=win32.Dispatch('outlook.application')
                mail2=outlook2.CreateItem(0)
                msg2="Hi "+senderUsername+',\n\nProcessing finished!!\n\nPlease collect output from attachements.\n\nThank you!!'
                mail2.Subject='Calling stock APi'
                mail2.To=getSenderEmailAddress
                mail2.Attachments.Add(filepath)
                mail2.Body=msg2
                mail2.Send()
                print("Message 2 Sent!!")
        except :
            pass

