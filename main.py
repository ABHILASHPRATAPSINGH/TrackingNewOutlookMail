
import win32com.client
import pythoncom
import re
from outlookWork import OutlookMailTask




class Handler_Class:

    def OnNewMailEx(self,receivedItemsIDs):
        # ReceivedItemsIds is a collection of mail IDs separated by ','
        # You know some, sometimes more than 1 mail is received at the same moment.

        for ID in receivedItemsIDs.split(','):
            mail=outlook.Session.GetItemFromID(ID)
            OutlookMailTask(mail).performMailTask()

if __name__=='__main__':
    outlook=win32com.client .DispatchWithEvents("Outlook.Application",Handler_Class)
    pythoncom.PumpMessages()
    # print("Hello")