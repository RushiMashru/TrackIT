import win32com.client
import win32timezone
import ctypes
import pythoncom
import re
import time
import psutil
import ManipulateExcelFile as MEF
import datetime as dt
from termcolor import colored
import colorama

colorama.init()

print(colored("\n\tTrackIT: Monitoring Outlook mail's...", "green"), "\n")


class Handler_Class(object):
    def __init__(self):
        try:
            f = open('eLog.txt', 'a+')
            try:
                self.SUBJECT = '| TrackIT |'
                inbox = self.Application.GetNamespace(
                    "MAPI").GetDefaultFolder(6)
                messages = inbox.Items
                for message in messages:
                    try:
                        if self.check_subject(message.Subject, self.SUBJECT):
                            votingResponse = message.VotingResponse
                            trackID = self.get_id(message.Subject)

                            print("\t", colored('Date: ', 'green'),
                                  dt.datetime.strftime(message.ReceivedTime, '%m/%d/%Y'))
                            print("\t", colored('Subject: ', 'green'),
                                  message.Subject)
                            print("\t", colored('Voted:', 'green'),
                                  votingResponse, "\n")

                            MEF.UpdateResponseFromReceiverStatus(
                                trackID, 13, votingResponse)
                    except Exception as e:
                        f.write(str(
                            dt.datetime.now()) + " [email_inner_logic_exception]:\n\t" + str(e) + "\n\n\n")

            except Exception as e:
                f.write(str(dt.datetime.now()) + ":\n\t" + str(e) + "\n\n\n")

        finally:
            f.close()

    def OnQuit(self):
        ctypes.windll.user32.PostQuitMessage(0)

    def OnNewMailEx(self, receivedItemsIDs):
        try:
            f = open('eLog.txt', 'a+')
            self.SUBJECT = '| TrackIT |'
            for ID in receivedItemsIDs.split(","):
                try:
                    mail = self.Session.GetItemFromID(ID)
                    subject = mail.Subject
                    if self.check_subject(subject, self.SUBJECT):
                        votingResponse = mail.VotingResponse
                        trackID = self.get_id(subject)

                        print("\t", colored('Subject:', 'green'), subject)
                        print("\t", colored('Voted: ', 'green'),
                              votingResponse, "\n")

                        MEF.UpdateResponseFromReceiverStatus(
                            trackID, 13, votingResponse)
                    try:
                        print('command', command)
                    except:
                        pass
                except Exception as e:
                    f.write(str(dt.datetime.now()) +
                            " [new_email_inner_logic_exception]:\n\t" + str(e) + "\n\n\n")
        finally:
            f.close()

    def check_subject(self, subjectStr, matchingStr):
        if (subjectStr.find(matchingStr) == -1):
            return False
        else:
            return True

    def get_id(self, subjectStr):
        result = re.findall(r'\d+', subjectStr)
        trackId = int(result[0])
        #print('trackId', trackId)
        return trackId


def check_outlook_open():
    list_process = []
    for pid in psutil.pids():
        p = psutil.Process(pid)
        list_process.append(p.name())

    if 'OUTLOOK.EXE' in list_process:
        return True
    else:
        return False


while True:
    try:
        outlook_open = check_outlook_open()
    except:
        outlook_open = False

    if outlook_open == True:
        outlook = win32com.client.DispatchWithEvents(
            "Outlook.Application", Handler_Class)
        pythoncom.PumpMessages()

    time.sleep(10)
