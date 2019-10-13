import ManipulateExcelFile as MEF
import SendMailThroughOutlook as outlook
import datetime as dt
from termcolor import colored

currentDate = dt.datetime.now().strftime("%m/%d/%Y")
EmailNotSent = 'No'
EmailSent = 'Yes'
RespondedNo = 'No'
RespondedYes = 'Yes'


def TakeFollowUp():
    AdjustDate()
    CheckAndSendEmail()


def TrackFollowUp():
    TrackList = MEF.GetListFromExcel()

    for item in TrackList:
        if(item.SrNo != None):
            if(currentDate == dt.datetime.strftime(item.FollowUpDate, '%m/%d/%Y') and item.EmailSentBySystem == EmailSent and (item.ResponseFromReceiver == None or item.ResponseFromReceiver == RespondedNo)):
                Max_Row = int(MEF.GetMaxRow())
                anticipatedDate = item.FollowUpDate + dt.timedelta(days=3)
                DefaultersList = []
                DefaultersList.append(Max_Row + 1)
                DefaultersList.append(dt.datetime.strptime(
                    anticipatedDate.strftime("%m/%d/%Y"), "%m/%d/%Y").date())
                DefaultersList.append(item.PCName)
                DefaultersList.append(item.PCEmail)
                DefaultersList.append(item.PMOName)
                DefaultersList.append(item.PMOEmail)
                DefaultersList.append(item.PMONameCont2)
                DefaultersList.append(item.PMOEmail)
                DefaultersList.append(item.ClientName)
                DefaultersList.append(item.ProjectNameWBS)
                DefaultersList.append(item.ProjectNameCond2)
                DefaultersList.append('No')
                DefaultersList.append('NA')
                DefaultersList.append(item.Subject)
                DefaultersList.append(item.Body)
                DefaultersList.append('NA')
                DefaultersList.append('NA')
                DefaultersList.append('NA')
                DefaultersList.append('NA')
                DefaultersList.append('NA')
                DefaultersList.append(item.SrNo)
                MEF.AddNewEntryToExcel(DefaultersList)
                print(colored("\t\t" + item.PCEmail + " of " +
                              item.ProjectNameWBS + " has missed the dead line.", "red"))
                print(colored("\t\t" + anticipatedDate.strftime("%m/%d/%Y") +
                              " is next anticipated date for reminder.", "yellow"), "\n")


def GetResponseOfReceiverFromRefSrNo(RefSrNo):
    if(RefSrNo):
        oTrackIT = MEF.GetAllColumnsByRow(RefSrNo)
        return oTrackIT.ResponseFromReceiver

    return None


def AdjustDate():
    FollowUpList = MEF.GetListFromExcel()
    for item in FollowUpList:
        if(item.SrNo != None and item.PCEmail and item.PMOEmail and dt.datetime.now().date() > item.FollowUpDate.date()):
            if(item.EmailSentBySystem == EmailNotSent):
                if((item.RefSrNo == None and item.ResponseFromReceiver == None) or (item.RefSrNo and item.ResponseFromReceiver != RespondedYes)):
                    date = dt.datetime.strptime(currentDate, "%m/%d/%Y").date()
                    MEF.UpdateFollowUpDate(item.SrNo, 2, date)


def CheckAndSendEmail():
    FollowUpList = MEF.GetListFromExcel()
    for item in FollowUpList:
        if(item.SrNo != None and item.PCEmail and item.PMOEmail):
            # print(item.SrNo, item.PCEmail, item.PMOEmail,
            #       item.ProjectName, item.PCName, item.Subject, item.Body, item.EmailSentBySystem)
            if(currentDate == dt.datetime.strftime(item.FollowUpDate, '%m/%d/%Y') and item.EmailSentBySystem == EmailNotSent):
                if(item.RefSrNo == None):
                    print(colored('\t\tEmail sent to ' + item.PMOEmail +
                                  ' and ' + item.PMOEmail2, "green"))
                    outlook.SendMail(item.SrNo, item.PCEmail, item.PMOEmail, item.PMOEmail2, item.ProjectNameWBS,
                                     item.ClientName, item.PMOName, item.Subject, item.Body, IsHtml=True)
                    MEF.UpdateEmailSentBySystemStatus(item.SrNo, 12)
                elif(GetResponseOfReceiverFromRefSrNo(item.RefSrNo) != RespondedYes):
                    print(colored('\t\tEmail sent to ' + item.PMOEmail +
                                  ' and ' + item.PMOEmail2, "yellow"))
                    outlook.SendMail(item.SrNo, item.PCEmail, item.PMOEmail, item.PMOEmail2, item.ProjectNameWBS,
                                     item.ClientName, item.PMOName, item.Subject, item.Body, IsHtml=True)
                    MEF.UpdateEmailSentBySystemStatus(item.SrNo, 12)
                else:
                    MEF.UpdateResponseFromReceiverStatus(item.SrNo, 13, 'Yes')
