import datetime as dt


class TrackFollowUpModel(object):

    def __init__(self, SrNo, FollowUpDate, PCName, PCEmail, PMOName, PMOEmail, PMONameCont2, PMOEmail2, ClientName, ProjectNameWBS, ProjectNameCond2, EmailSentBySystem, ResponseFromReceiver, Subject, Body, UserFollowUps, UserNotes, DocUploadCompleted, ConfirmedBy, UserSanityCheck, RefSrNo):
        self.SrNo = SrNo
        self.FollowUpDate = FollowUpDate
        self.PCName = PCName
        self.PCEmail = PCEmail
        self.PMOName = PMOName
        self.PMOEmail = PMOEmail
        self.PMONameCont2 = PMONameCont2
        self.PMOEmail2 = PMOEmail2
        self.ClientName = ClientName
        self.ProjectNameWBS = ProjectNameWBS
        self.ProjectNameCond2 = ProjectNameCond2
        self.EmailSentBySystem = EmailSentBySystem
        self.ResponseFromReceiver = ResponseFromReceiver
        self.Subject = Subject
        self.Body = Body
        self.UserFollowUps = UserFollowUps
        self.UserNotes = UserNotes
        self.DocUploadCompleted = DocUploadCompleted
        self.ConfirmedBy = ConfirmedBy
        self.UserSanityCheck = UserSanityCheck
        self.RefSrNo = RefSrNo
