import win32com.client as win32


def TestSendMail():
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'mrushi@deloitte.com'
    mail.Subject = 'Track Followup'
    # mail.HTMLBody = '<h2>You are safe</h2><h3>Respond with subject Yes or No</h3><p>Thanks in advance</p><p>Rushi Mashru</p>'
    mail.Body = 'Testing: This mail is to track your follow up on, please respond back by voting Yes or No'
    mail.VotingOptions = "Yes;No"
    attachment = r'C:\\Users\\mrushi\\Desktop\\ImpFiles\\TrackIT\\TrackFollowup.xlsx'
    mail.Attachments.Add(attachment)
    mail.Send()


def SendMail(SrNo, PCEmail, PMOEmail, PMOEmail2, ProjectName, ClientName, PMOName, Subject, Body, IsHtml=False):
    outlook = win32.Dispatch('outlook.application')
    ccList = [PMOEmail2, PCEmail]
    mail = outlook.CreateItem(0)
    mail.To = PMOEmail
    mail.CC = ';'.join(ccList)
    mail.Subject = f'Sr:{SrNo} | TrackIT | {Subject} | {ClientName}'

    if(IsHtml):
        mail.HTMLBody = f'<p>Hi, {PMOName} - </p><p>Would you kindly confirm if the relevant project documents of the {ProjectName} engagement have been uploaded into Source.</p><p>{PMOName} please reply by using voting button.</p><p>If possible, a response within 2 weeks would be much appreciated.</p><p>{Body}</p>Many thanks,<p style="margin:0px;">Teresa</p><p>--</p><p style="margin:0px;">Teresa Palmieri</p><p style="margin:0px;">Engagement Review Program Manager | Service Excellence</p>'
    else:
        mail.Body = f'Testing: This mail is to track your follow up on {ProjectName} project, kindly respond back by voting Yes or No'

    mail.VotingOptions = "Yes;No"
    mail.Send()


def SendMailWithAttachment(To, SrNo, ProjectName, attachmentPath, IsHtml=False):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = To
    mail.Subject = f'Sr:{SrNo} Track Followup of {ProjectName}'

    if(IsHtml):
        mail.HTMLBody = f'<h2>Testing:</h2><h3>This email is to track your follow up on {ProjectName} project</h3></br><h3>Respond back by voting Yes or No</h3><p>Thanks in advance</p>Regards,<p>Rushi Mashru</p>'
    else:
        mail.Body = f'Testing: This mail is to track your follow up on {ProjectName} project, please respond back with Yes or No in subject'

    mail.VotingOptions = "Yes;No"
    attachment = attachmentPath
    mail.Attachments.Add(attachment)
    mail.Send()
