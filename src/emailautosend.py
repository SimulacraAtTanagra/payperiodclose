from re import search
import win32com.client as win32
from tabulate import tabulate
def getemail(search_string):
    try:    
        search_string = str(search_string)
        outlook = win32.Dispatch('outlook.application')
        gal = outlook.Session.GetGlobalAddressList()
        entries = gal.AddressEntries
        ae = entries[search_string]
        email_address = None
        
        if search(f'{search_string}$',str(ae)) != None:
           pass
        else:
           return('')
        
        if 'EX' == ae.Type:
            eu = ae.GetExchangeUser()
            email_address = eu.PrimarySmtpAddress
           
        
        if 'SMTP' == ae.Type:
            email_address = ae.Address
        return(email_address)
    except:
        return('')

def html_to_text(html):
    text=html[0]+html[1:]
    for i in ['<p>','</p>','<html>','<head>','</body>','</html>']:
        text=text.replace(i,'')
    return(text)

def mailthis(recipientlist,cc, df, subject,obj):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipientlist
    mail.Cc = cc
    mail.Subject = subject
    
    
    # To attach a file to the email (optional):
    #attachment  = "Path to the attachment"
    #mail.Attachments.Add(attachment)
    text = """
    Good Day,
    
    
    
    {table}
    
    Best Regards,
    Shane Ayers
    Acting Human Resources Information Systems Manager
    Office of Human Resources
    York College
    The City University of New York"""
    
    html = """
    <html>
    <head>
    <style>     
     table, th, td {{ border: 1px solid black; border-collapse: collapse; }}
      th, td {{ padding: 10px; }}
    </style>
    </head>
    <body><p>Good Day,</p>
    <p></p>
    {table}
    <p>Best Regards,</p>
    <p>Shane Ayers</p>
    <p>Acting Human Resources Information Systems Manager</p>
    <p>Office of Human Resources</p>
    <p>York College</p>
    <p>The City University of New York</p>
    </body></html>
    """
    
    # above line took every col inside csv as list
    text = text.format(table=tabulate(df, headers=(list(df.columns.values)), tablefmt="grid"))
    html = html.format(table=tabulate(df, headers=(list(df.columns.values)), tablefmt="html"))
    mail.Body = text
    mail.HTMLBody = html
    if obj!='':
        mail.Attachments.Add(obj)
    mail.Send()
    
def mailthat(subject,to=None,cc=None,bcc=None,text=None,html=None,atch=None):
    
    outlook = win32.Dispatch('outlook.application')
#    oacctouse = acc
#    for oacc in outlook.Session.Accounts:
#        if oacc.SmtpAddress == "humanresources@york.cuny.edu":
#            oacctouse = oacc
#            break
    mail = outlook.CreateItem(0)
    if to:
        mail.To = to
    if cc:
        mail.CC = cc
    if bcc:
        mail.BCC = bcc
    mail.Subject = subject
#    mail.SendUsingAccount = acc
    mail.ReadReceiptRequested = True
#    if oacctouse:
#        mail._oleobj_.Invoke(*(64209, 0, 8, 0, oacctouse))  # Msg.SendUsingAccount = oacctouse

    mail.OriginatorDeliveryReportRequested = True       
    if not text:
        if html:
            text=html_to_text(html)
        else:
            text = """Good Day,
            
            Attached find your reappointment letter for January. Please read it completely before signing, indicating either acceptance of the reappointment or denial of same, and return via e-mail as an attachment, preferably with the original filename or with your name in the filename. You may return the signed letter to Ms. Annie Jackson if you are a College Assistant or to Ms. Marilyn Williams if you are another classified hourly title. Please note that opening this document in a web browser like Chrome may display it without details such as your Name, Rate, or Title. Please open in Adobe for best results. 
            
            Best Regards, 
            Shane Ayers
            
            Human Resources Information Systems Manager
            Office of Human Resources
            York College
            The City University of New York
            """
    if not html:
        html = """
        <html>
        <head>
        <p> Good Afternoon, </p>
        <p> </p>
        <p>Attached find your reappointment letter for January. Please read it completely before signing, indicating either acceptance of the reappointment or denial of same, and return via e-mail as an attachment, preferably with the original filename or with your name in the filename. You may return the signed letter to Ms. Annie Jackson if you are a College Assistant or to Ms. Marilyn Williams if you are another classified hourly title.</p>
        <p>Please note that opening this document in a web browser like Chrome may display it without details such as your Name, Rate, or Title. Please open in Adobe for best results.</p>
        <p> </p>
        <p>Best Regards,</p>
        <p>Shane Ayers</p>
        <p>Human Resources Information Systems Manager</p>
        <p>Office of Human Resources</p>
        <p>York College</p>
        <p>The City University of New York</p>
        </body></html>
        """
    
                
    mail.Body = text
    mail.HTMLBody = html
    #To attach a file to the email (optional):
    if atch:
        mail.Attachments.Add(atch)
    mail.Send()
