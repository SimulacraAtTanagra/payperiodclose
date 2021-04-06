from re import search
import win32com.client as win32
from tabulate import tabulate
import pythoncom

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
    
def mailthat(subject,to=None,cc=None,bcc=None,acc=None,text=None,html=None,atch=None,disp=None,temp=None,recp=None,deli=None):
    """
    This function has only subject as the required argument. 
    All other arguments are optional. to,cc, and bcc are self-explanatory.
    acc gives the option to use something other than the default account.
    text and html are self explanatory but text is not required if html is provided.
    atch is the full file location of any desired attachment(s)
    disp is option to open the e-mails for display *instead* of than sending.
    temp is an optional argument to work from a template or create a new item
    temp is file location of template.
    recp is optional read receipts, anything can go in this field
    deli is optional delivery receipts, anything can go in this field
    """
    
    outlook = win32.Dispatch('outlook.application')
    if temp:
        mail=outlook.CreateItemFromTemplate(temp)
    else:
        mail = outlook.CreateItem(0)
    if to:
        mail.To = to
    if cc:
        mail.CC = cc
    if bcc:
        mail.BCC = bcc
    mail.Subject = subject
    if acc:
        mail.SendUsingAccount = acc
    if recp:
        mail.ReadReceiptRequested = True
    if deli:
        mail.OriginatorDeliveryReportRequested = True       
    if not text:
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
    #to display the e-mail on screeen rather than sending
    if disp:
        mail.Display(False)
    else:
        mail.Send()
