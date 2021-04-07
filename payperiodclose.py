import openpyxl
import pandas as pd
import re
import os
from datetime import datetime,timedelta
import win32com.client as win32
from tabulate import tabulate
from admin import newest, colclean
from emailautosend import mailthat

#TODO encapsulate code into main function
#TODO add function to handle npay502 files
#TODO anonymize by referencing a distlist file hosted elswhere
#TODO comment the code
#TODO add a reference to a json file with pay periods and date ranges for file naming
#TODO add a check for pay periods so that this is looking during the right dates


pd.options.display.float_format = '{:.0f}'.format

datedict={27:27, 9: 1, 10: 2, 11: 3, 12: 4, 13: 5,
 14: 6, 15: 7, 16: 8, 17: 9, 18: 10, 19: 11, 20: 12, 21: 13,22: 14, 23: 15,
 24: 16, 25: 17, 26: 18, 1: 19, 2: 20, 3: 21, 4: 22, 5: 23, 6: 24, 7: 25, 8: 26}
rdatedict={v:k for k,v in datedict.items()}


#x=[i for i in range(9,27)]
#x.extend([i for i in range(1,9)])
#y=[i for i in range(1,27)]
#ppdict={a:y[ia] for ia,a in enumerate(x)}

def data_collect(path,fname,itera=None):
    if itera:
        itera=itera
    else:
        itera=0
    frames=newest(path,fname,itera=itera)
    return(frames)  #return Itera number of filepaths    
    
def pp_collect(path):
    fname = "CrystalReportViewer"
    itera=2
    return(data_collect(path,fname,itera=itera)) 
    
def npay_collect(path):
    fname="NPAY502"
    itera=2
    return(data_collect(path,fname,itera=itera))

def convert_pp(df):
    df['pr']=df.pr.apply(lambda x: datedict[x])
    period=int(df.pr.max())
    period= rdatedict[period]
    return(period)

def rename_file(path,fname,df,path2=None):  #leaving option for renaming in other location
    if path2:
        path2=path2
    else:
        path2=path
    period=convert_pp(df)
    year=int(datetime.now().strftime("%Y"))[2:]
    #TODO fix this filename hardcoding
    if df[(df.title.isnull()==False)&(df.title.str.contains('Adj'))].shape[0] > 0:
        if df[(df.ps_emp_id.isnull()==True)].shape[0] >0:
            newpath = os.path.join(path2,str(f'aems pp {period} {year}.xls'))
            os.rename(os.path.join(path,fname),newpath)
    else:
        if df[(df.ps_emp_id.isnull()==True)].shape[0] >0:
            newpath = os.path.join(path2,str(f'pr pp {period} {year}.xls'))
            os.rename(os.path.join(path,fname),newpath)

def processing_npay(path,frames,path2=None):
    if path2:
        path2=path2
    else:
        path2=path
    npay=[]
    for i in frames:
        with open(i,'r') as f:
            npay.extend(f.readlines())
    return(npay)
#TODO add function here to create npay502 manually from the associated pp files.

def create_npay(path,npay):
    newpath=os.path.join(path,'NPAY502.txt')
    with open(newpath,'w') as f:
        f.writelines(npay)

def rename_npay(path):
    files=newest(path,'NPAY502',2)
    #since format is always NPAY502 YYYY-MM-DD.txt...
    file=[file for file in files if'NPAY502.txt' not in file][0]
    lastdate=file.split('\\')[-1].split(' ')[1].split('.')[0]
    lastdate=datetime.strptime(lastdate, "%Y-%m-%d")
    newdate=(lastdate+timedelta(days=14)).strftime("%Y-%m-%d")
    oldname=os.path.join(path,'NPAY502.txt')
    newname=os.path.join(path,("NPAY502 "+newdate+'.txt'))
    os.rename(oldname,newname)

def npaymain(path,path2=None):
    if path2:
        path2=path2
    else:
        path2=path
    npay=processing_npay(path,npay_collect(path))
    create_npay(path2,npay)
    rename_npay(path2)
#missin_n = missin_n[(missin_n.empl_id.isnull()==False)].astype({"empl_id": int})
def npaysend(path):
    obj=newest(path,'NPAY502')
    html= f'\n<html>\n<head>\n<p>Good Day, </p>\n<p> </p>\n<p>Attached please find the NPAY502 report.</p>\n<p> </p>\n<p>Best Regards,</p>\n<p>Shane Ayers</p>\n<p>Human Resources Information Systems Manager</p>\n<p>Office of Human Resources</p>\n<p>York College</p>\n<p>The City University of New York</p>\n</body></html>\n'
    text=html[0]+html[1:]
    for i in ['<p>','</p>','<html>','<head>','</body>','</html>']:
        text=text.replace(i,'')
    #TODO replace these hardcoded names with a json call and dict lookup
    to='University_Payroll_Interface_Processing@cuny.edu;'
    cc='adavis901@york.cuny.edu;'
    bcc=''
    subject=obj.split('\\')[-1].split('.')[0][:-4]
    #this is the part where you use the file both as attachment and as subject
    mailthat(subject,to=to,cc=cc,bcc=bcc,text=text,html=html,atch=obj)
    #plus custom string for send function
    #plus actual sending
    
    
def processing_pp(path,frames,path2=None):  #finds missing, renames files, returns dfs
    if path2:
        path2=path2
    else:
        path2=path
    dframes=[]
    missing_ns=[]
    fnames=[fname[len(path):] for fname in frames]
    for ix,i in enumerate(frames):
        df = colclean(pd.read_excel(i))
        df=convert_pp(df)
        dframes.append(df)
        missin_n=df[df.title == "xyz"]
        missin_n = missin_n.append(df[(df.ps_emp_id.isnull()==True)])
        missin_n=list(missin_n['name','title','empl_id'].to_records(index=False))
        missing_ns.extend(missin_n)
        rename_file(path,fnames[ix],df,path2=path2)
    df = pd.DataFrame(missing_ns, columns =['name','title','empl_id']) 
    df = df[(df.empl_id.isnull()==False)].astype({"empl_id": int})
    dframes.append(df)
    return(dframes) #a list of dataframes at least 2 items long

#TODO encapsulate this send function. Use class(h) letter file as example  
def sendfunc(df):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    #TODO read these from a dictionary stored in json
    sendlist='jamican@york.cuny.edu;hgordon@york.cuny.edu;adavis901@york.cuny.edu;eford1@york.cuny.edu'
    if df[(df.title.isnull()==False)&(df.title.str.contains('Adj'))].shape[0] > 0:
        sendlist=sendlist
    else:
        sendlist = sendlist+';bmajor@york.cuny.edu'
    mail.Subject = "Pay period close"
    
    mail.To=sendlist
    
    text = """{table}"""
    
    html = """
    <html>
    <head>
    <style>     
     table, th, td {{ border: 1px solid black; border-collapse: collapse; }}
      th, td {{ padding: 10px; }}
    </style>
    </head>
    <p>{table}<p>
    <p>Best Regards,</p>
    <p>Shane Ayers</p>
    <p>Acting Human Resources Information Systems Manager</p>
    <p>Office of Human Resources</p>
    <p>York College</p>
    <p>The City University of New York</p>
    </body></html>
    """
    
            
    # above line took every col inside csv as list
    try:
        text = text.format(table=tabulate(missin_n, headers=(list(missin_n.columns.values)), tablefmt="grid"))
        html = html.format(table=tabulate(missin_n, headers=(list(missin_n.columns.values)), tablefmt="html"))
    except:
        pass
    mail.Body = text
    mail.HTMLBody = html
    #To attach a file to the email (optional):
    
    attachment  = newpath
    mail.Attachments.Add(attachment)
    mail.Send()

path = "S:\\Downloads\\"     # Give the location of the file      
npaypay=r'Y:\Reports\NPAY files\Processed\2021'   

#ef main():
    