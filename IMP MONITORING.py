from email import message
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
import os
from os import path
import time
from pathlib import Path
from datetime import datetime
import smtplib
from datetime import date



###############################Capturing Date Segment-1
current_date = time.strftime("%Y-%m-%d")
print("Date Extracted and Formatted:", current_date)

############################### Validating Star Reports Segment-2

star_reports_path = Path(r"\\bnacorp01sf\star_actuate_reporting\Service Center Management Reports\Service Accounting\GL Activity\Daily")
star_reports_flag=['Failure','Success']
flag=0
bg='#FF0000'  
star_reports_files=list(star_reports_path.iterdir())
star_reports_latest_file = max(star_reports_files, key=os.path.getmtime)    
#print(star_reports_latest_file)
star_reports_date=datetime.fromtimestamp(star_reports_latest_file.stat().st_ctime)
#print("date:",str(star_reports_date.date()))
current_date = time.strftime("%Y-%m-%d")
#print("current_date:",str(current_date))
if str(star_reports_date.date()) == str(current_date):
    flag=1
    bg='#008000'
    #print(star_reports_flag[flag])

############################### Validating Inbound Files ( VRI , VSU, INV ) Segment-3

Inbound_path = Path(r"\\pdw01k02ap01a\bbyapps\bby-repl-data\Delete")

Inbound_path_file_names=['VSU', 'VRI', 'RGM', 'INV']
Inbound_path_files=list(Inbound_path.iterdir())
Inbound_path_files = Inbound_path_files[::-1]
Inbound_path_file_updated_dates={}
for Inbound_path_file in Inbound_path_files:
    Inbound_path_file_name = str(Inbound_path_file)[48:51]
    if Inbound_path_file_name in Inbound_path_file_names:
        Inbound_path_creation_time = datetime.fromtimestamp(Inbound_path_file.stat().st_ctime)
        Inbound_path_file_updated_dates[Inbound_path_file_name]=str(Inbound_path_creation_time.date())
        Inbound_path_file_names.remove(Inbound_path_file_name)
        if Inbound_path_file_names==[]:
            break


############################### Validating Outbound Files (RGM ) Segment-4
        
Outbound_path = Path(r"\\pdw01k02ap01a\bbyapps\bby-repl-data\Export\Archive")
Outbound_path_files=list(Outbound_path.iterdir())
Outbound_path_files = Outbound_path_files[::-1]
Outbound_path_file_names=['RGM']
Outbound_path_file_updated_dates={}
for Outbound_path_file in Outbound_path_files:
    Outbound_path_file_name = str(Outbound_path_file)[53:56]
    if Outbound_path_file_name in Outbound_path_file_names:
        Outbound_path_creation_time = datetime.fromtimestamp(Outbound_path_file.stat().st_ctime)
        Outbound_path_file_updated_dates[Outbound_path_file_name]=str(Outbound_path_creation_time.date())
        Outbound_path_file_names.remove(Outbound_path_file_name)
        if Outbound_path_file_names==[]:
            break

#print(Outbound_path_file_updated_dates)

###############################Outlook mail Segment -5
messageObject = MIMEMultipart()
fromEmail = ['A3048773@bestbuy.com']
toEmail = ['A3048773@bestbuy.com']
cc = ['A3048773@bestbuy.com']
    
messageObject['From'] = fromEmail
messageObject['To'] = ', '.join(toEmail)
messageObject['Cc'] = ', '.join(cc)
messageObject['Subject'] = "IMP MONITORING STATUS "+current_date+" "
messageObject['X-Priority'] = '2'


html1 = "<!DOCTYPE html><html><head><style>table{border-collapse: collapse;}td,th{border: 1px solid; padding: 4px; border-collapse: collapse;}</style></head><body><p style='font-size:100%;'>Hi All,</br>Since the flow stopped and Job is not processing any records, hence we are removing the K02ND104 job stream from the monitoring sheet.</br>Please find the monitoring status :</br>Evening Monitoring</p><table cellpadding='0' cellspacing='0'><tr><th style='background-color:#b3ccff'>Description</th><th style='background-color:#b3ccff'>Status</th></tr><tr><td>STAR Reports</td><td style='background-color:"+bg+"'>"+star_reports_flag[flag]+"</td></tr><tr><td>Files sent to NP</td><td>Inbound (last received):</br>VSU :"+Inbound_path_file_updated_dates['VSU']+" </br>INV :"+Inbound_path_file_updated_dates['INV']+" </br>RGM :"+Outbound_path_file_updated_dates['RGM']+" </br></td></tr><tr><td>VRI files monitoring</td><td>VRI :"+Inbound_path_file_updated_dates['VRI']+" </td></td></table></body></html>"


img=html1+'<p>Thanks & Regards,</p><p>Thanks and Regards STAR ASM</p>'

msgText = MIMEText(img, 'html')
    
messageObject.attach(msgText)
   
   
try:
    server  = smtplib.SMTP('mail10.bestbuy.com')
    server.ehlo()
    server.sendmail(fromEmail,(toEmail+cc),messageObject.as_string())
    server.close()
except Exception as e:
    print(e)


print("mail sent")


































