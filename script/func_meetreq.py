import pandas as pd
import smtplib
import getpass
import os, datetime
import json
import openpyxl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders

def meeting_req():
    for event in wb_main[wb_main['meetreq_status'] != "done"]["event_code"].unique().tolist():
        list_trainer = wb_trainer[wb_trainer['event_code'] == event]["trainer_email"].unique().tolist()
        list_trainee = wb_trainee[wb_trainee['event_code'] == event]["trainee_email"].unique().tolist()
        list_room = wb_trainroom[(wb_trainroom['event_code'] == event) & (wb_trainroom['meeting_room_category'] == 'system')]["meeting_room_email"].unique().tolist()
        list_optional = []
        for i in wb_cc[wb_cc['dept'].isin(wb_trainee['dept'].unique().tolist())].iloc[:,2:].values.tolist():
            for l in i:
                if l not in list_optional and str(l) != 'nan':
                    list_optional.append(l)
        attendees = list_trainer + list_trainee
        if len(list_room) != 0:
            attendees = attendees + list_room
        organizer = "ORGANIZER;CN=organiser:mailto:prameswari.kristal"+CRLF+"@nutrifood.co.id"
        fro = "prameswari.kristal@nutrifood.co.id"

        ddtstart = wb_main[wb_main['event_code'] == event]['event_date'].tolist()[0]
        dtend = ddtstart + datetime.timedelta(minutes= wb_main[wb_main['event_code'] == event]['event_duration'].tolist()[0])
        dtstamp = datetime.datetime.now().strftime("%Y%m%dT%H%M%SZ")
        dtstart = ddtstart.strftime("%Y%m%dT%H%M%SZ")
        dtend = dtend.strftime("%Y%m%dT%H%M%SZ")

        description = "DESCRIPTION: training invitation from YDL-Nutrifood"+CRLF
        attendee = ""
        for att in attendees:
            attendee += "ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-    PARTICIPANT;PARTSTAT=ACCEPTED;RSVP=TRUE"+CRLF+" ;CN="+att+";X-NUM-GUESTS=0:"+CRLF+" mailto:"+att+CRLF
        ical = "BEGIN:VCALENDAR"+CRLF+"PRODID:pyICSParser"+CRLF+"VERSION:2.0"+CRLF+"CALSCALE:GREGORIAN"+CRLF
        ical +="METHOD:REQUEST"+CRLF+"BEGIN:VEVENT"+CRLF+"DTSTART:"+dtstart+CRLF+"DTEND:"+dtend+CRLF+"DTSTAMP:"+dtstamp+CRLF+organizer+CRLF
        ical += "UID:FIXMEUID"+dtstamp+CRLF
        ical += attendee+"CREATED:"+dtstamp+CRLF+description+"LAST-MODIFIED:"+dtstamp+CRLF+"LOCATION:"+CRLF+"SEQUENCE:0"+CRLF+"STATUS:CONFIRMED"+CRLF
        ical += "SUMMARY:test "+ddtstart.strftime("%Y%m%d @ %H:%M")+CRLF+"TRANSP:OPAQUE"+CRLF+"END:VEVENT"+CRLF+"END:VCALENDAR"+CRLF

        eml_body = "Email body visible in the invite of outlook and outlook.com but not google calendar"
        msg = MIMEMultipart('mixed')
        msg['Reply-To']=fro
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = wb_main[wb_main['event_code'] == event]['event_name'].tolist()[0]
        msg['From'] = fro
        msg['To'] = ",".join(attendees)
        msg['Location'] = wb_trainroom[wb_trainroom['event_code'] == event]['meeting_room'].tolist()[0]

        part_email = MIMEText(eml_body,"html")
        part_cal = MIMEText(ical,'calendar;method=REQUEST')

        msgAlternative = MIMEMultipart('alternative')
        msg.attach(msgAlternative)

        ical_atch = MIMEBase('application/ics',' ;name="%s"'%("invite.ics"))
        ical_atch.set_payload(ical)
        encoders.encode_base64(ical_atch)
        ical_atch.add_header('Content-Disposition', 'attachment; filename="%s"'%("invite.ics"))

        # eml_atch = MIMEBase('text/plain','')
        # encoders.encode_base64(eml_atch)
        # eml_atch.add_header('Content-Transfer-Encoding', "")

        msgAlternative.attach(part_email)
        msgAlternative.attach(part_cal)

        mailServer = smtplib.SMTP('smtp.gmail.com', 587)
        mailServer.ehlo()
        mailServer.starttls()
        mailServer.ehlo()
        mailServer.login(login, password)
        mailServer.sendmail(fro, attendees, msg.as_string())
        mailServer.close()

with open('pass.json', 'r') as f:
    ds = json.load(f)

# email configuration
login = ds['login']['email']
password = ds['login']['pass']

CRLF = "\r\n"

wb_main = pd.read_excel("/Users/fahimhadimaula/Documents/F01 - YDL Generator/main_database_MU.xlsx", sheet_name="main")
wb_trainer = pd.read_excel("/Users/fahimhadimaula/Documents/F01 - YDL Generator/main_database_MU.xlsx", sheet_name="trainer")
wb_trainee = pd.read_excel("/Users/fahimhadimaula/Documents/F01 - YDL Generator/main_database_MU.xlsx", sheet_name="trainee")
wb_cc = pd.read_excel("/Users/fahimhadimaula/Documents/F01 - YDL Generator/main_database_MU.xlsx", sheet_name="CC")
wb_trainroom = pd.read_excel("/Users/fahimhadimaula/Documents/F01 - YDL Generator/main_database_MU.xlsx", sheet_name="training_room")

if __name__ == "__main__":
    meeting_req()