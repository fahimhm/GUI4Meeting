import pandas as pd
import xlwings as xw
import smtplib
import os, datetime
import json
import openpyxl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders

CRLF = "\r\n"
def df_from_excel(path):
    app = xw.App(visible=False)
    # book = app.books.open(path)
    book = xw.Book(path)
    book.save()
    book.close()
    app.kill()
    return pd.read_excel(path,header=0)

def open_exc():
    global wb_main, wb_trainer, wb_trainee, wb_cc, wb_trainroom
    wb_main = df_from_excel(path)
    wb_trainer = pd.read_excel(path, sheet_name="trainer")
    wb_trainee = pd.read_excel(path, sheet_name="trainee")
    wb_cc = pd.read_excel(path, sheet_name="CC")
    wb_trainroom = pd.read_excel(path, sheet_name="training_room")

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
        organizer = "ORGANIZER;CN=organiser:mailto:prameswari.kristal@nutrifood.co.id"
        fro = "prameswari.kristal@nutrifood.co.id"

        ddtstart = wb_main[wb_main['event_code'] == event]['event_date'].tolist()[0]
        dtend = ddtstart + datetime.timedelta(minutes= wb_main[wb_main['event_code'] == event]['event_duration'].tolist()[0])
        dtstamp = datetime.datetime.now().strftime("%Y%m%dT%H%M%SZ")
        dtstart = ddtstart.strftime("%Y%m%dT%H%M%SZ")
        dtend = dtend.strftime("%Y%m%dT%H%M%SZ")

        description = "DESCRIPTION: training invitation from YDL-Nutrifood"+CRLF
        LOC = wb_trainroom[wb_trainroom['event_code'] == event]["meeting_room"].unique().tolist()[0]
        attendee = ""
        for att in attendees:
            attendee += "ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-    PARTICIPANT;PARTSTAT=ACCEPTED;RSVP=TRUE"+CRLF+" ;CN="+att+";X-NUM-GUESTS=0:"+CRLF+" mailto:"+att+CRLF
        ical = "BEGIN:VCALENDAR"+CRLF+"PRODID:pyICSParser"+CRLF+"VERSION:2.0"+CRLF+"CALSCALE:GREGORIAN"+CRLF
        ical +="METHOD:REQUEST"+CRLF+"BEGIN:VEVENT"+CRLF+"DTSTART:"+dtstart+CRLF+"DTEND:"+dtend+CRLF+"DTSTAMP:"+dtstamp+CRLF+organizer+CRLF
        ical += "UID:FIXMEUID"+dtstamp+CRLF
        ical += attendee+"CREATED:"+dtstamp+CRLF+description+"LAST-MODIFIED:"+dtstamp+CRLF+"LOCATION:"+LOC+CRLF+"SEQUENCE:0"+CRLF+"STATUS:CONFIRMED"+CRLF
        ical += "SUMMARY:test "+ddtstart.strftime("%Y%m%d @ %H:%M")+CRLF+"TRANSP:OPAQUE"+CRLF+"END:VEVENT"+CRLF+"END:VCALENDAR"+CRLF

        body1 = ("Dear rekan-rekan,<br>" \
"Mengundang rekan-rekan POK mengikuti:<br>" \
"%(judul)s <br>" \
"<br>" \
"Hari, Tanggal: %(hari)s, %(tanggal)s <br>" \
"Durasi: %(durasi)i jam <br>" \
"Tempat: %(tempat)s <br>" \
"Trainer: %(trainer)s <br>" \
"Trainee:<br>" % {"judul": wb_main[wb_main['event_code'] == event]['event_name'][0],
                "tanggal": str(wb_main[wb_main['event_code'] == event]['event_date'][0]),
                "durasi": wb_main[wb_main['event_code'] == event]['event_duration'][0]/60,
                "hari": wb_main[wb_main['event_code'] == event]['event_day'][0],
                "tempat": wb_trainroom[wb_trainroom['event_code'] == event]['meeting_room'][0],
                "trainer": wb_trainer[wb_trainer['event_code'] == event]['trainer_name'][0]})

        forbody2 = wb_trainee[wb_trainee['event_code'] == event][['trainee_name', 'dept']]
        forbody2.drop_duplicates(keep='first', inplace=True)
        body2 = ""
        for i in forbody2.index:
            listpeserta = "-- " + forbody2.trainee_name[i] + " - " + forbody2.dept[i] + "<br>"
            body2 += listpeserta

        body3 = "<br><br>Mohon bantuan rekan-rekan untuk dapat hadir tepat waktu, mengisi evaluasi training maupun posttest (bila ada).<br>" \
        "Terimakasih ya<br>" \
        "Salam<br><br>" \
        "Tim YDL-Nutrifood"

        eml_body = body1 + body2 + body3
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

        idx = wb_main[wb_main['event_code'] == event].index[0]
        update_excel(idx)

        mailServer = smtplib.SMTP('smtp.gmail.com', 587)
        mailServer.ehlo()
        mailServer.starttls()
        mailServer.ehlo()
        mailServer.login(login, password)
        mailServer.sendmail(fro, attendees, msg.as_string())
        mailServer.close()

def update_excel(r):
    workbook = openpyxl.load_workbook(path)
    worksheet = workbook['main']
    mycell = worksheet.cell(row=(r+2), column=11)
    mycell.value = 'done'
    workbook.save(path)

def extract_excel():
    global wb_trainer, wb_trainee, wb_trainroom
    wb_trainer = wb_trainer.loc[:,['event_code', 'trainer_name', 'trainer_email']]
    exc = wb_main.merge(wb_trainer, on='event_code', how='outer')
    wb_trainee = wb_trainee.loc[:, ['event_code', 'survey_id', 'trainee_name', 'NIK', 'dept', 'trainee_email', 'presensi', 'absent_remark', 'trainee_status', 'nilai_post_test', 'eti_trainer_materi', 'eti_trainer_penampilan', 'eti_trainer_interaksi', 'eti_trainer_waktu', 'eti_materi_bobot', 'eti_materi_jelas', 'eti_materi_objective', 'eti_metode_objective', 'eti_organizer', 'eti_trainee_relevan', 'eti_trainee_manfaat', 'eti_essay_1', 'eti_essay_2', 'eti_essay_3']]
    exc = exc.merge(wb_trainee, on='event_code', how='outer')
    exc = exc.merge(wb_cc, on='dept', how='left')
    wb_trainroom = wb_trainroom.loc[:, ['event_code', 'meeting_room', 'meeting_room_email', 'meeting_room_category']]
    exc = exc.merge(wb_trainroom, on='event_code', how='left')
    with pd.ExcelWriter("extract_all.xlsx") as writer:
        exc.to_excel(writer, sheet_name='all', index=False)

def create_db():
    try:
        db_main = pd.read_csv(r'database/db_main.csv')
        db_trainer = pd.read_csv(r'database/db_trainer.csv')
        db_trainee = pd.read_csv(r'database/db_trainee.csv')
        db_cc = pd.read_csv(r'database/db_cc.csv')
        db_trainroom = pd.read_csv(r'database/db_trainroom.csv')
    except:
        db_main = pd.DataFrame(columns=['event_code', 'event_id_main', 'event_id_lain', 'event_name', 'event_day', 'event_date', 'event_number', 'event_duration', 'event_category', 'eval_training_code', 'meetreq_status', 'fdh_status', 'eti_status', 'orange_status', 'report_status', 'deadline'])
        db_trainer = pd.DataFrame(columns=['event_code', 'event_name', 'trainer_name', 'trainer_email'])
        db_trainee = pd.DataFrame(columns=['event_code', 'event_name', 'eval_training_code', 'survey_id', 'trainee_name', 'NIK', 'dept', 'trainee_email', 'presensi', 'absent_remark', 'trainee_status', 'nilai_post_test', 'eti_trainer_materi', 'eti_trainer_penampilan', 'eti_trainer_interaksi', 'eti_trainer_waktu', 'eti_materi_bobot', 'eti_materi_jelas', 'eti_materi_objective', 'eti_metode_objective', 'eti_organizer', 'eti_trainee_relevan', 'eti_trainee_manfaat', 'eti_essay_1', 'eti_essay_2', 'eti_essay_3'])
        db_cc = pd.DataFrame(columns=['dept', 'atasan', 'cc_1', 'cc_2', 'cc_3', 'cc_4', 'cc_5'])
        db_trainroom = pd.DataFrame(columns=['event_code', 'event_name', 'meeting_room', 'meeting_room_email', 'meeting_room_category'])
    finally:
        open_exc()
        db_main = pd.concat([db_main, wb_main], axis=0, ignore_index=True, sort=False)
        db_trainer = pd.concat([db_trainer, wb_trainer], axis=0, ignore_index=True, sort=False)
        db_trainee = pd.concat([db_trainee, wb_trainee], axis=0, ignore_index=True, sort=False)
        db_cc = pd.concat([db_cc, wb_cc], axis=0, ignore_index=True, sort=False)
        db_trainroom = pd.concat([db_trainroom, wb_trainroom], axis=0, ignore_index=True, sort=False)

        db_main.drop_duplicates(keep='first', inplace=True)
        db_trainer.drop_duplicates(keep='first', inplace=True)
        db_trainee.drop_duplicates(keep='first', inplace=True)
        db_cc.drop_duplicates(keep='first', inplace=True)
        db_trainroom.drop_duplicates(keep='first', inplace=True)

        db_main.to_csv(r'database/db_main.csv', index=False)
        db_trainer.to_csv(r'database/db_trainer.csv', index=False)
        db_trainee.to_csv(r'database/db_trainee.csv', index=False)
        db_cc.to_csv(r'database/db_cc.csv', index=False)
        db_trainroom.to_csv(r'database/db_trainroom.csv', index=False)

def log_conf():
    with open('pass.json', 'r') as f:
        ds = json.load(f)
    global login, password, path
    login = ds['login']['email']
    password = ds['login']['pass']
    path = ds['path']

if __name__ == "__main__":
    log_conf()
    open_exc()
    meeting_req()
    extract_excel()
    create_db()