import pandas as pd
import xlwings as xw
import smtplib
import os, datetime
import json
import openpyxl
import re
import time
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate, formataddr
from email.header import Header
from email import encoders
import tkinter as tk
from tkinter import ttk

CRLF = "\r\n"
LARGE_FONT = ("Verdana", 12)

def log_conf():
    with open(os.path.join(os.getcwd(), 'pass.json'), 'r', encoding='utf-8') as f:
        ds = json.load(f)
    global login, password, path
    login = ds['login']['email']
    password = ds['login']['pass']
    path = ds['path']

def df_from_excel(path):
    book = xw.Book(path)
    book.save()
    # book.close()
    return pd.read_excel(path,header=0)

def open_exc(path):
    global wb_main, wb_trainer, wb_trainee, wb_cc, wb_trainroom, wb_eti
    wb_main = df_from_excel(path)
    wb_trainer = pd.read_excel(path, sheet_name="trainer")
    wb_trainee = pd.read_excel(path, sheet_name="trainee")
    wb_cc = pd.read_excel(path, sheet_name="CC")
    wb_trainroom = pd.read_excel(path, sheet_name="training_room")
    wb_eti = pd.read_excel(path, sheet_name="ETI")

def meeting_req(df):
    for event in df["event_code"].unique().tolist():
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
            attendees += list_room
        organizer = "ORGANIZER;CN=organiser:mailto:prameswari.kristal@nutrifood.co.id"
        fro = "prameswari.kristal@nutrifood.co.id"

        ddtstart = wb_main[wb_main['event_code'] == event]['event_date'].tolist()[0]
        ddtstart += datetime.timedelta(minutes=-(7*60))
        dtend = ddtstart + datetime.timedelta(minutes= wb_main[wb_main['event_code'] == event]['event_duration'].tolist()[0])
        dtstamp = datetime.datetime.now().strftime("%Y%m%dT%H%M%SZ")
        dtstart = ddtstart.strftime("%Y%m%dT%H%M%SZ")
        dtend = dtend.strftime("%Y%m%dT%H%M%SZ")

        body1 = ("Dear rekan-rekan,<br>" \
                "Mengundang rekan-rekan POK mengikuti:<br>" \
                "%(judul)s <br>" \
                "<br>" \
                "Hari, Tanggal: %(hari)s, %(tanggal)s <br>" \
                "Durasi: %(durasi)i jam <br>" \
                "Tempat: %(tempat)s <br>" \
                "Trainer: %(trainer)s <br>" \
                "Trainee:<br>" % {"judul": wb_main[wb_main['event_code'] == event]['event_name'].iloc[0],
                                "tanggal": str(wb_main[wb_main['event_code'] == event]['event_date'].iloc[0]),
                                "durasi": wb_main[wb_main['event_code'] == event]['event_duration'].iloc[0]/60,
                                "hari": wb_main[wb_main['event_code'] == event]['event_day'].iloc[0],
                                "tempat": wb_trainroom[wb_trainroom['event_code'] == event]['meeting_room'].iloc[0],
                                "trainer": wb_trainer[wb_trainer['event_code'] == event]['trainer_name'].iloc[0]})

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

        description = body1 + body2 + body3 +CRLF
        # LOC = wb_trainroom[wb_trainroom['event_code'] == event]["meeting_room"].unique().tolist()[0]
        LOC = ",".join(list_room)
        attendee = ""
        for att in attendees:
            attendee += "ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=ACCEPTED;RSVP=TRUE"+CRLF+" ;CN="+att+";X-NUM-GUESTS=0:"+CRLF+" mailto:"+att+CRLF
        for optt in list_optional:
            attendee += "ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=OPT-PARTICIPANT;PARTSTAT=ACCEPTED;RSVP=TRUE"+CRLF+" ;CN="+optt+";X-NUM-GUESTS=0:"+CRLF+" mailto:"+optt+CRLF
        ical = "BEGIN:VCALENDAR"+CRLF+"PRODID:pyICSParser"+CRLF+"VERSION:2.0"+CRLF+"CALSCALE:GREGORIAN"+CRLF
        ical +="METHOD:REQUEST"+CRLF+"BEGIN:VEVENT"+CRLF+"DTSTART:"+dtstart+CRLF+"DTEND:"+dtend+CRLF+"DTSTAMP:"+dtstamp+CRLF+organizer+CRLF
        ical += "UID:FIXMEUID"+dtstamp+CRLF
        ical += attendee+"CREATED:"+dtstamp+CRLF+description+"LAST-MODIFIED:"+dtstamp+CRLF+"LOCATION:"+LOC+CRLF+"SEQUENCE:0"+CRLF+"STATUS:CONFIRMED"+CRLF
        ical += "SUMMARY: "+wb_main[wb_main['event_code'] == event]['event_name'].unique()[0]+CRLF+"TRANSP:OPAQUE"+CRLF+"END:VEVENT"+CRLF+"END:VCALENDAR"+CRLF

        msg = MIMEMultipart('mixed')
        msg['Reply-To']=fro
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = wb_main[wb_main['event_code'] == event]['event_name'].unique()[0]
        msg['From'] = fro
        msg['To'] = ",".join(attendees)
        msg['CC'] = ",".join(list_optional)
        # msg['Location'] = wb_trainroom[wb_trainroom['event_code'] == event]['meeting_room'].tolist()[0]
        msg['Location'] = ",".join(list_room)

        part_email = MIMEText(description,"html")
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
        update_excel(idx, column="K")

        mailServer = smtplib.SMTP('smtp.gmail.com', 587)
        mailServer.ehlo()
        mailServer.starttls()
        mailServer.ehlo()
        mailServer.login(login, password)
        mailServer.sendmail(fro, attendees+list_optional, msg.as_string())
        mailServer.close()

        print(event, "-", wb_main[wb_main['event_code'] == event]['event_name'].iloc[0], "."*(40-len(wb_main[wb_main['event_code'] == event]['event_name'].iloc[0])), "done")

def update_excel(r, column):
    book = xw.Book(path)
    sht = book.sheets['main']
    cell = "{}".format(column) + str(r+2)
    sht.range(cell).value = 'done'
    book.save()
    # book.close()

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
        db_eti = pd.read_csv(r'database/db_eti.csv')
    except:
        db_main = pd.DataFrame(columns=['event_code', 'event_id_main', 'event_id_lain', 'event_name', 'event_day', 'event_date', 'event_number', 'event_duration', 'event_category', 'eval_training_code', 'meetreq_status', 'fdh_status', 'eti_status', 'orange_status', 'report_status', 'deadline'])
        db_trainer = pd.DataFrame(columns=['event_code', 'event_name', 'trainer_name', 'trainer_email'])
        db_trainee = pd.DataFrame(columns=['event_code', 'event_name', 'eval_training_code', 'survey_id', 'trainee_name', 'NIK', 'dept', 'trainee_email', 'presensi', 'absent_remark', 'trainee_status', 'nilai_post_test', 'eti_trainer_materi', 'eti_trainer_penampilan', 'eti_trainer_interaksi', 'eti_trainer_waktu', 'eti_materi_bobot', 'eti_materi_jelas', 'eti_materi_objective', 'eti_metode_objective', 'eti_organizer', 'eti_trainee_relevan', 'eti_trainee_manfaat', 'eti_essay_1', 'eti_essay_2', 'eti_essay_3'])
        db_cc = pd.DataFrame(columns=['dept', 'atasan', 'cc_1', 'cc_2', 'cc_3', 'cc_4', 'cc_5'])
        db_trainroom = pd.DataFrame(columns=['event_code', 'event_name', 'meeting_room', 'meeting_room_email', 'meeting_room_category'])
        db_eti = pd.DataFrame(columns=['eval_training_code', 'eti_trainer_materi', 'eti_trainer_penampilan', 'eti_trainer_interaksi', 'eti_trainer_waktu', 'eti_materi_bobot', 'eti_materi_jelas', 'eti_materi_objective', 'eti_metode', 'eti_organizer', 'eti_trainee_relevan', 'eti_trainee_manfaat', 'eti_essay_1', 'eti_essay_2', 'eti_essay_3'])
    finally:
        open_exc(path)
        db_main = pd.concat([db_main, wb_main], axis=0, ignore_index=True, sort=False)
        db_trainer = pd.concat([db_trainer, wb_trainer], axis=0, ignore_index=True, sort=False)
        db_trainee = pd.concat([db_trainee, wb_trainee], axis=0, ignore_index=True, sort=False)
        db_cc = pd.concat([db_cc, wb_cc], axis=0, ignore_index=True, sort=False)
        db_trainroom = pd.concat([db_trainroom, wb_trainroom], axis=0, ignore_index=True, sort=False)
        db_eti = pd.concat([db_eti, wb_eti], axis=0, ignore_index=True, sort=False)

        db_main.drop_duplicates(subset="event_code", keep='last', inplace=True)
        db_trainer.drop_duplicates(keep='first', inplace=True)
        db_trainee.drop_duplicates(keep='first', inplace=True)
        db_cc.drop_duplicates(keep='first', inplace=True)
        db_trainroom.drop_duplicates(keep='first', inplace=True)
        db_eti.drop_duplicates(keep='first', inplace=True)

        db_main.to_csv(r'database/db_main.csv', index=False)
        db_trainer.to_csv(r'database/db_trainer.csv', index=False)
        db_trainee.to_csv(r'database/db_trainee.csv', index=False)
        db_cc.to_csv(r'database/db_cc.csv', index=False)
        db_trainroom.to_csv(r'database/db_trainroom.csv', index=False)
        db_eti.to_csv(r'database/db_eti.csv', index=False)

def training_report(df):
    for event in df[(df['meetreq_status'] == 'done') & (df['report_status'].isnull()) & (df['fdh_status'] == 'done')]['event_code'].unique():
        list_to = []
        df_dept = wb_trainee[wb_trainee['event_code'] == event]
        for i in wb_cc[wb_cc['dept'].isin(df_dept['dept'].unique().tolist())].iloc[:,2:].values.tolist():
            for l in i:
                if l not in list_to and str(l) != 'nan':
                    list_to.append(l)
        
        list_cc = wb_trainer[wb_trainer['event_code'] == event].loc[:,'trainer_email'].values.tolist()
        
        # fro = "prameswari.kristal@nutrifood.co.id"
        fro = formataddr((str(Header('YDL', 'utf-8')), 'prameswari.kristal@nutrifood.co.id'))
        rt = ["prameswari.kristal@nutrifood.co.id", "ranilia.lestari@nutrifood.co.id"]

        new_df1 = wb_main[wb_main['event_code'] == event]
        new_df2 = wb_trainer[wb_trainer['event_code'] == event]
        new_df3 = wb_trainee[wb_trainee['event_code'] == event]
        new_df = new_df1.merge(new_df2, on=['event_code', 'event_name'], how='left')
        new_df = new_df.merge(new_df3, on=['event_code', 'event_name'], how='left')
        new_df = new_df.merge(wb_cc, on='dept', how='left')
        
        table_1 = new_df[['event_name', 'trainer_name', 'event_date']].set_index('event_name')
        table_1 = table_1.drop_duplicates(keep='last')
        table_1 = table_1.transpose()

        table_2 = new_df[['trainee_name', 'dept', 'atasan', 'nilai_post_test', 'presensi', 'absent_remark']]

        total_id_1 = 'totalID1'
        header_id_1 = 'headerID1'
        total_id_2 = 'totalID2'
        header_id_2 = 'headerID2'
        style_1_in_html = """<style>table#{total_table} {{color='black';font-size:13px; text-align:left; border:0.2px solid black; border-collapse:collapse; table-layout:fixed; padding:10px; width=100%; height="250"; text-align:left}} thead#{header_table} {{background-color: #fff645; color:#000000}}</style>""".format(total_table=total_id_1, header_table=header_id_1)
        style_2_in_html = """<style>table#{total_table} {{color='black';font-size:13px; text-align:center; border:0.2px solid black; border-collapse:collapse; table-layout:fixed; padding:10px; width=100%; height="250"; text-align:center}} thead#{header_table} {{background-color: #fff645; color:#000000}}</style>""".format(total_table=total_id_2, header_table=header_id_2)
        table_1_in_html = table_1.to_html()
        table_1_in_html = re.sub(r'<table', r'<table id=%s ' % total_id_1, table_1_in_html)
        table_1_in_html = re.sub(r'<thead', r'<thead id=%s ' % header_id_1, table_1_in_html)
        table_2_in_html = table_2.to_html(index=False)
        table_2_in_html = re.sub(r'<table', r'<table id=%s ' % total_id_2, table_2_in_html)
        table_2_in_html = re.sub(r'<thead', r'<thead id=%s ' % header_id_2, table_2_in_html)
        body1 = "<p>Dear rekan-rekan leader,<br/>Berikut adalah report dari pelaksanaan training:<br/></p>"
        body2 = style_1_in_html + table_1_in_html
        body3 = style_2_in_html + table_2_in_html
        body4 = "<p>Terimakasih,<br/><br/>Generated by pyYDL, any issue(s) please inform prameswari.kristal@nutrifood.co.id<br/></p>"
        body5 = "<br/>"
        body = body1 + body2 + body5 + body3 + body4

        msg = MIMEMultipart()
        msg['From'] = fro
        msg['To'] = ",".join(list_to)
        msg['Cc'] = ",".join(list_cc)
        msg['Subject'] = "Report {}".format(new_df['event_name'].unique()[0])
        msg.add_header('reply-to', ",".join(rt))
        msg.attach(MIMEText(body, 'html'))

        idx = wb_main[wb_main['event_code'] == event].index[0]
        update_excel(idx, column="O")

        mailServer = smtplib.SMTP('smtp.gmail.com', 587)
        mailServer.ehlo()
        mailServer.starttls()
        mailServer.ehlo()
        mailServer.login(login, password)
        mailServer.sendmail(msg['From'], list_to+list_cc, msg.as_string())
        mailServer.close()

        print("Report", event, "-", wb_main[wb_main['event_code'] == event]['event_name'].iloc[0], "."*(35-len(wb_main[wb_main['event_code'] == event]['event_name'].iloc[0])), "done")

def eti_report(df):
    for event in df[(df['meetreq_status'] == 'done') & (df['eti_status'].isnull()) & (df['fdh_status'] == 'done')]['event_code'].unique():
        list_to = wb_trainer[wb_trainer['event_code'] == event].loc[:,'trainer_email'].values.tolist()

        fro = 'prameswari.kristal@nutrifood.co.id'

        topic = wb_main[wb_main['event_code'] == event].loc[:,'event_name'].values.tolist()[0]
        hartang = str(wb_main[wb_main['event_code'] == event].loc[:,'event_day'].values.tolist()[0]) + " / " + str(wb_main[wb_main['event_code'] == event].loc[:,'event_date'].values.tolist()[0].strftime("%Y%m%d"))
        wakdar = str(wb_main[wb_main['event_code'] == event].loc[:,'event_date'].values.tolist()[0].strftime("T%H%M%SZ")) + " / " + str(wb_main[wb_main['event_code'] == event].loc[:,'event_duration'].values.tolist()[0]) + " menit"
        loc = wb_trainroom[wb_trainroom['event_code'] == event].loc[:,'meeting_room'].values[0]
        trainer = wb_trainer[wb_trainer['event_code'] == event].loc[:,'trainer_name'].values[0]
        sumtrainee = len(wb_trainee[wb_trainee['event_code'] == event].loc[:,'trainee_name'].values.tolist())

        code = wb_main[wb_main['event_code'] == event].loc[:,'eval_training_code'].values[0]
        mean_eti_trainer_materi = wb_eti[wb_eti['eval_training_code'] == code]['eti_trainer_materi'].mean()
        mean_eti_trainer_penampilan = wb_eti[wb_eti['eval_training_code'] == code]['eti_trainer_penampilan'].mean()
        mean_eti_trainer_interaksi = wb_eti[wb_eti['eval_training_code'] == code]['eti_trainer_interaksi'].mean()
        mean_eti_trainer_waktu = wb_eti[wb_eti['eval_training_code'] == code]['eti_trainer_waktu'].mean()
        mean_eti_materi_bobot = wb_eti[wb_eti['eval_training_code'] == code]['eti_materi_bobot'].mean()
        mean_eti_materi_jelas = wb_eti[wb_eti['eval_training_code'] == code]['eti_materi_jelas'].mean()
        mean_eti_materi_objective = wb_eti[wb_eti['eval_training_code'] == code]['eti_materi_objective'].mean()
        mean_eti_metode_objective = wb_eti[wb_eti['eval_training_code'] == code]['eti_metode_objective'].mean()
        mean_eti_organizer = wb_eti[wb_eti['eval_training_code'] == code]['eti_organizer'].mean()
        mean_eti_trainee_relevan = wb_eti[wb_eti['eval_training_code'] == code]['eti_trainee_relevan'].mean()
        mean_eti_trainee_manfaat = wb_eti[wb_eti['eval_training_code'] == code]['eti_trainee_manfaat'].mean()
        all_eti_essay_1 = "\n".join(wb_eti[wb_eti['eval_training_code'] == code]['eti_essay_1'])
        all_eti_essay_2 = "\n".join(wb_eti[wb_eti['eval_training_code'] == code]['eti_essay_2'])
        all_eti_essay_3 = "\n".join(wb_eti[wb_eti['eval_training_code'] == code]['eti_essay_3'])

        body1 = (f"Dear {trainer},<br>" \
                "Terimakasih sudah membawakan materi training. Berikut adalah Evaluasi Training Internal dari peserta training.<br><br>" \
                f"Topik\t\t: {topic} <br>" \
                f"Hari/tanggal\t\t: {hartang} <br>" \
                f"Waktu/durasi\t\t: {wakdar} <br>" \
                f"Tempat\t\t: {loc} <br>" \
                f"Jumlah peserta\t\t: {sumtrainee} <br>")
        body2 = f"""<table>
                    <tr>
                        <th>Aspek Trainer</th>
                        <th>Skala Nilai</th>
                    </tr>
                    <tr>
                        <td>Penguasaan materi</td>
                        <td>{mean_eti_trainer_materi}</td>
                    </tr>
                    <tr>
                        <td>Penampilan & body language</td>
                        <td>{mean_eti_trainer_penampilan}</td>
                    </tr>
                    <tr>
                        <td>Kemampuan interaksi</td>
                        <td>{mean_eti_trainer_interaksi}</td>
                    </tr>
                    <tr>
                        <td>Alokasi waktu training</td>
                        <td>{mean_eti_trainer_waktu}</td>
                    </tr>
                </table><br><br>"""
        body3 = f"""<table>
                    <tr>
                        <th>Aspek Materi</th>
                        <th>Skala Nilai</th>
                    </tr>
                    <tr>
                        <td>Bobot</td>
                        <td>{mean_eti_materi_bobot}</td>
                    </tr>
                    <tr>
                        <td>Kejelasan</td>
                        <td>{mean_eti_materi_jelas}</td>
                    </tr>
                    <tr>
                        <td>Kesesuaian materi dgn objective training</td>
                        <td>{mean_eti_materi_objective}</td>
                    </tr>
                </table><br><br>"""
        body4 = f"""<table>
                    <tr>
                        <th>Aspek Metode</th>
                        <th>Skala Nilai</th>
                    </tr>
                    <tr>
                        <td>Kesesuaian metode dgn objective training</td>
                        <td>{mean_eti_metode_objective}</td>
                    </tr>
                </table><br><br>"""
        body5 = f"""<table>
                    <tr>
                        <th>Aspek Organizer</th>
                        <th>Skala Nilai</th>
                    </tr>
                    <tr>
                        <td>Layout, suhu & kebersihan ruangan</td>
                        <td>{mean_eti_organizer}</td>
                    </tr>
                </table><br><br>"""
        body6 = f"""<table>
                    <tr>
                        <th>Aspek Trainee</th>
                        <th>Skala Nilai</th>
                    </tr>
                    <tr>
                        <td>Relevansi ke pekerjaan</td>
                        <td>{mean_eti_trainee_relevan}</td>
                    </tr>
                    <tr>
                        <td>Manfaat ke pekerjaan</td>
                        <td>{mean_eti_trainee_manfaat}</td>
                    </tr>
                    <tr>
                        <td>Poin-poin penting yg bermanfaat bagi pekerjaan</td>
                        <td>{all_eti_essay_1}</td>
                    </tr>
                    <tr>
                        <td>Poin-poin yg akan diimplementasikan dalam pekerjaan</td>
                        <td>{all_eti_essay_2}</td>
                    </tr>
                </table><br><br>"""
        body7 = f"""<table>
                    <tr>
                        <th>Usulan Perbaikan</th>
                    </tr>
                    <tr>
                        <td>{all_eti_essay_3}</td>
                    </tr>
                </table><br><br>"""

def email_training():
    tic = time.time()
    log_conf()
    open_exc(path)
    meeting_req(df=wb_main[wb_main['meetreq_status'] != "done"])
    extract_excel()
    create_db()
    toc = time.time()
    print("Email berhasil. Durasi proses: {0:.2f} detik".format(toc - tic))

def email_report_training():
    tic = time.time()
    log_conf()
    open_exc(path)
    training_report(df=wb_main)
    extract_excel()
    create_db()
    toc = time.time()
    print("Email report berhasil. Durasi proses: {0:.2f} detik".format(toc - tic))

class YDLapp(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        container.master.title("YDL Tools")

        self.frames = {}

        for F in (StartPage, changeConf, reportPage):
            frame = F(container, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        # ----
        self.show_frame(StartPage)

        menu = tk.Menu(container.master)
        container.master.config(menu=menu)

        file = tk.Menu(menu)
        file.add_command(label="Home", command = lambda: self.show_frame(StartPage))
        file.add_command(label="Report", command = lambda: self.show_frame(reportPage))
        file.add_command(label="Configuration", command = lambda: self.show_frame(changeConf))
        menu.add_cascade(label="File", menu=file)
    
    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()

class StartPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        label = tk.Label(self, text="Meeting Request\nJadwal Training\nKlik Send!", font=LARGE_FONT)
        label.pack(pady=10, padx=10)

        button1 = ttk.Button(self, text="Send", command=email_training)
        button1.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

class changeConf(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        tk.Label(self, text="\n\n").grid(row=0)
        tk.Label(self, text="\tEmail").grid(row=2)
        tk.Label(self, text="\tPassword").grid(row=3)
        tk.Label(self, text="\tPath").grid(row=4)

        email = tk.Entry(self)
        password = tk.Entry(self, show="*")
        path = tk.Entry(self)

        email.grid(row=2, column=1)
        password.grid(row=3, column=1)
        path.grid(row=4, column=1)

        def change_conf():
            with open(os.path.join(os.getcwd(), 'pass.json'), 'r', encoding='utf-8') as f:
                ds = json.load(f)
            
            ds['login']['email'] = email.get()
            ds['login']['pass'] = password.get()
            ds['path'] = path.get()

            with open(os.path.join(os.getcwd(), 'pass.json'), 'w', encoding='utf-8') as f:
                json.dump(ds, f)

            email.delete(0, tk.END)
            password.delete(0, tk.END)
            path.delete(0, tk.END)

        tk.Button(self, text='OK', command=change_conf).grid(row=5, column=1, sticky=tk.W, pady=4)

class reportPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        label = tk.Label(self, text='Report Training\n\nKlik Send!', font=LARGE_FONT)
        label.pack(pady=10, padx=10)

        button1 = ttk.Button(self, text='Send', command=lambda: email_report_training())
        button1.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

app = YDLapp()
app.geometry("400x300")
app.mainloop()