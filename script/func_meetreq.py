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
    global wb_main, wb_trainer, wb_trainee, wb_cc, wb_trainroom
    wb_main = df_from_excel(path)
    wb_trainer = pd.read_excel(path, sheet_name="trainer")
    wb_trainee = pd.read_excel(path, sheet_name="trainee")
    wb_cc = pd.read_excel(path, sheet_name="CC")
    wb_trainroom = pd.read_excel(path, sheet_name="training_room")

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
    book = xw.Book(path)
    sht = book.sheets['main']
    cell = "K" + str(r+2)
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
    except:
        db_main = pd.DataFrame(columns=['event_code', 'event_id_main', 'event_id_lain', 'event_name', 'event_day', 'event_date', 'event_number', 'event_duration', 'event_category', 'eval_training_code', 'meetreq_status', 'fdh_status', 'eti_status', 'orange_status', 'report_status', 'deadline'])
        db_trainer = pd.DataFrame(columns=['event_code', 'event_name', 'trainer_name', 'trainer_email'])
        db_trainee = pd.DataFrame(columns=['event_code', 'event_name', 'eval_training_code', 'survey_id', 'trainee_name', 'NIK', 'dept', 'trainee_email', 'presensi', 'absent_remark', 'trainee_status', 'nilai_post_test', 'eti_trainer_materi', 'eti_trainer_penampilan', 'eti_trainer_interaksi', 'eti_trainer_waktu', 'eti_materi_bobot', 'eti_materi_jelas', 'eti_materi_objective', 'eti_metode_objective', 'eti_organizer', 'eti_trainee_relevan', 'eti_trainee_manfaat', 'eti_essay_1', 'eti_essay_2', 'eti_essay_3'])
        db_cc = pd.DataFrame(columns=['dept', 'atasan', 'cc_1', 'cc_2', 'cc_3', 'cc_4', 'cc_5'])
        db_trainroom = pd.DataFrame(columns=['event_code', 'event_name', 'meeting_room', 'meeting_room_email', 'meeting_room_category'])
    finally:
        open_exc(path)
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

def email_training():
    log_conf()
    open_exc(path)
    meeting_req(df=wb_main[wb_main['meetreq_status'] != "done"])
    extract_excel()
    create_db()
    print("Email berhasil")

class YDLapp(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        container.master.title("YDL Tools")

        self.frames = {}

        for F in (StartPage, changeConf):
            frame = F(container, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        # ----
        self.show_frame(StartPage)

        menu = tk.Menu(container.master)
        container.master.config(menu=menu)

        file = tk.Menu(menu)
        file.add_command(label="Home", command = lambda: self.show_frame(StartPage))
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

        button1 = ttk.Button(self, text="Send", command=lambda: email_training())
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

app = YDLapp()
app.geometry("400x300")
app.mainloop()