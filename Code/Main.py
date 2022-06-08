from tkinter import Label, Text, Button, Frame, Toplevel, NORMAL, END, messagebox, filedialog, StringVar, PhotoImage, Tk, Entry, Radiobutton
import speech_recognition
from pygame import mixer
from pandas import read_excel, isnull
from docx import Document
from PyPDF4 import PdfFileReader
import pytesseract
from PIL import Image
import imghdr
from email.message import EmailMessage
import smtplib
import os
from OpenMessageBox import MessageBox

MessageBox()

convert_type = ''
check = False
attachments = []


def secondwindow():
    window = Toplevel()
    window.rowconfigure(0, weight=1)
    window.columnconfigure(0, weight=1)
    window_app_width = 600
    window_app_height = 250
    window_screen_width = root.winfo_screenwidth()
    window_screen_height = root.winfo_screenheight()
    x = int((window_screen_width / 2) - (window_app_width / 2))
    y = int((window_screen_height / 2) - (window_app_height / 2))
    window.geometry(f'{window_app_width}x{window_app_height}+{x}+{y}')
    window.title("Mass Mail Sender")
    window.wm_iconbitmap(r'E:\MMS\Images\one.ico')
    window.resizable(0, 0)

    page1 = Frame(window)
    page2 = Frame(window)
    page3 = Frame(window)
    page4 = Frame(window)
    page5 = Frame(window)
    page6 = Frame(window)
    page7 = Frame(window)
    page8 = Frame(window)
    page9 = Frame(window)
    page10 = Frame(window)
    page11 = Frame(window)

    for frame in (page1, page2, page3, page4, page5, page6,
                  page7, page8, page9, page10, page11):
        frame.grid(row=0, column=0, sticky='nsew')

    def show_frame(frame):
        frame.tkraise()

    def check_single_or_bulk():
        if var_choice.get() == "single":
            pg2_button_browse.place_forget()
            pag2_label_bulk.place_forget()
            label_single.place(x=100, y=82)
            excel_img_label.place_forget()
            emailaddress_img_label.place(x=30, y=50)
            email_entry.config(state='normal')
            email_entry.delete(0, 'end')
        if var_choice.get() == "bulk":
            pg2_button_browse.place(x=450, y=12)
            label_single.place_forget()
            pag2_label_bulk.place(x=100, y=82)
            emailaddress_img_label.place_forget()
            excel_img_label.place(x=30, y=50)
            email_entry.delete(0, 'end')
            email_entry.config(state='readonly')

    def cancel():
        result = messagebox.askyesno('Notification', 'Do you want to exit')
        if result:
            window.destroy()
        else:
            pass

    def excel_browse():
        global emails
        excel_filename = filedialog.askopenfile(
            initialdir="/",
            title="Select A File",
            filetypes=(
                ("excel files",
                 "*.xlsx"),
                ("all files",
                 "*.*")))
        if excel_filename is not None:
            data = read_excel(excel_filename.name)
            if 'Email' in data.columns:
                emails = list(data['Email'])
                c = []
                for i in emails:
                    if isnull(i) is False:
                        c.append(i)
                emails = c
                if len(emails) > 0:
                    email_entry.config(state=NORMAL)
                    email_entry.delete(0, 'end')
                    email_entry.insert(0,
                                       str(excel_filename.name.split("/")[-1]))
                    email_entry.config(state='readonly')
                    total_label.config(text="TOTAL: " + str(len(emails)))
                    total_sent.config(text="SENT: ")
                    total_left.config(text="LEFT: ")
                    total_fail.config(text="FAILED: ")
                else:
                    messagebox.showerror(
                        "Error", "This file doesn't have any emails")
            else:
                messagebox.showerror(
                    "Error", "Please select file which have Email Column")

    def word_browse():
        global full_text, convert_type
        convert_type = "word"
        word_filename = filedialog.askopenfile(
            initialdir="/",
            title="Select A File",
            filetypes=(
                ("word files",
                 "*.docx"),
                ("all files",
                 "*.*")))
        if word_filename is not None:
            doc = Document(word_filename.name)
            full_text = []
            for para in doc.paragraphs:
                full_text.append(para.text.replace(u"\xa0", u" "))
            full_text = '\n'.join(full_text)
            if len(full_text) > 0:
                pag4_entry.config(state=NORMAL)
                pag4_entry.delete(0, END)
                pag4_entry.insert(0, str(word_filename.name.split("/")[-1]))
                pag4_entry.config(state='readonly')
            else:
                messagebox.showerror("Error", "This file is empty")
        else:
            pass

    def pdf_browse():
        global content, convert_type
        convert_type = "pdf"
        pdf_filename = filedialog.askopenfile(
            initialdir="/",
            title="Select A File",
            filetypes=(
                ("pdf files",
                 "*.pdf"),
                ("all files",
                 "*.*")))
        if pdf_filename is not None:
            pdf = PdfFileReader(pdf_filename.name)  # Load PDF into pyPDF
            # Iterate pages
            content = ""
            for i in range(0, pdf.getNumPages()):
                # Extract text from page and add to content
                content += pdf.getPage(i).extractText() + "\n"
            # Collapse whitespace
            content = ' '.join(content.replace(u"\xa0", " ").strip().split())
            if len(content) > 0:
                pag5_entry.config(state=NORMAL)
                pag5_entry.delete(0, END)
                pag5_entry.insert(0, str(pdf_filename.name.split("/")[-1]))
                pag5_entry.config(state='readonly')
            else:
                messagebox.showerror(
                    "Error", "This file is empty", parent=page5)
        else:
            pass

    def photo_browse():
        global img_text, convert_type
        convert_type = "img"
        photo_filename = filedialog.askopenfile(
            initialdir="/",
            title="Select A File",
            filetypes=(
                ("image files",
                 "*.png"),
                ("all files",
                 "*.*")))
        if photo_filename is not None:
            # Defining paths to tesseract.exe
            # and the image we would be using
            path_to_tesseract = 'E:\\MMS\\Code\\Tesseract-OCR\\tesseract.exe'
            # Opening the image & storing it in an image object
            img = Image.open(photo_filename.name)
            # Providing the tesseract executable
            # location to pytesseract library
            pytesseract.tesseract_cmd = path_to_tesseract
            # Passing the image object to image_to_string() function and This
            # function will extract the text from the image
            img_text = pytesseract.image_to_string(img)
            if len(img_text) > 0:
                pag6_entry.config(state=NORMAL)
                pag6_entry.delete(0, END)
                pag6_entry.insert(0, str(photo_filename.name.split("/")[-1]))
                pag6_entry.config(state='readonly')
            else:
                messagebox.showerror(
                    "Error", "This file is empty", parent=page6)

    def excel_clear():
        email_entry.config(state=NORMAL)
        email_entry.delete(0, END)
        email_entry.config(state="readonly")
        total_label.config(text="")
        total_sent.config(text="")
        total_left.config(text="")
        total_fail.config(text="")

    def word_clear():
        pag4_entry.config(state=NORMAL)
        pag4_entry.delete(0, 'end')
        pag4_entry.config(state='readonly')

    def pdf_clear():
        pag5_entry.config(state=NORMAL)
        pag5_entry.delete(0, 'end')
        pag5_entry.config(state='readonly')

    def photo_clear():
        pag6_entry.config(state=NORMAL)
        pag6_entry.delete(0, 'end')
        pag6_entry.config(state='readonly')

    def speak():
        mixer.init()
        mixer.music.load(r'E:\MMS\Music\music1.mp3')
        mixer.music.play()
        sr = speech_recognition.Recognizer()
        with speech_recognition.Microphone() as m:
            try:
                sr.adjust_for_ambient_noise(m, duration=0.2)
                audio = sr.listen(m)
                text = sr.recognize_google(audio)
                pag8_speechbox.insert(END, text + '.')
            except Exception:
                pass

    def attachment():
        global filename, filetype, filepath, check, attachments
        check = True
        filepath = filedialog.askopenfilename(
            initialdir='/', title='Select File')
        attachments.append(filepath)
        filename = os.path.basename(filepath)
        pag10_entry.insert(END, f'{filename}\t')

    def check_file_type():
        if check:
            for filepath in attachments:
                filetype = filepath.split('.')
                filetype = filetype[1]
                if filetype == 'png' or filetype == 'PNG' or filetype == 'jpg' or filetype == 'jpeg':
                    with open(filepath, 'rb') as f:
                        message.add_attachment(
                            f.read(),
                            maintype='image',
                            subtype=imghdr.what(filepath),
                            filename=os.path.basename(filepath))
                else:
                    with open(filepath, 'rb') as f:
                        message.add_attachment(
                            f.read(),
                            maintype='application',
                            subtype='octet-stream',
                            filename=os.path.basename(filepath))

    def send_email():
        global s_count, f_count
        if var_choice.get() == 'single':
            status = SendingEmail(email_entry.get(), pag9_entry.get())
            if status == 's':
                messagebox.showinfo('Success', 'Email is sent successfuly')
            if status == 'f':
                messagebox.showerror('Error', 'Email is not sent')
        if var_choice.get() == 'bulk':
            s_count = 0
            f_count = 0
            for x in emails:
                status = SendingEmail(x, pag9_entry.get())
                if status == 's':
                    s_count += 1
                if status == 'f':
                    f_count += 1
                total_label.config(text='')
                total_sent.config(text='SENT: ' + str(s_count))
                total_left.config(
                    text='LEFT: ' + str(len(emails) - (s_count + f_count)))
                total_fail.config(text='FAILED: ' + str(f_count))
                total_label.update()
                total_sent.update()
                total_left.update()
                total_fail.update()
            messagebox.showinfo('Success', 'Emails are sent successfully')

    def SendingEmail(toAddress, subject):
        global message
        message = EmailMessage()
        message['subject'] = subject
        message['to'] = toAddress
        message['from'] = email
        total_body = ""
        if convert_type == 'word':
            total_body = full_text
        if convert_type == 'pdf':
            total_body = content
        if convert_type == 'img':
            total_body = img_text
        if len(pag7_textbox.get('1.0', 'end-1c')) > 0:
            total_body = pag7_textbox.get('1.0', 'end-1c')
        if len(pag8_speechbox.get('1.0', 'end-1c')) > 0:
            total_body = pag8_speechbox.get('1.0', 'end-1c')
        message.set_content(total_body)
        check_file_type()
        server.send_message(message)
        x = server.ehlo()
        if x[0] == 250:
            return 's'
        else:
            return 'f'
        server.quit()

    show_frame(page1)

    # ==================================== Page 1 ============================
    page1.config(background='dodger blue')
    page1_img = PhotoImage(file=r'E:\MMS\Images\EmailLogo.png')
    img_label = Label(page1, image=page1_img, bg='dodger blue').pack()
    var_choice = StringVar()
    single = Radiobutton(
        page1,
        text="Single",
        value="single",
        variable=var_choice,
        font=(
            "times new roman",
            30,
            "bold"),
        bg="dodger blue",
        activebackground="dodger blue",
        command=check_single_or_bulk)
    single.place(x=100, y=100)
    bulk = Radiobutton(
        page1,
        text="Bulk",
        value="bulk",
        variable=var_choice,
        font=(
            "times new roman",
            30,
            "bold"),
        bg="dodger blue",
        activebackground="dodger blue",
        command=check_single_or_bulk)
    bulk.place(x=350, y=100)
    var_choice.set('single')
    button_frame = Frame(
        page1,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=50,
        bg='dodger blue')
    button_frame.place(x=15, y=190)
    pg1_button = Button(
        button_frame,
        text="Next >",
        command=lambda: show_frame(page2),
        width=10,
        bg='brown',
        fg='white')
    pg1_button.place(x=285, y=10)
    pg1_button = Button(
        button_frame,
        text="Cancel",
        command=cancel,
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg1_button.place(x=405, y=10)

    # ==================================== Page 2 ============================
    page2.config(background='dodger blue')
    page1_img = PhotoImage(file=r'E:\MMS\Images\EmailLogo.png')
    img_label = Label(page2, image=page1_img, bg='dodger blue').pack()
    pg2_frame = Frame(
        page2,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=70,
        bg='dodger blue')
    pg2_frame.place(x=15, y=97)
    emailaddress_img = PhotoImage(
        file=r'E:\MMS\Images\emailaddress.png')
    emailaddress_img_label = Label(
        page2,
        image=emailaddress_img,
        bg='dodger blue',
        width=55,
        height=55)
    excel_img = PhotoImage(file=r'E:\MMS\Images\Excelimage.png')
    excel_img_label = Label(
        page2,
        image=excel_img,
        bg='dodger blue',
        width=55,
        height=55)
    pag2_label_bulk = Label(
        page2,
        text="Select Excel file for Email address:",
        width=24,
        fg='black',
        bg='dodger blue',
        font=(
            "Goudy Old Style",
            13,
            "bold"))
    label_single = Label(
        page2,
        text="Enter Email Address:",
        width=15,
        fg='black',
        bg='dodger blue',
        font=(
            "Goudy Old Style",
            13,
            "bold"))
    pg2_button_browse = Button(
        pg2_frame,
        text="Browse...",
        command=excel_browse,
        width=10,
        bg='white',
        fg='black',
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg2_button_clear = Button(
        pg2_frame,
        text="Clear",
        command=excel_clear,
        width=10,
        bg='white',
        fg='black',
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg2_button_clear.place(x=48, y=40)
    # Entries
    email_entry = Entry(pg2_frame, width=60, bg='light goldenrod')
    email_entry.place(x=48, y=15)
    # Buttons
    button_frame = Frame(
        page2,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=50,
        bg='dodger blue')
    button_frame.place(x=15, y=190)
    pg2_button_next = Button(
        button_frame,
        text="Next >",
        command=lambda: show_frame(page3),
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg2_button_next.place(x=285, y=10)
    pg2_button = Button(
        button_frame,
        text="< Back",
        width=10,
        bg='brown',
        fg='white',
        command=lambda: show_frame(page1),
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg2_button.place(x=50, y=10)
    pg2_button_cancel = Button(
        button_frame,
        text="Cancel",
        command=cancel,
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg2_button_cancel.place(x=405, y=10)

    # ==================================== Page 3 ============================
    page3.config(background='dodger blue')
    img_label = Label(page3, image=page1_img, bg='dodger blue').pack()
    box_frame = Frame(
        page3,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=90,
        bg='dodger blue')
    box_frame.place(x=15, y=78)
    pag3_label = Label(
        page3,
        text="Select body format:",
        width=15,
        fg='black',
        bg='dodger blue',
        font=(
            "Goudy Old Style",
            13,
            "bold"))
    pag3_label.place(x=225, y=65)
    word_img = PhotoImage(file=r'E:\MMS\Images\Wordimage.png')
    pdf_img = PhotoImage(file=r'E:\MMS\Images\pdfimage.png')
    speech_img = PhotoImage(file=r'E:\MMS\Images\speechimage.png')
    text_img = PhotoImage(file=r'E:\MMS\Images\textimage.png')
    photo_img = PhotoImage(file=r'E:\MMS\Images\photoimage.png')
    word_button = Button(
        box_frame,
        image=word_img,
        bd=0,
        bg='dodger blue',
        activebackground='dodger blue',
        cursor="hand2",
        command=lambda: show_frame(page4))
    word_button.place(x=250, y=10)
    pdf_button = Button(
        box_frame,
        image=pdf_img,
        bd=0,
        bg='dodger blue',
        activebackground='dodger blue',
        cursor="hand2",
        command=lambda: show_frame(page5))
    pdf_button.place(x=150, y=10)
    text_button = Button(
        box_frame,
        image=text_img,
        bd=0,
        bg='dodger blue',
        activebackground='dodger blue',
        cursor="hand2",
        command=lambda: show_frame(page7))
    text_button.place(x=50, y=10)
    speech_button = Button(
        box_frame,
        image=speech_img,
        bd=0,
        bg='dodger blue',
        activebackground='dodger blue',
        cursor="hand2",
        command=lambda: show_frame(page8))
    speech_button.place(x=350, y=10)
    photo_button = Button(
        box_frame,
        image=photo_img,
        bd=0,
        bg='dodger blue',
        activebackground='dodger blue',
        cursor="hand2",
        command=lambda: show_frame(page6))
    photo_button.place(x=450, y=10)
    button_frame = Frame(
        page3,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=50,
        bg='dodger blue')
    button_frame.place(x=15, y=190)
    pg3_button = Button(
        button_frame,
        text="< Back",
        width=10,
        bg='brown',
        fg='white',
        command=lambda: show_frame(page2),
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg3_button.place(x=50, y=10)
    pg3_button = Button(
        button_frame,
        text="Cancel",
        command=cancel,
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg3_button.place(x=405, y=10)

    # ==================================== word Page  ========================
    page4.config(background='dodger blue')
    img_label = Label(page4, image=page1_img, bg='dodger blue').pack()
    text_frame = Frame(
        page4,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=70,
        bg='dodger blue')
    text_frame.place(x=15, y=97)
    word_img2 = PhotoImage(file=r'E:\MMS\Images\Wordimage.png')
    word_img_label2 = Label(
        page4,
        image=word_img2,
        bg='dodger blue',
        width=55,
        height=55)
    word_img_label2.place(x=30, y=50)
    word_label = Label(
        page4,
        text="Select Word File:",
        width=13,
        fg='black',
        bg='dodger blue',
        font=(
            "Goudy Old Style",
            13,
            "bold"))
    word_label.place(x=100, y=82)
    # Entries
    pag4_entry = Entry(
        text_frame,
        width=60,
        bg='light goldenrod',
        state='readonly')
    pag4_entry.place(x=48, y=15)
    pg4_button_browse = Button(
        text_frame,
        text="Browse...",
        command=word_browse,
        width=10,
        bg='white',
        fg='black',
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg4_button_browse.place(x=450, y=12)
    pg4_button_clear = Button(
        text_frame,
        text="Clear",
        command=word_clear,
        width=10,
        bg='white',
        fg='black',
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg4_button_clear.place(x=48, y=40)
    button_frame = Frame(
        page4,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=50,
        bg='dodger blue')
    button_frame.place(x=15, y=190)
    pg4_button = Button(
        button_frame,
        text="Next >",
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove',
        command=lambda: show_frame(page9))
    pg4_button.place(x=285, y=10)
    pg4_button = Button(
        button_frame,
        text="< Back",
        width=10,
        bg='brown',
        fg='white',
        command=lambda: show_frame(page3),
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg4_button.place(x=50, y=10)
    pg4_button = Button(
        button_frame,
        text="Cancel",
        command=cancel,
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg4_button.place(x=405, y=10)

    # ==================================== pdf Page ==========================
    page5.config(background='dodger blue')
    img_label = Label(page5, image=page1_img, bg='dodger blue').pack()
    text_frame = Frame(
        page5,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=70,
        bg='dodger blue')
    text_frame.place(x=15, y=97)
    pdf_img2 = PhotoImage(file=r'E:\MMS\Images\pdfimage.png')
    pdf_img_label2 = Label(
        page5,
        image=pdf_img2,
        bg='dodger blue',
        width=55,
        height=55)
    pdf_img_label2.place(x=30, y=50)
    pdf_label = Label(
        page5,
        text="Select Pdf File:",
        width=13,
        fg='black',
        bg='dodger blue',
        font=(
            "Goudy Old Style",
            13,
            "bold"))
    pdf_label.place(x=100, y=82)
    # Entries
    pag5_entry = Entry(
        text_frame,
        width=60,
        state='readonly',
        bg='light goldenrod')
    pag5_entry.place(x=48, y=15)
    pg5_button_clear = Button(
        text_frame,
        text="Clear",
        command=photo_clear,
        width=10,
        bg='white',
        fg='black',
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg5_button_clear.place(x=48, y=40)
    pg5_button_browse = Button(
        text_frame,
        text="Browse...",
        command=pdf_browse,
        width=10,
        bg='white',
        fg='black',
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg5_button_browse.place(x=450, y=12)
    button_frame = Frame(
        page5,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=50,
        bg='dodger blue')
    button_frame.place(x=15, y=190)
    pg5_button = Button(
        button_frame,
        text="Next >",
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove',
        command=lambda: show_frame(page9))
    pg5_button.place(x=285, y=10)
    pg5_button = Button(
        button_frame,
        text="< Back",
        width=10,
        bg='brown',
        fg='white',
        command=lambda: show_frame(page3),
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg5_button.place(x=50, y=10)
    pg5_button = Button(
        button_frame,
        text="Cancel",
        command=cancel,
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg5_button.place(x=405, y=10)

    # ==================================== photo Page ========================
    page6.config(background='dodger blue')
    img_label = Label(page6, image=page1_img, bg='dodger blue').pack()
    text_frame = Frame(
        page6,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=70,
        bg='dodger blue')
    text_frame.place(x=15, y=97)
    photo_img2 = PhotoImage(file=r'E:\MMS\Images\photoimage.png')
    photo_img_label2 = Label(
        page6,
        image=photo_img2,
        bg='dodger blue',
        width=55,
        height=55)
    photo_img_label2.place(x=30, y=50)
    photo_label = Label(
        page6,
        text="Select Image File:",
        width=13,
        fg='black',
        bg='dodger blue',
        font=(
            "Goudy Old Style",
            13,
            "bold"))
    photo_label.place(x=100, y=82)
    # Entries
    pag6_entry = Entry(
        text_frame,
        width=60,
        state='readonly',
        bg='light goldenrod')
    pag6_entry.place(x=48, y=12)
    pg6_button_clear = Button(
        text_frame,
        text="Clear",
        command=photo_clear,
        width=10,
        bg='white',
        fg='black',
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg6_button_clear.place(x=48, y=40)
    pg6_button_browse = Button(
        text_frame,
        text="Browse...",
        command=photo_browse,
        width=10,
        bg='white',
        fg='black',
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg6_button_browse.place(x=450, y=12)
    button_frame = Frame(
        page6,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=50,
        bg='dodger blue')
    button_frame.place(x=15, y=190)
    pg6_button = Button(
        button_frame,
        text="Next >",
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove',
        command=lambda: show_frame(page9))
    pg6_button.place(x=285, y=10)
    pg6_button = Button(
        button_frame,
        text="< Back",
        width=10,
        bg='brown',
        fg='white',
        command=lambda: show_frame(page3),
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg6_button.place(x=50, y=10)
    pg6_button = Button(
        button_frame,
        text="Cancel",
        command=cancel,
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg6_button.place(x=405, y=10)

    # ==================================== text Page =========================
    page7.config(background='dodger blue')
    img_label = Label(page7, image=page1_img, bg='dodger blue').pack()
    text_box_frame = Frame(
        page7,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=100,
        bg='dodger blue')
    text_box_frame.place(x=15, y=67)
    # Storage
    pag7_textbox = Text(
        text_box_frame,
        width=60,
        height=4,
        bg='light goldenrod')
    pag7_textbox.place(x=20, y=20)
    pg7_button_clear = Button(
        text_box_frame,
        text="Clear",
        width=5,
        bg='white',
        fg='black',
        cursor="hand2",
        relief='raised',
        overrelief='groove',
        command=pag7_textbox.delete(1.0, END))
    pg7_button_clear.place(x=510, y=43)
    pag7_label = Label(
        page7,
        text="Enter Message:",
        width=12,
        fg='black',
        bg='dodger blue',
        font=(
            "Goudy Old Style",
            13,
            "bold"))
    pag7_label.place(x=100, y=53)
    text_img2 = PhotoImage(file=r'E:\MMS\Images\textimage.png')
    text_img_label2 = Label(
        page7,
        image=text_img2,
        bg='dodger blue',
        width=55,
        height=55)
    text_img_label2.place(x=30, y=25)
    button_frame = Frame(
        page7,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=50,
        bg='dodger blue')
    button_frame.place(x=15, y=190)
    pg7_button = Button(
        button_frame,
        text="Next >",
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove',
        command=lambda: show_frame(page9))
    pg7_button.place(x=285, y=10)
    pg7_button = Button(
        button_frame,
        text="< Back",
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove',
        command=lambda: show_frame(page3))
    pg7_button.place(x=50, y=10)
    pg7_button = Button(
        button_frame,
        text="Cancel",
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove',
        command=cancel)
    pg7_button.place(x=405, y=10)

    # ==================================== speech Page =======================
    page8.config(background='dodger blue')
    speech_label = Label(page8, image=page1_img, bg='dodger blue')
    speech_label.pack()
    speech_box_frame = Frame(
        page8,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=100,
        bg='dodger blue')
    speech_box_frame.place(x=15, y=67)
    # Storage
    pag8_speechbox = Text(
        speech_box_frame,
        width=60,
        height=4,
        bg='light goldenrod')
    pag8_speechbox.place(x=20, y=20)
    pg8_button_clear = Button(
        speech_box_frame,
        text="Clear",
        width=5,
        bg='white',
        fg='black',
        cursor="hand2",
        relief='raised',
        overrelief='groove',
        command=pag8_speechbox.delete(1.0, END))
    pg8_button_clear.place(x=510, y=43)
    pag8_label = Label(
        page8,
        text="Click Microphone and speak the Message:",
        width=30,
        fg='black',
        bg='dodger blue',
        font=(
            "Goudy Old Style",
            13,
            "bold"))
    pag8_label.place(x=100, y=53)
    speech_img2 = PhotoImage(file=r'E:\MMS\Images\speechimage2.png')
    speech_button2 = Button(
        page8,
        image=speech_img2,
        bd=0,
        width=55,
        height=58,
        bg='dodger blue',
        activebackground='dodger blue',
        cursor="hand2",
        command=speak)
    speech_button2.place(x=30, y=25)
    button_frame = Frame(
        page8,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=50,
        bg='dodger blue')
    button_frame.place(x=15, y=190)
    pg8_button = Button(
        button_frame,
        text="Next >",
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove',
        command=lambda: show_frame(page9))
    pg8_button.place(x=285, y=10)
    pg8_button = Button(
        button_frame,
        text="< Back",
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove',
        command=lambda: show_frame(page3))
    pg8_button.place(x=50, y=10)
    pg8_button = Button(
        button_frame,
        text="Cancel",
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove',
        command=cancel)
    pg8_button.place(x=405, y=10)

    # ==================================== subject Page ======================
    page9.config(background='dodger blue')
    subject_label = Label(page9, image=page1_img, bg='dodger blue').pack()
    subject_frame = Frame(
        page9,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=70,
        bg='dodger blue')
    subject_frame.place(x=15, y=97)
    subject_img = PhotoImage(file=r'E:\MMS\Images\subjectimage.png')
    subject_img_label = Label(
        page9,
        image=subject_img,
        bg='dodger blue',
        width=55,
        height=55)
    subject_img_label.place(x=30, y=50)
    subject_label = Label(
        page9,
        text="Subject:",
        width=8,
        fg='black',
        bg='dodger blue',
        font=(
            "Goudy Old Style",
            13,
            "bold"))
    subject_label.place(x=100, y=82)
    # Entries
    pag9_entry = Entry(subject_frame, width=60, bg='light goldenrod')
    pag9_entry.place(x=48, y=25)
    pg9_button_clear = Button(
        subject_frame,
        text="Clear",
        width=10,
        bg='white',
        fg='black',
        cursor="hand2",
        relief='raised',
        overrelief='groove',
        command=pag9_entry.delete(0, 'end'))
    pg9_button_clear.place(x=450, y=23)
    button_frame = Frame(
        page9,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=50,
        bg='dodger blue')
    button_frame.place(x=15, y=190)
    pg9_button = Button(
        button_frame,
        text="Next >",
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove',
        command=lambda: show_frame(page10))
    pg9_button.place(x=285, y=10)
    pg9_button = Button(
        button_frame,
        text="< Back",
        width=10,
        bg='brown',
        fg='white',
        command=lambda: show_frame(page3),
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg9_button.place(x=50, y=10)
    pg9_button = Button(
        button_frame,
        text="Cancel",
        command=cancel,
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg9_button.place(x=405, y=10)

    # ==================================== attachment Page ===================
    page10.config(background='dodger blue')
    subject_label = Label(page10, image=page1_img, bg='dodger blue').pack()
    Frame10 = Frame(
        page10,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=70,
        bg='dodger blue')
    Frame10.place(x=15, y=97)
    attachment_img = PhotoImage(
        file=r'E:\MMS\Images\attachmentimage.png')
    attachment_button = Button(
        Frame10,
        image=attachment_img,
        bd=0,
        width=55,
        height=55,
        bg='dodger blue',
        activebackground='dodger blue',
        cursor="hand2",
        command=attachment)
    attachment_button.place(x=450, y=2)
    pag10_entry = Text(Frame10, width=45, height=0, bg='light goldenrod')
    pag10_entry.place(x=48, y=12)
    pg10_button_clear = Button(
        Frame10,
        text="Clear",
        width=10,
        bg='white',
        fg='black',
        cursor="hand2",
        relief='raised',
        overrelief='groove',
        command=pag10_entry.delete(1.0, END))
    pg10_button_clear.place(x=48, y=40)
    # Entries
    subject_label = Label(
        page10,
        text="Attachment:",
        width=10,
        fg='black',
        bg='dodger blue',
        font=(
            "Goudy Old Style",
            13,
            "bold"))
    subject_label.place(x=100, y=82)
    button_frame = Frame(
        page10,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=50,
        bg='dodger blue')
    button_frame.place(x=15, y=190)
    pg10_button = Button(
        button_frame,
        text="Next >",
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove',
        command=lambda: show_frame(page11))
    pg10_button.place(x=285, y=10)
    pg10_button = Button(
        button_frame,
        text="< Back",
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove',
        command=lambda: show_frame(page9))
    pg10_button.place(x=50, y=10)
    pg10_button = Button(
        button_frame,
        text="Cancel",
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove',
        command=cancel)
    pg10_button.place(x=405, y=10)

    # ==================================== status page =======================
    page11.config(background='dodger blue')
    img_label = Label(page11, image=page1_img, bg='dodger blue')
    img_label.pack()
    box_frame = Frame(
        page11,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=90,
        bg='dodger blue')
    box_frame.place(x=15, y=78)
    total_label = Label(
        box_frame,
        text='TOTAL:',
        font=(
            "times new roman",
            20,
            'bold'),
        bg='dodger blue')
    total_label.place(x=20, y=25)
    total_sent = Label(
        box_frame,
        text='SENT:',
        font=(
            "times new roman",
            20,
            'bold'),
        bg='dodger blue')
    total_sent.place(x=170, y=25)
    total_left = Label(
        box_frame,
        text='LEFT:',
        font=(
            "times new roman",
            20,
            'bold'),
        bg='dodger blue')
    total_left.place(x=300, y=25)
    total_fail = Label(
        box_frame,
        text='FAILED:',
        font=(
            "times new roman",
            20,
            'bold'),
        bg='dodger blue')
    total_fail.place(x=420, y=25)
    pag10_label = Label(
        page11,
        text="Status:",
        width=6,
        fg='black',
        bg='dodger blue',
        font=(
            "Goudy Old Style",
            13,
            "bold"))
    pag10_label.place(x=266, y=65)
    button_frame = Frame(
        page11,
        highlightbackground="black",
        highlightthickness=1,
        width=570,
        height=50,
        bg='dodger blue')
    button_frame.place(x=15, y=190)
    pg11_button = Button(
        button_frame,
        text="Send",
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove',
        command=send_email)
    pg11_button.place(x=285, y=10)
    pg11_button = Button(
        button_frame,
        text="< Back",
        width=10,
        bg='brown',
        fg='white',
        command=lambda: show_frame(page10),
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg11_button.place(x=50, y=10)
    pg11_button = Button(
        button_frame,
        text="Cancel",
        command=cancel,
        width=10,
        bg='brown',
        fg='white',
        cursor="hand2",
        relief='raised',
        overrelief='groove')
    pg11_button.place(x=405, y=10)

    window.mainloop()


def Login():
    global email, password, server
    if len(Rmail.get()) == 0 and len(Rpassword.get()) == 0:
        messagebox.showwarning("Error", "fill fields")
    else:
        email = str(Rmail.get())
        password = str(Rpassword.get())
        try:
            server = smtplib.SMTP("smtp.gmail.com", 587)
        except Exception as err:
            messagebox.showerror("Error", err)
            server = None
        if server is not None:
            # sending smtp "hello" message
            server.starttls()
            try:
                # login in to server
                server.login(email, password)
                secondwindow()
                # checking the server
            except smtplib.SMTPAuthenticationError:
                messagebox.showerror(
                    "Error",
                    "->Username and Password Not Validate Credentials\n->check you less secure app setting\n->Email not sent!")
                # sending mail
                server.quit()


def Reset():
    emailE.delete(0, 'end')
    passwordE.delete(0, 'end')


def show_hide_password():
    if passwordE['show'] == '*':
        passwordE.configure(show='')
        show_hide_btn.configure(image=show_face)
    else:
        passwordE.configure(show='*')
        show_hide_btn.configure(image=hide_face)

# main screen


root = Tk()
root_app_width = 600
root_app_height = 200
root_screen_width = root.winfo_screenwidth()
root_screen_height = root.winfo_screenheight()
x = int((root_screen_width / 2) - (root_app_width / 2))
y = int((root_screen_height / 2) - (root_app_height / 2))
root.geometry(f'{root_app_width}x{root_app_height}+{x}+{y}')
root.title("Mass Mail Sender")
root.wm_iconbitmap(r'E:\MMS\Images\one.ico')
root.resizable(0, 0)
root.config(bg='dodger blue')
# Graphics
logoImage = PhotoImage(file=r'E:\MMS\Images\EmailLogo.png')
show_face = PhotoImage(file=r'E:\MMS\Images\view_show.png')
hide_face = PhotoImage(file=r'E:\MMS\Images\view_hide.png')
titleLabel = Label(
    root,
    text="LOGIN",
    image=logoImage,
    bg='dodger blue').pack()
Label_1 = Label(
    root,
    text="Email:",
    width=20,
    fg='black',
    bg='dodger blue',
    font=(
        "Goudy Old Style",
        13,
         "bold"))
Label_1.place(x=5, y=70)
Label_2 = Label(
    root,
    text="Password:",
    width=20,
    fg='black',
    bg='dodger blue',
    font=(
        "Goudy Old Style",
        13,
         "bold"))
Label_2.place(x=16, y=90)

# Storage
Rmail = StringVar()
Rpassword = StringVar()

# Entries
emailE = Entry(root, width=40, textvariable=Rmail)
emailE.place(x=180, y=70)
passwordE = Entry(root, width=40, show='*', textvariable=Rpassword)
passwordE.place(x=180, y=90)

# Buttons
widget = Button(
    root,
    text="Login",
    command=Login,
    width=10,
    bg='brown',
    fg='white')
widget.place(x=190, y=135)
widget1 = Button(
    root,
    text="Reset",
    command=Reset,
    width=10,
    bg='brown',
    fg='white')
widget1.place(x=325, y=135)
show_hide_btn = Button(
    root,
    image=hide_face,
    bd=0,
    bg='dodger blue',
    activebackground='dodger blue',
    cursor="hand2",
    width=32,
    height=32,
    command=show_hide_password)
show_hide_btn.place(x=440, y=82)
root.mainloop()
