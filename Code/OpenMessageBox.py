from tkinter import Label, Button, PhotoImage, Tk, Frame
import webbrowser


def openlink():
    webbrowser.open('https://myaccount.google.com/lesssecureapps')


def MessageBox():
    window = Tk()
    window_app_width = 410
    window_app_height = 127
    window_screen_width = window.winfo_screenwidth()
    window_screen_height = window.winfo_screenheight()
    x = int((window_screen_width / 2) - (window_app_width / 2))
    y = int((window_screen_height / 2) - (window_app_height / 2))
    window.geometry(f'{window_app_width}x{window_app_height}+{x}+{y}')
    window.title("Information")
    window.wm_iconbitmap(r'E:\MMS\Images\one.ico')
    window.resizable(0, 0)
    info_frame = Frame(window, width=409, height=85, bg='white')
    info_frame.place(x=0, y=0)
    info_photo = PhotoImage(file=r'E:\MMS\Images\info_icon.png')
    info_photo_place = Label(info_frame, image=info_photo, bg='white')
    info_photo_place.place(x=10, y=10)
    info_text_place1 = Label(info_frame, text="Before using application please make sure that you have", font=("Product Sans", 9), bg='white')
    info_text_place1.place(x=65, y=8)
    info_text_place2 = Label(info_frame, text="turn on the less secure app setting's. Click on", font=("Product Sans", 9), bg='white')
    info_text_place2.place(x=65, y=25)
    info_text_place3 = Label(info_frame, text="\"Setting\"", font=("Product Sans", 9, "bold"), bg="white")
    info_text_place3.place(x=319, y=25)
    info_text_place4 = Label(info_frame, text="to on the less seurece setting or click", font=("Product Sans", 9), bg="white")
    info_text_place4.place(x=65, y=42)
    info_text_place5 = Label(info_frame, text="\"Yes\"", font=("Product Sans", 9, "bold"), bg="white")
    info_text_place5.place(x=272, y=42)
    info_text_place6 = Label(info_frame, text="to continue.", font=("Product Sans", 9), bg="white")
    info_text_place6.place(x=311, y=42)
    common_img = PhotoImage(width=1, height=1)
    pg2_button = Button(window, text="Setting", image=common_img, compound='c', width=65, height=14, font=("Product Sans", 9), bg='#d6d6d6', fg='black', cursor="hand2", command=openlink)
    pg2_button.place(x=236, y=95)
    yes_button = Button(window, text="Yes", image=common_img, compound='c', width=65, height=14, bg='#d6d6d6', fg='black', font=("Product Sans", 9), cursor="hand2", command=window.destroy)
    yes_button.place(x=320, y=95)
    window.mainloop()
