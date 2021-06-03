import tkinter as tk
from tkinter import *
from PIL import ImageTk, Image

# for facedetection
import face_recognition
import sys
import cv2

# ......for excel
import pandas as pd
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np

root = Tk()
root.title("jimmy")
root.geometry("1366x768")
root.configure(backg="black")

# root.configure(backg="pink")

import datetime
import os
from openpyxl import Workbook
from tkinter import *
from tkinter import messagebox
import random
import time
from typing import Counter
import pyttsx3
import speech_recognition as sr
import wikipedia  
import datetime
import webbrowser
import pipwin
import xlrd
import win32api
import pywhatkit
import playsound
import smtplib
#......for excel
import pandas as pd
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np

from tkinter import*

global approvement   #ture only when the face will get proper approvement

#for facedetection
import face_recognition
import sys
import cv2
global files

#........../\...........
#catch picture dynamicly
def capture_image():
    videoCaptureObject = cv2.VideoCapture(0)
    result = True
    while result:
        ret, frame = videoCaptureObject.read()
        cv2.imwrite("images/NewPicture.jpg", frame)
        result = False
    videoCaptureObject.release()
    cv2.destroyAllWindows()

#to analyse the face match with tarining data
def analyze_user():
    print("analyzing face, please wait.....")
    baseimg = face_recognition.load_image_file('images/sohan_pic1.jpg')
    baseimg = cv2.cvtColor(baseimg, cv2.COLOR_BGR2RGB)
    #print(len(baseimg))
    #cv2.imshow("test", baseimg)
    #cv2.waitKey(0)
    encodeface = face_recognition.face_encodings(baseimg)[0]
    #cv2.rectangle(baseimg, (myface[3], myface[0]), (myface[1], myface[2]), (255, 0, 255), 2)

    sampleimg = face_recognition.load_image_file('images/NewPicture.jpg')
    sampleimg = cv2.cvtColor(sampleimg, cv2.COLOR_BGR2RGB)
    #print(len(sampleimg))

    try:
        #cv2.imshow("test", sampleimg)
        #cv2.waitKey(0)
        encodesamplefacetest = face_recognition.face_encodings(sampleimg)[0]
        result = face_recognition.compare_faces([encodeface], encodesamplefacetest)
        resultstring = str(result)

        if resultstring == "[True]":
            approvement = True
            return approvement
        else:
            approvement = False
            return approvement
        
    except IndexError as e:
        speak("sorry, there is some issue in face recognition engine")
        sys.exit()


mail_information = {}
mail_information = {'rooman':'ummerooman54@gmail.com','sohan':'sohannaik27@gmail.com'}

global memo   #web database
memo = {'open youtube':'www.youtube.com','open google':'www.google.com','open facebook':'www.facebook.com',
'open web':'www.google.com','open gmail':'www.gmail.com','open amazon':'www.amazon.com','open Amazom Prime':'www.primevideo.com',
'open prime':'www.primevideo.com','open whatsapp':'web.whatsapp.com','open flipkart':'www.flipkart.com'}
whatsappp = {'sohan':'9606272347'}#contact database


# for voice...........................................
engine = pyttsx3.init('sapi5')  # fr voices
voices = engine.getProperty('voices')
# print (voices[1].id)pre
engine.setProperty('voice', voices[1].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def takeCommand():
    # it takes the micro phone input and returns the string as outout
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening....\n")
        r.pause_threshold = 1
        print("talk")
        audio = r.listen(source)

    try:
        print("Recognizing sir...\n")
        query = r.recognize_google(audio, language='en-in')
        result = "you said: {}\n".format(query)
        print(result)
    except sr.UnknownValueError:
        # print(e)
        print("sohan, i didn't get you, say that again plz\n")
        return "None"
    except sr.RequestError:
        print("sohan, my speach service is down\n")
        return "None"
    return query


def searchFor(ask=False):
    with sr.Microphone() as source:
        if ask:
            print(ask)
            speak(ask)

        r = sr.Recognizer()
        audio = r.listen(source)
        voice_data = ''
        try:
            voice_data = r.recognize_google(audio, language='en-in')
        except sr.UnknownValueError:
            speak('sorry, i didnt get you!')

        except sr.RequestError:
            speak(user_name)
            speak('sorry, looks like your internet us down!!')
        return voice_data

        
def wishMe():
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        speak("good morning there!")
    elif 12 <= hour < 18:
        speak("good afternoon There!")
    else:
        speak("good evening there..!")


def call_jimmy(approved):
    if approved == True:
        #playsound.playsound('JARVIS.mp3')
        assis = ['hey jimmy','jimmy','you there','you there jimmy','hello jimmy','hi jimmy']
        speak('im jimmy here, how can i help you.')
        
        query = takeCommand().lower()
        if query in memo:
            speak('okay'+user_name+'opening'+memo[query]+' for you..')
            webbrowser.open(query)
        elif 'on stock' in query or 'show stock' in query or 'stock market' in query:
            speak("let me see, what are the current stock going on, here you go")
            stockk = "https://www1.nseindia.com/live_market/dynaContent/live_watch/equities_stock_watch.htm"
            webbrowser.open(stockk)
        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(strTime)

        elif 'search about' in query:
            search = searchFor('what do you want to search sohan?')
            url = 'https://google.com/search?q=' + search
            webbrowser.get().open(url)
            speak('here is what i have found about' + search)
                            
        elif 'Tell me about' in query:
            search = searchFor('what do you want to search sohan?')
            url = 'https://google.com/search?q=' + search
            webbrowser.get().open(url)
            speak('here is what i have found about' + search)

        elif 'find the location' in query:
            location = searchFor('okay'+user_name+'what location you want')
            url = 'https://google.nl/maps/place/' + location + '/&amp;'
            webbrowser.get().open(url)
            print('here is what i have found in map' + location)

        elif 'shutdown' in query or 'offline' in query:
            speak('goodbye, '+user_name+' lets catch up later')
            exit()

        elif 'send message' in query:
            speak('whom do you want to send messgae ?')
            mess_name ='sohan'
            speak('what will be the messgae for '+mess_name)
            mess_text = searchFor("message")
            t1 = datetime.datetime.now().hour
            t2 = datetime.datetime.now().minute
            speak("okay will send message to"+mess_name+'and message is'+mess_text)
            t2 +=2
            code = "+91"
            code = code+user_phone
            pywhatkit.sendwhatmsg(code,mess_text,t1,t2)
            #query = ""

        elif 'send mail' in query or 'reply mail' in query or 'send email' in query:

            speak("whome do you want to send mail")
            person_name = takeCommand()
            person_name = person_name.lower()
            idinty_person = mail_information[person_name]
            speak("what do you want to convey")
            content = takeCommand()
            try:
                server = smtplib.SMTP('smtp.gmail.com',465)
                server.ehlo()
                server.starttls()
                server.login(user_email,user_email_password)
                server.sendmail(user_email,idinty_person,content)
                server.close()
                speak("okay,"+user_name+"mail sent!")
            except:
                speak("sorry"+user_name+"couldn't send the email")

        elif 'play music' in query or "play song" in query:
            speak("Here you go with music")
            # music_dir = "G:\\Song"
            music_dir = "C:\\Users\\naik\\Music\\songs\\New"
            songs = os.listdir(music_dir)
            r1 = random.randint(0,80)
            os.startfile(os.path.join(music_dir, songs[r1]))

        elif 'your developer' in query or "who created you" in query or "who developed you" in query:
            speak("Umme rooman, and Sohan Naik are my master and creater")
            speak("I was developed in the year june, two thousand twenty one")

        elif 'open brave' in query:
            codePath = r"C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe"
            os.startfile(codePath)

        elif 'open stackoverflow' in query or 'stack overflow' in query:
            speak("Sure will open stack overflow for you")
            webbrowser.open("stackoverflow.com")

        elif 'my name' in query:
            speak('your name is'+user_name)
            print(user_name)

        elif 'wikipedia' in query:
            speak('okay will search wikipedia..')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=1)
            # print("{}\n".format(results))
            speak(user_name)
            speak("Here you go, according to known:")
            speak(results)
        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(strTime)
        elif 'find the location' in query or 'location' in query or 'locate' in query:
            location = searchFor('what is the location, sohan')
            url = 'https://google.nl/maps/place/' + location + '/&amp;'
            webbrowser.get().open(url)
            speak('here is what i have found in map' + location)
        
        elif 'can you help' in query or 'help me' in query or 'how can you help me' in query or 'what can you do' in query or 'can you do' in query:
            speak("i can help you in doing task easy, just ask me to search wikipeida or open somthing from web")
            speak("also you can ask me to play music, i have good collections of songs")
        
        else:
            speak("I apologize "+user_name+"there is some truble in recognize your commond, can you repeat the command")
        query = ""
        call_interactive_page_to_talk(True)


def call_interactive_page_to_talk(approve):
    root3 = Tk()
    root3.geometry('1266x768')
    root3.title("Communicating")

    def call_jimmy_totake_commond():
        root3.destroy()
        call_jimmy(approve)
            
    root3.configure(backg="black")
    label_2 = Label(root3, text="Start Your Talk",width=15,font=("bold", 45),fg="yellow",bg="black")
    label_2.place(x=380,y=53)
    log=PhotoImage(file="talk.png")
    img=Button(root3,image=log,borderwidth=0,bg="black",command=call_jimmy_totake_commond)
    img.place(x=550,y=200)
    root3.mainloop()


                
#................................main jimmy runs from here...........
            
def main_run():
    df=pd.read_excel("E:\\mirror\\user_data_storage.xlsx",sheet_name="Sheet1")
    datas = df['data']
    print(datas)
    #data->medium...new 0 then 1 ...2.3.4.5
    if datas[0] == "":
        speak("looks like you have not registered as user")
    else:
        global user_name
        global user_email
        global user_phone
        global user_email_password
        global mail_information
        user_name =  str(datas[0])
        user_email = datas[1]
        user_email_password = datas[2]
        user_phone = datas[3]
        
        wishMe()  #wishes the person
        capture_image()#captues the image
        approvement = analyze_user() # recognition analysis
        if approvement == TRUE:
            call_interactive_page_to_talk(approvement)
        #call_jimmy()


def get_login():
    root.destroy()
    main_run()
    

def get_registered():
    root.destroy()
    root1 = Tk()
    root1.geometry('1366x768')
    root1.configure(backg="black")
    root1.title("Registration Form")
    #bg = PhotoImage(file=.png")
    #my_lable = Label(root1, image=bg)
    #my_lable.place(x=0, y=0)

    label_r0 = Label(root1, text="Registration form", width=20, font=("bold", 35),fg="yellow",bg="black")
    label_r0.place(x=380, y=0)
    bg= PhotoImage(file="reg.png")
    my_lable = Label(root1, image=bg,width=200,height=200,bg="black",borderwidth=0)
    my_lable.place(x=550, y=60)

    label_r1 = Label(root1, text="Name", anchor="c", width=20, font=("bold", 10), fg="white",bg="black")
    label_r1.place(x=450, y=300)

    entry_1 = Entry(root1)
    entry_1.place(x=650, y=300)


    label_r2 = Label(root1, text="Email", width=20, font=("bold", 10),fg="white",bg="black")
    label_r2.place(x=450, y=350)


    entry_2 = Entry(root1)
    entry_2.place(x=650, y=350)


    label_r3 = Label(root1, text="Password", width=20, font=("bold", 10),fg="white",bg="black")
    label_r3.place(x=450, y=400)


    entry_3 = Entry(root1, show="*")
    entry_3.place(x=650, y=400)


    label_r4 = Label(root1, text="Gender", width=20, font=("bold", 10),fg="white",bg="black")
    label_r4.place(x=450, y=450)


    var = IntVar()
    Radiobutton(root1, text="Male", padx=15, variable=var, value=1).place(x=650, y=450)
    Radiobutton(root1, text="Female", padx=15, variable=var, value=2).place(x=750, y=450)

    label_r5 = Label(root1, text="Phone:", width=20, font=("bold", 10),fg="white",bg="black")
    label_r5.place(x=450, y=500)


    entry_5 = Entry(root1)
    entry_5.place(x=650, y=500)


    # .................................................................i left here for adding data in excel file
    def submit_register_data():
        df = pd.DataFrame([entry_1.get(), entry_2.get(), entry_3.get(), entry_5.get()])
        writer = ExcelWriter('user_data_storage.xlsx')
        df.to_excel(writer, 'Sheet1', index=False)
        writer.save()
        messagebox.showinfo("sucess", "you have registered, please exit and re-launch the application, and get LoggedIn")

    def scan_image():
        # ........../\...........
        # catch picture dynamicly
        videoCaptureObject = cv2.VideoCapture(0)
        result = True
        while result:
            ret, frame = videoCaptureObject.read()
            cv2.imwrite("images/sohan_pic1.jpg", frame)
            result = False
        videoCaptureObject.release()
        cv2.destroyAllWindows()

    def clear_text():
        entry_1.delete(0, 'end')

    Button(root1, text='Submit', width=15, bg='black', fg='yellow', command=submit_register_data).place(x=450, y=550)

    Button(root1, text='Reset', width=15, bg='black', fg='yellow', command=clear_text).place(x=600, y=550)

    Button(root1, text='scan my face', width=15, bg='black', fg='yellow', command=scan_image).place(x=750, y=550)
    # ..........TO STORE THE Data in xcel file............

    root1.mainloop()
#................................***********************>>>>>>>>>>>>>>>this is where our program  starts 
#bg = PhotoImage(file="login.png")
#my_lable = Label(root, image=bg)
#my_lable.place(x=0, y=0, relwidth=1, relheight=1)
my_text = Label(root, text="JIMMY-YOUR PERSONAL ASSISTANT", font=("Algerian", 50), fg="white",bg="black")
bg = PhotoImage(file="per.png")
my_lable = Label(root, image=bg,borderwidth=0)
my_lable.place(x=550, y=150)
my_text.pack(pady=50)


my_button1 = Button(root, text="LOGIN", width=35,font=("Bahnschrift SemiLight SemiConde", 12), fg="white",bg="black",command=get_login)
my_button1.place(x=400,y=500)
my_button2 = Button(root, text="REGISTER",width=35, font=("Bahnschrift SemiLight SemiConde", 12), bg="Black",fg="white",command=get_registered)
my_button2.place(x=700,y=500)
root.mainloop()

