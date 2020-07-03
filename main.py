import pygame
import pyautogui
import psutil
import string
import tkinter as tk
from tkinter import messagebox
from tkinter import *
from ttkthemes import *
import pyttsx3
import webbrowser
import smtplib
import random
import speech_recognition as sr
import wikipedia
import datetime
import cv2
import wolframalpha
import os
import sys
import requests
import ctypes
import re
import pygame
import wmi
import pyautogui
import tkinter as tk
from tkinter import messagebox
from tkinter import *
from ttkthemes import *
import time
from random import choice as ch


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[len(voices) - 1].id)


def speak(audio):
    print('Eliza: ' + audio)
#    engine.say(audio)
 #   engine.runAndWait()
    textbox.configure(state=NORMAL)
    textbox.insert('end', '\nEliza:\t')
    textbox.insert('end', audio)
    textbox.configure(state=DISABLED)
    textbox.delete(0.0,'end')
    textbox.update()
    engine.say(audio)
    engine.runAndWait()

def timeSett():
    currentH = int(datetime.datetime.now().hour)
    if currentH >= 0 and currentH < 12:
        speak('Good Morning!')

    if currentH >= 12 and currentH < 18:
        speak('Good Afternoon!')

    if currentH >= 18 and currentH != 0:
        speak('Good Evening!')


def myCommand():

    r = sr.Recognizer()
    while(True):
        with sr.Microphone() as source:
            print("Listening...")
            r.pause_threshold = 1
            r.adjust_for_ambient_noise(source, duration=1)
            audio = r.listen(source)
        try:
            query = r.recognize_google(audio)
            print('User: ' + query + '\n')
            textbox.configure(state=NORMAL)
            textbox.insert('end', '\nYOU:\t')
            textbox.insert('end', query)
            textbox.update()
            textbox.delete(0.0, 5.0)
            textbox.configure(state=DISABLED)
            break
        except sr.UnknownValueError:
            speak('Sorry! I didn\'t get that! Try again!')
            #textbox.insert('end', '\nCommand:\t')
            #query = str(input('Command: '))

    return query


def process():

    while True:

        query = myCommand()
        query = query.lower()

        if 'open youtube' in query:
            speak("okay! opening youtube...")
            webbrowser.open('www.youtube.com')

        elif 'open google' in query:
            speak('okay! opening google...')
            webbrowser.open('www.google.co.in')

        elif 'open gmail' in query:
            speak('okay! Opening your gmail...')
            webbrowser.open('www.gmail.com')

        elif "what\'s up" in query or 'how are you' in query:
            stMsgs = ['Just doing my thing!', 'I am fine!', 'Nice!', 'I am nice and full of energy']
            speak(random.choice(stMsgs))

        elif 'email' in query or 'mail' in query:
            speak('Who is the recipient? ')
            recipient = myCommand()

            if 'send it to me' in recipient:
                try:
                    speak('What should I say? ')
                    content = myCommand()

                    server = smtplib.SMTP("smtp.gmail.com", 587)
                    server.ehlo()
                    server.starttls()
                    server.login("*******", "******")
                    server.sendmail("*******", "", content)
                    server.close()
                    speak('Email sent!')

                except:
                    speak('Sorry! I am unable to send your message at this moment!')

            elif 'rudrajit' in recipient:
                    try:
                        speak('What should I say? ')
                        content = myCommand()

                        server = smtplib.SMTP("smtp.gmail.com", 587)
                        server.ehlo()
                        server.starttls()
                        server.login("*******", "*******")
                        server.sendmail("*******", "rudrajit1729@gmail.com", content)
                        server.close()
                        speak('Email sent!')


                    except:
                        speak('Sorry! I am unable to send your message at this moment!')


        elif 'nothing' in query or 'abort' in query or 'stop' in query:
            pygame.mixer.music.load('sound.wav')
            pygame.mixer.music.play(0)
            speak('okay')
            speak('Bye, have a good day.')
            sys.exit()


        elif 'hello' in query or 'hi' in query or 'hi Eliza' in query or 'hello Eliza' in query or 'hey Eliza' in query:
            speak('Hello there! How may I help you?')

        elif 'bye' in query:
            pygame.mixer.music.load('sound.wav')
            pygame.mixer.music.play(0)
            speak('Bye, have a good day.')
            sys.exit()

        elif 'play music' in query:
            music_folder = "music"
            music = ["music1", "music2", "music3", "music4", "music5"]
            random_music = random.choice(music) + '.mp3'
            speak('Okay, here is your music! Enjoy!')
            os.system(random_music)



        elif 'joke' in query:
            res = requests.get(
                'https://icanhazdadjoke.com/',
                headers={"Accept": "application/json"}
            )
            if res.status_code == requests.codes.ok:
                speak(str(res.json()['joke']))
            else:
                speak('oops!I ran out of jokes')

        elif 'decrease' in query or 'decrease brightness' in query:
            speak("decreasing brightness..")
            dec = wmi.WMI(namespace='wmi')
            methods = dec.WmiMonitorBrightnessMethods()[0]
            methods.WmiSetBrightness(10, 0)
            #return
        elif 'increase ' in query or 'increase brightness' in query:
            speak("increasing brightness..")
            ins = wmi.WMI(namespace='wmi')
            methods = ins.WmiMonitorBrightnessMethods()[0]
            methods.WmiSetBrightness(100, 0)
            #return

        elif 'screenshot' in query or 'screen shot' in query or 'snapshot' in query or "Take a screenshot:" in query:
            name = random.randint(1000, 300000)
            speak("Taking Screenshot..")
            speak('Check your screenshots folder, I saved there')
            pic = pyautogui.screenshot()
            pic.save("./screenshots/screenshot"+str(name)+".jpg")


        elif "news" in query:
            my_url = " https://newsapi.org/v1/articles?source=bbc-news&sortBy=top&apiKey=9113b9445d8c4e399c473f41890d6126"
            my_open_bbc_page = requests.get(my_url).json()
            my_article = my_open_bbc_page["articles"]
            my_results = []
            for ar in my_article:
                my_results.append(ar["title"])
            speak("Here is the top ten headlines from BBC world.")
            for i in range(len(my_results)):
                print(i + 1)

                engine.say(my_results[i])
                engine.runAndWait()

        elif "lock my pc" in query:
            speak("ok!")
            ctypes.windll.user32.LockWorkStation()
            return
        elif "restart" in query or "system restart" in query or "reboot" in query:
            speak("ok!")
            speak("system restarting..")
            os.system("res.bat")

        elif "shutdown" in query or "system shutdown" in query or "shutdown the system" in query:
            #speak("is it sure!")
           # myCommand()
            speak("ok!")
            speak("system shutdowning..")
            os.system("shut.bat")
           # if "yes" in query:

                #os.system("shutdown /s /f /t 1")
           # if "no" in query:
             #   speak("ok!")
            #    return

        elif 'open website' in query:

            reg_ex = re.search('open website (.+)', query)
            if reg_ex:
                domain = reg_ex.group(1)
                url = 'https://www.' + domain
                webbrowser.open(url)
                speak("Opening website...")
                speak('Done!')

            else:
                pass

        elif "search the web" in query:
            speak("Sure. What do you want to search?")
            search = myCommand()
            try:
                webbrowser.open('http://google.com/search?q=' + search)
            except:
                speak("Can't search at the moment! Check your internet connection.")
        else:
            query = query
            speak('Searching...')
            try:
                try:
                    client = wolframalpha.Client("A4VGW6-6YT3XAW3X5")
                    res = client.query(query)
                    results = next(res.results).text
                    #speak('Got it! ')
                    speak(results)

                except:
                    results = wikipedia.summary(query, sentences=2)
                    speak('Got it!')
                    #speak('Wikipedia says: ')
                    engine.say(results)
                    engine.runAndWait()

            except:
                speak("Sorry! can't search at the moment. Opening google... ")
                webbrowser.open('http://google.com/search?q=' + query)

        speak('Next Command!')


def start():
    pygame.init()
    pygame.mixer.music.load('sound.wav')
    pygame.mixer.music.play(0)
    time.sleep(1)
    timeSett()
    speak('This is Eliza at your service!')
    rand = ["What can I do for you?", "How can I help?", "Hi there, what can I do ?", "Do you need help?"]
    voice = ch(rand)
    speak(voice)
    process()

if __name__ == '__main__':

    root = tk.Tk()
    frame = tk.Frame(root).pack()
    root.title(" Eliza")
    root.geometry('500x750')
    root.resizable(0,0)
    root.iconbitmap(r'logo.ico')
    style = ttk.Style()
    style.theme_use('winnative')
    root.config(background='black')
    photo = PhotoImage(file='backg.png')
    c = tk.Canvas(root, bg='black', width=900, height=350)
    c.create_image(250, 380, image=photo)
    c.pack()
    #c.create_oval(100, 100, 400, 400, width=3, fill="")
    label = tk.Label(root, text="P  E  R  S  O  N  A  L     A  S  S  I S T  A  N  T").pack()
    startBotton = tk.Button(root, height=0, width=20, relief=FLAT, fg='black', bg="red", text="--- S  T  A  R  T ---",activeforeground="black", command=lambda: start())
    #play = lambda: PlaySound('Sound.wav', SND_FILENAME)
    #startbutton = Button(root, command=play)
    startBotton.pack(side=tk.TOP)
    c1 = tk.Canvas(root, bg='black', relief = FLAT,width=500, height=175)
    #photo1 = PhotoImage(file='backg.png')
    #c1.create_image(-58,-40, image=photo1,anchor='nw')
    #c1.pack()
    def about():
        tk.messagebox.showinfo('About Eliza! ','Eliza is a Virtual Voice Assistant designed for a service'
                                              ' which can access and control the functions and the web in the system.\n')
    menubar = Menu(root)
    root.config(menu=menubar)
    subMenu = Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Help", menu=subMenu)
    subMenu.add_command(label="About", command=about)
    subMenu.add_separator()
    subMenu.add_command(label="Exit", command=quit)
    exitButton = tk.Button(root, bg="black", relief=FLAT, command=quit)
    photo2 = tk.PhotoImage(file="quit.png", )
    exitButton.config(image=photo2, width="100", height="15")
    exitButton.pack(side=tk.BOTTOM)
    textbox = tk.Text(root, bg='black',fg="white",font=('sans',12,'bold','italic'), height=8, width=50)
    #textbox.insert(END, "T H I S  I S  R U B Y")
    #textbox.image_create(END)
    textbox.pack(fill="both", expand=True)
    time.sleep(1)
    speak("Initializing..")
    # time.sleep(1)
    # speak("Getting ready..")
    #time.sleep(1)
    speak("Done")
    root.mainloop()
