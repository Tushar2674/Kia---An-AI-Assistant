import speech_recognition as sr    #To convert speech into text
import pyttsx3                     #To convert text into speech
import datetime                    #To get the date and time
import wikipedia                   #To get information from wikipedia
import webbrowser                  #To open websites
import os                          #To open files
import time                        #To calculate time
import subprocess                  #To open files
import pyjokes                     #For jokes
from playsound import playsound    #To play sounds 
import sys 
import operator
import pyautogui
import cv2
import sounddevice as sd
import pyaudio
import numpy as np
import requests
import pywhatkit as wk 
from tkinter import *
import ctypes
import keyboard                    #To add keyboard activation
name_assistant = "Kia" #The name of the assistant


engine = pyttsx3.init('sapi5')   #'sapi5' is the argument you have to use for windows, I am not sure what it is for Mac and Linux
voices = engine.getProperty('voices')  #To get the voice
engine.setProperty('voice', voices[1].id) #defines the gender of the voice. 
engine.setProperty("rate",150)#speed of speaking


# Note: voices[1].id sets it to female and voices[0].id sets it to male

def speak(text):
    engine.say(text)
    print(name_assistant ," : " ,  text)
    engine.runAndWait()
def wishMe():
    hour=datetime.datetime.now().hour
    if hour>=0 and hour<12:
        speak("Hello,Good Morning")
   
    elif hour>=12 and hour<18:  # This uses the 24 hour system so 18 is actually 6 p.m 
        speak("Hello,Good Afternoon")
 
    else:
        speak("Hello,Good Evening")

month_names = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
ordinalnames = [ '1st', '2nd', '3rd', ' 4th', '5th', '6th', '7th', '8th', '9th', '10th', '11th', '12th', '13th', '14th', '15th', '16th', '17th', '18th', '19th', '20th', '21st', '22nd', '23rd','24rd', '25th', '26th', '27th', '28th', '29th', '30th', '31st']


def date():
    now = datetime.datetime.now()
    my_date = datetime.datetime.today()

    month_name = now.month
    day_name = now.day
    month_names = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
    ordinalnames = [ '1st', '2nd', '3rd', ' 4th', '5th', '6th', '7th', '8th', '9th', '10th', '11th', '12th', '13th', '14th', '15th', '16th', '17th', '18th', '19th', '20th', '21st', '22nd', '23rd','24rd', '25th', '26th', '27th', '28th', '29th', '30th', '31st'] 
    

    speak("Today is "+ month_names[month_name-1] +" " + ordinalnames[day_name-1] + '.')



def note(text):
    date = datetime.datetime.now()
    file_name = str(date).replace(":", "-") + "-note.txt"
    with open(file_name, "w") as f:
        f.write(text)

    subprocess.Popen(["notepad.exe", file_name])

def get_audio(): 

    r = sr.Recognizer() 
    audio = '' 

    with sr.Microphone() as source: 

        print("Listening")
        r.threshold =1 
        audio = r.listen(source) 
        
    try:
        print("Recognizing....")
        text = r.recognize_google(audio, language ='en-in') 
        print('You' + ': '+ text)
        return text


    except:
        print("I request your Pardon....")
        speak("I request your Pardon....")
        return "None"
    return text

def Process_audio():
    wishMe()
    if __name__=='__main__':
        while True:
            
            #operation 1 
            app_string = ["open word", "open powerpoint", "open excel", "open visual studio code ","open notepad",  "open chrome"]
            app_link = [r'\Microsoft Office 2013\Word 2013.lnk',r'\Microsoft Office 20413\PowerPoint 2013.lnk', r'\Microsoft Office 2013\Excel 2013.lnk', r'\Visual Studio Installer.lnk', r'\Accessories\Notepad.lnk', r'\Google Chrome.lnk' ]
            app_dest = r'c:\ProgramData\Microsoft\Windows\Start Menu\Programs' # Write the location of your file
            statement =get_audio().lower()
            
            results = ''
            
            #operation 2
            if "today date" in statement :
                date()
            #operation 3    
            elif "your name" in statement :
                print(" My name is "+name_assistant)
                speak (" My name is "+name_assistant)
            #operation 4
            elif "calculator" in statement :
                speak("calculator mode ON")
                r=sr.Recognizer()
                with sr.Microphone() as source:
                    speak("Ready")
                    speak("Listening...")
                    r.adjust_for_ambient_noise(source)
                    audio=r.listen(source)
                mystring=r.recognize_google(audio)
                print(mystring)
                def get_operator(op):
                    return {
                        "+" : operator.add,
                        "-" : operator.sub,
                        "x" : operator.mul,
                        "/" : operator.__truediv__,
                    }[op]
                def eval_binary_expr(op1,oper,op2):
                    op1,op2=int(op1),int(op2)
                    return get_operator(oper)(op1,op2)
                speak("your result is")
                speak(eval_binary_expr(*(mystring.split())))
            #operation 5
            elif "what is my ip address"in statement:
                speak("checking")
                try:
                    ip=requests.get("https://api.ipify.org").text
                    print(ip)
                    speak("your ip address is")
                    speak(ip)
                except Exception as e:
                    speak("your network is slow , please try again after some time ")
            #operation 6
            elif "hello" in statement or "hi" in statement:
                wishMe()               
            #operation 7
            elif "good bye" in statement or "ok bye" in statement :
                speak('Your personal assistant ' + name_assistant +' is shutting down, Good bye')
                sys.exit()
            #operation 8
            elif 'just open google' in statement:
                webbrowser.open_new_tab("google.com")
                speak("google is open now")
            #operation 9
            elif 'open google' in statement:
                    speak("google is open now ")
                    speak("what do you want to search ")
                    
                    def wik(w):
                        
                    
                        if 'who is' in w:
                            w = w.replace("who is", "")
                            results = wikipedia.summary(w, sentences = 1)
                            speak("According to google ")
                            speak(results)
                            speak("do you want to search more ")
                            b0=get_audio().lower()
                            if "yes" in b0:
                                speak("what do you want to search ")
                                c0=get_audio()
                                wik(c0)
                            else:  
                                breakpoint
                                speak("google is close now")
                            
                        if 'what is' in w:
                            w = w.replace("what is", "")
                            result = wikipedia.summary(w, sentences = 1)
                            speak("According to google")
                            speak(result)
                            speak("do you want to search more ")
                            b1=get_audio().lower()
                            if "yes" in b1:
                                speak("what do you want to search ")
                                c1=get_audio()
                                wik(c1)
                            else:   
                                breakpoint
                                speak("google is close now")
                                
                        else:
                            w = w.replace("tell me ", "")
                            results = wikipedia.summary(w, sentences = 1)
                            speak("According to google ")
                            speak(results)
                            speak("do you want to search more ")
                            b2=get_audio().lower()
                            if "yes" in b2:
                                speak("what do you want to search ")
                                c2=get_audio()
                                wik(c2)
                            else:   
                                breakpoint
                                speak("google is close now")
                                
                    a=get_audio()
                    wik(a)

            #operation 11
            elif "screenshot" in statement:
                speak('tell me a name for the file')
                name = get_audio().lower()
                time.sleep(3)
                img = pyautogui.screenshot() 
                img.save(f"{name}.png") 
                speak("screenshot saved")

            #operation 12
            elif "shutdown the system" in statement:
                os.system("shutdown /s /t 5")
            #operation 13    
            elif "lock the system" in statement:
                os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
            #operation 14
            elif "restart the system " in statement:
                os.system("shutdown /r /t 5")
            #operation 15
            elif "hibernate the system " in statement:
                os.system("windll32.exe powerprof.dill,SetSuspendState 0,1,0")
            #operation 16
            elif 'joke' in statement:
                speak(pyjokes.get_joke())    
            #operation 17
            elif 'search on youtube' in statement:
                speak("what will you like to watch")
                qrry=get_audio()
                wk.playonyt(f"{qrry}")
                time.sleep(1)

            #operation 18
            elif 'open instagram' in statement:
                webbrowser.open_new_tab("https://www.instagram.com")
                speak("instagram is open now")
                time.sleep(5)
            #operation 19
            elif 'open web whatsapp' in statement:
                webbrowser.open_new_tab("https://web.whatsapp.com")
                speak("webwhatsapp is open now")
                time.sleep(5)
            #operation 20
            elif 'close google' in statement:
                os.system("taskkill /f /im msedge.exe")
            #operation 21
            if 'open gmail' in statement:
                    webbrowser.open_new_tab("mail.google.com")
                    speak("Google Mail open now")
                    time.sleep(5)
            #operation 22
            if 'open netflix' in statement:
                    webbrowser.open_new_tab("netflix.com/browse") 
                    speak("Netflix open now")
            #operation 23
            if 'open prime video' in statement:
                    webbrowser.open_new_tab("primevideo.com") 
                    speak("Amazon Prime Video open now")
                    time.sleep(5)

            if app_string[0] in statement:
                os.startfile(app_dest + app_link[0])

                speak("Microsoft office Word is opening now")

            if app_string[1] in statement:
                os.startfile(app_dest + app_link[1])
                speak("Microsoft office PowerPoint is opening now")

            if app_string[2] in statement:
                os.startfile(app_dest + app_link[2])
                speak("Microsoft office Excel is opening now")
        
            if app_string[3] in statement:

                os.startfile(app_dest + app_link[3])
                speak("Visual Studio Code is opening now")


            if app_string[4] in statement:
                os.startfile(app_dest + app_link[4])
                speak("Notepad is opening now")
        
            if app_string[5] in statement:
                os.startfile(app_dest + app_link[5])
                speak("Google chrome is opening now")
                       
            #operation 24
            if 'news' in statement:
                news = webbrowser.open_new_tab("https://timesofindia.indiatimes.com/city/mangalore")
                speak('Here are some headlines from the Times of India, Happy reading')
                time.sleep(6)
            #operation 25
            if 'cricket' in statement:
                news = webbrowser.open_new_tab("cricbuzz.com")
                speak('This is live news from cricbuzz')
                time.sleep(6)
            #operation 26
            if 'corona' in statement:
                news = webbrowser.open_new_tab("https://www.worldometers.info/coronavirus/")
                speak('Here are the latest covid-19 numbers')
                time.sleep(6)
            #operation 27
            if 'time' in statement:
                strTime=datetime.datetime.now().strftime("%H:%M:%S")
                speak(f"the time is {strTime}")
            #operation 28
            if 'date' in statement:
                date()
            #operation 29
            if 'who are you' in statement or 'what can you do' in statement:
                    speak('I am '+name_assistant+' your personal assistant. I am programmed to minor tasks like opening youtube, google chrome, gmail and search wikipedia etcetra. I am very intelligent and trying to be updated day by day .') 
            #operation 30
            if "who made you" in statement or "who created you" in statement or "who discovered you" in statement:
                speak("I was built by Tushar")
            #operation 31
            if 'make a note' in statement:
                statement = statement.replace("make a note", "")
                note(statement)
            #operation 32
            if 'note this' in statement:    
                statement = statement.replace("note this", "")
                note(statement)   
            #operation 33
            elif 'open youtube' in statement:
                webbrowser.open_new_tab("https://www.youtube.com")
                speak("youtube is open now")
                time.sleep(5) 
            #operation 34 
            elif "volume up" in statement:
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
                pyautogui.press("volumeup")
            #operation 35
            elif "volume down" in statement:
                pyautogui.press("volumedown") 
                pyautogui.press("volumedown") 
                pyautogui.press("volumedown") 
                pyautogui.press("volumedown") 
                pyautogui.press("volumedown") 
                pyautogui.press("volumedown") 
                pyautogui.press("volumedown") 
                pyautogui.press("volumedown") 
                pyautogui.press("volumedown") 
                pyautogui.press("volumedown") 
                pyautogui.press("volumedown") 
                pyautogui.press("volumedown") 
                pyautogui.press("volumedown")  
            elif "open command prompt" in statement:
                os.system("start cmd") 
            elif "close command prompt" in statement:
                os.system("taskkill /f /im cmd.exe")
            elif "mute" in statement:
                pyautogui.press("volumemute")
            elif "refresh" in statement:
                pyautogui.moveTo(630,560, 1)
                pyautogui.click(x=630, y=560, clicks=1, interval=0.5, button='right')
                pyautogui.moveTo(1030,690, 1)
                pyautogui.click(x=1030, y=690, clicks=1, interval=0.5, button='left')
            elif "scroll down" in statement:
                pyautogui.scroll(1000)
            elif 'type' in statement:
                statement = statement.replace("type", "")
                pyautogui.write(f"{statement}")
            elif "welcome" in statement:
                wishMe()
                pyautogui.press("win")
                time.sleep(2)
                pyautogui.typewrite("C:\\Users\\admin\\Desktop\\Project_KIA.pptx")
                time.sleep(3)
                pyautogui.press("enter")
                time.sleep(6)
                pyautogui.press("f5")
                time.sleep(1)
                speak("I am "+name_assistant+" ,your personal assistant . I am  meticulously designed for user convenience, seamlessly facilitates various tasks with unparalleled ease ")
                time.sleep(1)
                pyautogui.press("right")
                time.sleep(1)
                speak("In today's fast-paced world, personal assistants have become an essential tool for managing our daily tasks and increasing productivity . I am a cutting-edge personal assistant designed to provide seamless assistance and convenience to users")
                time.sleep(1)
                pyautogui.press("right")
                time.sleep(1)
                speak("users can effortlessly launch applications such as Word, PowerPoint, Excel,notepad,netflix, whatsapp, google and many more")
                time.sleep(1)
                pyautogui.press("right")
                time.sleep(1)
                speak(" i do a warm greeting upon activation, ensuring a welcoming interaction . Alongside assisting with date, time, and calculation inquiries")
                time.sleep(1)
                pyautogui.press("right")
                time.sleep(1)
                speak(" With the ability to adjust volume levels and even operate the camera to capture snapshots, I versatility knows no bounds")
                time.sleep(1)
                pyautogui.press("right")
                time.sleep(1)
                speak("users can surf the internet using google and youtube . I do offers a streamlined experience. Users can initiate searches for specific videos on YouTube effortlessly, enhancing productivity")
                time.sleep(1)
                pyautogui.press("right")
                time.sleep(1)
                speak("i grants users control over system functions like shutdown, restart, lock, and hibernate, enhancing operational efficiency")
                time.sleep(1)
                pyautogui.press("right")
                time.sleep(1)    
                speak("However, My most remarkable feature lies in mine intuitive typing functionality, where users simply articulate their thoughts, and it transcribes seamlessly. This comprehensive suite of capabilities underscores mine indispensable role in enhancing user productivity and convenience")
                pyautogui.press("right")
                time.sleep(1)    
                speak("thank you")
                time.sleep(1)
                pyautogui.press("esc")
                time.sleep(1)
                pyautogui.hotkey("alt","f4")
                speak(results)
            elif "camera" in statement:
                pyautogui.press("win")
                time.sleep(2)
                pyautogui.typewrite("camera")
                time.sleep(3)
                pyautogui.press("enter")
                time.sleep(1)
                speak("what do you want to have a photo or a video")
                z0=get_audio().lower()
                if "photo" in z0:
                    speak("clicking photograph in 5 seconds")
                    time.sleep(3)
                    speak("clicking")
                    time.sleep(1)
                    pyautogui.press("space")
                if "video" in z0:
                    pyautogui.moveTo(1840,370)
                    time.sleep(2)
                    pyautogui.leftClick()
                    time.sleep(2)
                    speak("Recording video in 5 seconds")
                    time.sleep(3)
                    speak("Recording")
                    time.sleep(1)
                    pyautogui.press("space")

            elif "stop" in statement:

                breakpoint
            elif "open quick finder" in statement:
                pyautogui.press("win")
                time.sleep(2)
                pyautogui.typewrite("C:\\Users\\admin\\Desktop\\Gaurav\\main.py")
                time.sleep(2)
                pyautogui.press("enter")
                time.sleep(2)
                pyautogui.hotkey("ctrl","f5")
                time.sleep(2)
                pyautogui.press("downkey")
                time.sleep(2)
                pyautogui.press("enter")

            
           
root = Tk()
var = IntVar()
def toggle():
    if var.get() == 1:
        Process_audio()
        
    else:
        print("Button is OFF")
        get_audio="stop"
checkbutton = Checkbutton(root, text="Toggle", variable=var, command=toggle)
checkbutton.pack()

root.mainloop()