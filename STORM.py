from tkinter import *
import PIL.Image, PIL.ImageTk
import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import random
import roman
import pytesseract
from PIL import Image
from ecapture import ecapture as ec
import wolframalpha
import pyjokes
import shutil
import win32com.client as wincl
from urllib.request import urlopen
import subprocess



numbers= {'hundred':100, 'thousand':1000, 'lakh':100000}
a = {'name':'your email'}
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

window = Tk()

global var
global var1

var = StringVar()
var1 = StringVar()

def speak(audio):
    engine.say(audio)
    engine.runAndWait()
    
def usrname():
    speak("What should i call you sir")
    uname = takeCommand()
    speak("Welcome Mister")
    speak(uname)
    columns = shutil.get_terminal_size().columns
    speak("How can i Help you, Sir")
    
def wishme():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour <= 12:
        speak("Good Morning ")
    elif hour >= 12 and hour <= 18:
        speak("Good Afternoon ")
    else:
        speak("Good Evening ")
    speak("Myself STORM , I am your assistant") 
    
    
"""S- Speech 
T- Translating 
O- Operations Handling
R- Radical 
M- MetaAssistant"""

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        var.set("Listening...")
        window.update()
        print("Listening...")
        r.pause_threshold = 1
        r.energy_threshold = 400
        audio = r.listen(source)
    try:
        var.set("Recognizing...")
        window.update()
        print("Recognizing")
        query = r.recognize_google(audio, language='en-in')
    except Exception as e:
        return "None"
    var1.set(query)
    window.update()
    return query

def play():
    btn2['state'] = 'disabled'
    btn0['state'] = 'disabled'
    btn1.configure(bg = 'red')
    wishme()
    usrname()
    while True:
        btn1.configure(bg = 'red')
        query = takeCommand().lower()
        if 'exit' in query:
            var.set("Bye sir")
            btn1.configure(bg = '#5C85FB')
            btn2['state'] = 'normal'
            btn0['state'] = 'normal'
            window.update()
            speak("Bye sir")
            break

        elif 'wikipedia' in query:
            if 'open wikipedia' in query:
                webbrowser.open('wikipedia.com')
            else:
                try:
                    speak("searching wikipedia")
                    query = query.replace("according to wikipedia", "")
                    results = wikipedia.summary(query, sentences=2)
                    speak("According to wikipedia")
                    speak(results)
                except Exception as e:
                    speak('sorry sir could not find any results')

        elif 'open youtube' in query:
            speak('opening Youtube')
            webbrowser.open("www.youtube.com")

        elif 'open blackboard' in query:
            speak('opening blackboard')
            webbrowser.open("https://cuchd.blackboard.com/")

        elif 'open google' in query:
            speak('opening google')
            webbrowser.open("https://www.google.com/")

        elif 'hello' in query:
            speak("Hello Sir")
            
        elif 'open college profile' in query:
            speak("Here you go to cuims")
            webbrowser.open("https://uims.cuchd.in/uims/")
            
        elif 'open wolframalpha' in query:
            speak("Here you go")
            webbrowser.open("www.wolframalpha.com")
            
        elif 'open college website' in query:
            speak("Here you go to Chandigarh University webpage")
            webbrowser.open("https://www.cuchd.in")
            
        elif 'open facebook' in query:
            speak('opening facebook')
            webbrowser.open('www.facebook.com')
            
        elif 'open stackoverflow' in query:
            speak('opening stackoverflow')
            webbrowser.open('www.stackoverflow.com')
            
        elif 'open instagram' in query:
            speak('opening instagram')
            webbrowser.open('www.instagram.com')
            
        elif 'open geekforgeeks' in query:
            speak('opening geekforgeeks')
            webbrowser.open('www.geekforgeeks.com')
            
        elif 'joke' in query:
            print(pyjokes.get_joke())
            speak(pyjokes.get_joke())
            
        elif 'shutdown system' in query:
            speak("Hold On a Sec ! Your system is on its way to shut down")
            subprocess.call('shutdown / p /f')
            
        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("Recycle Bin Recycled")
            
        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop STORM from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)
            
        elif ('play music' in query) or ('change music' in query):
            speak('Here are your favorites')
            music_dir = 'directory to your song folder'
            songs = os.listdir(music_dir)
            n = random.randint(0,2)
            os.startfile(os.path.join(music_dir, songs[n]))

        elif 'the time' in query:
            strtime = datetime.datetime.now().strftime("%H:%M:%S")
            speak("Sir the time is %s" %strtime)

        elif 'the date' in query:
            strdate = datetime.datetime.today().strftime("%d %m %y")
            speak("Sir today's date is %s" %strdate) 
        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])
			
        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])
            
        elif 'thank you' in query:
            speak("Welcome Sir")
            
        elif "what is" in query or "who is" in query:
            try:		
                speak('Searching Wikipedia...')
                query = query.replace("wikipedia", "")
                results = wikipedia.summary(query, sentences = 3)
                speak("According to Wikipedia")
                print(results)
                speak(results)
            except Exception:
                speak("Sorry no results found")
                

        elif 'temperature' in query:
            app_id = "wolframalpha_appid"
            client = wolframalpha.Client(app_id)
            res = client.query(query)
            answer = next(res.results).text
            print("The answer is " + answer) 
            speak("The answer is " + answer) 
            
        elif 'can you do for me' in query:
            speak('I can do multiple tasks for you sir. tell me whatever you want to perform sir')

        elif 'open media player' in query:
            speak("opening V L C media player")
            path = "C:/Program Files/VideoLAN/VLC/vlc.exe" 
            os.startfile(path)

        elif 'your name' in query:
            speak('''myself STORM sir
                        S stands for Speech
                        T stands for Translating
                        O stands for Operation handling
                        R stands for Radical
                        M stands for MetaAssistant
                  ''')

        elif 'who created you' in query:
            speak('My Creator is Mohit Gupta')

        elif 'say hello' in query:
            speak('Hello Everyone! My self STORM')

        elif 'open VS Code' in query:
            speak("Opening VS Code")
            path = "C:/Users/hp/AppData/Local/Programs/Microsoft VS Code/Code.exe" 
            os.startfile(path)

        elif 'open browser' in query:
            speak("Opening Browser")
            path = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
            subprocess.call(path)
		
        elif "open python" in query:
            speak('opening python ')
            PYTHON=r"C:/Users/hp/AppData/Local/Programs/Python/Python37/python.exe"
            subprocess.call(PYTHON) 

        elif 'click photo' in query:
            ec.capture(0, "Storm Camera ", "img.jpg")
        
        elif 'read the photo' in query: 
            try:
                im = Image.open('OIP.jpg')
                text = pytesseract.image_to_string(im)
                speak(text)
            except Exception as e:
                print("Unable to read the data")
                print(e)
        else :
            try:
                app_id = "wolframalpha_appid"
                client = wolframalpha.Client(app_id)
                res = client.query(query)
                answer = next(res.results).text 
                print("The answer is " + answer) 
                speak("The answer is " + answer)
            except Exception:
                speak("Sorry, no result found")
                print("Sorry, no result found")
            pass
                

def update(ind):
    frame = frames[(ind)%100]
    ind += 1
    label.configure(image=frame)
    window.after(100, update, ind)

label2 = Label(window, textvariable = var1, bg = '#FAB60C')
label2.config(font=("Courier", 20))
var1.set('User Said:')
label2.pack()

label1 = Label(window, textvariable = var, bg = '#ADD8E6')
label1.config(font=("Courier", 20))
var.set('Welcome')
label1.pack()

frames = [PhotoImage(file='Assistant.gif',format = 'gif -index %i' %(i)) for i in range(100)]
window.title('STORM')
window.configure(bg='cyan')
label = Label(window, width = 500, height = 500, bg='cyan')
label.pack()
window.after(0, update, 0)

btn0 = Button(text = 'WISH ME',width = 20, command = wishme, bg = '#5C85FB')
btn0.config(font=("monotype corsiva", 12))
btn0.pack()
btn1 = Button(text = 'PLAY',width = 20,command = play, bg = '#5C85FB')
btn1.config(font=("monotype corsiva", 12))
btn1.pack()
btn2 = Button(text = 'EXIT',width = 20, command = window.destroy, bg = '#5C85FB')
btn2.config(font=("monotype corsiva", 12))
btn2.pack()


window.mainloop()
