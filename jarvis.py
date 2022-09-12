from win32com.client import Dispatch
import datetime
import speech_recognition as sr
import webbrowser
import wikipedia
import os
import random
import subprocess


list1 = ["Hello","hi"]
a = random.choice(list1)
def speak(audio) :
    speak = Dispatch(("sapi.spvoice"))
    speak.speak(audio)

if __name__ == '__main__':
   speak(a)



def wish_me():
    hour =datetime.datetime.now().strftime("%H:%M:%S")
    #speak(f"The time is {hour}")
    if hour<"10" :
        speak("Good morning, what can i do for you sir")

    elif hour<"16" :
        speak("GOod afternoon, what can i do for you sir")

    else :
        speak("Good evening, what can i do for you sir")
if __name__ == '__main__':
    wish_me()
    


def TakeCommand() :
    r = sr.Recognizer()
    with sr.Microphone() as source :
        print("Listening...")
        # r.pause_threshold = 1
        audio = r.listen(source)

    try :
        print("Recognizing...")
        Query = r.recognize_google(audio, language ='en-in')    
        print("User said",Query) 
        speak(f"You have said {Query}")
       
    except Exception as e :
        print("Say that again...")
        speak("Say that again please")
        return "None"    
    return Query

        
while(True) :  
    if __name__ == '__main__':
        Query = TakeCommand().lower()

    if "wikipedia" in Query :
        speak("Searching wikipedia...")
        Query = Query.replace("wikipedia","")
        results = wikipedia.summary(Query,sentences=2)
        speak("According to wikipedia")
        print(results)
        speak(results)

    elif ("my folder" or "pradeep") in Query :
        os.startfile("E:\Pr@deep")

       
    elif "movies" in Query :
        os.startfile("E:")
        
        
    elif "play song" in Query:
        music_dir = 'E:\\favourate'
        songs = os.listdir(music_dir)
        x = range(100)
        b = random.choice(x)
        print(songs)
        os.startfile(os.path.join(music_dir,songs[b])) 

    elif "who are you" in Query :
        speak("I am jarvis")

    elif "youtube" in Query :
        webbrowser.open("youtube.com")

    elif "google" in Query :
        webbrowser.open("google.com")

    elif "stack overflow" in Query :
        webbrowser.open("stackoverflow.com")

    elif "time" in Query :
        time =datetime.datetime.now().strftime("%H:%M:%S")
        print(time)
        speak(time)
    
    elif "add reminder" in Query :
        speak("what should i remember?")
        data = TakeCommand()
        f = open("helo.txt",'w')
        f.write(data)
        f.close()


    elif "read reminder" in Query :
        f = open("helo.txt",'r')
        i = f.read()
        speak(f"You have said {i}")
        f.close    

        