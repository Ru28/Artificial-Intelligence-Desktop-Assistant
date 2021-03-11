import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import random
import smtplib

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice',voices[1].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour=(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!") 
    elif hour>=12 and hour<18:
        speak("Good afternoon!")
    elif hour>=18 and hour<22:
        speak("Good Evening!")
    else:
        speak("Good Night!")
        
    speak("I am Jarvis Artificial Intelligence desktop , Please tell me how can I help you")

def takeCommand():
    #  it takes microphone input from the user and returns string output
    r=sr.Recognizer()

    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold=1
        audio=r.listen(source)


    try:
        print("Recongnizing...")
        query=r.recognize_google(audio,language="en-US")
        print(f"user said: {query}\n")


    except Exception as e:
        print("Say that again Please...")
        return "None"
    return query
        
def sendEmail(to,content):
    server= smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()
    server.login('youremail@gmail.com','your-password')
    server.sendmail('youremail@gmail.com',to,content)
    server.close()


if __name__=="__main__":
    wishMe()
    while True:
        query=takeCommand().lower()

        #Logic for executing task based on your query
        if 'wikipedia' in query:
            try:
                speak('Searching Wikipedia...')
                query = query.replace("wikipedia","")
                results  = wikipedia.summary(query,sentences=2)
                speak("According to wikipedia")
                print(results)
                speak(results)
            except Exception as e:
                print(e)
                print("Please Say again!")



        elif 'open youtube' in query:
            webbrowser.open("youtube.com")
        elif 'open linkedin' in query:
            webbrowser.open("linkedin.com")
        elif 'open facebook' in query:
            webbrowser.open("facebook.com")
        elif 'open instagram' in query:
            webbrowser.open("instagram.com")
        elif 'open twitter' in query:
            webbrowser.open("twitter.com")
        elif 'open google' in query:
            webbrowser.open("google.com")
        elif 'open stackoverflow' in query:
            webbrowser.open("stackoverflow.com")
        elif 'open language translator' in query:
            webbrowser.open("https://translate.google.co.in/")
        elif 'open snapchat' in query:
            webbrowser.open("snapchat.com")
        elif 'open jp university website' in query:
            webbrowser.open("https://www.juet.ac.in/")
        elif 'how to make samosa' in query:
            webbrowser.open("https://www.indianhealthyrecipes.com/samosa-recipe-make-samosa/")

        elif 'play music' in query:
            music_dir='D:\\I.T.Wala\\Music'
            songs=os.listdir(music_dir)
            random_number=random.randint(0,40)
            os.startfile(os.path.join(music_dir,songs[random_number]))

        elif 'the time' in query:
            strTime=datetime.datetime.now().strftime("%H:%M:%S")
            print(f"Sir,The time is {strTime}")
            speak(f"Sir, The time is {strTime}")

        elif 'open visual studio' in query:
            codePath="C:\\Users\\HP\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)
        
        elif 'email to xyzPerson' in query:
            try:
                speak("what should I say?")
                content=takeCommand()
                to="xyz@gmail.com"
                sendEmail(to,content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("System can not able to send the message at this moment. please try later")

        
        elif 'exit jarvis' in query:
            break

    print("You exit jarvis Artificial Intelligence desktop")   
    speak("You exit jarvis Artificial Intelligence desktop")


