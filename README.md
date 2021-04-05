# <h1>Artificial-Intelligence-Desktop-Voice-Assistant</h2>


Have you ever wondered how cool it would be to have your own A.I. assistant? Imagine how easier it would be to send emails without typing a single word, doing Wikipedia searches without opening web browsers, and performing many other daily tasks like playing music with the help of a single voice command.

### What can this A.I. assistant do for you?

* It can send emails on your behalf.

* It can play music for you.

* It can do Wikipedia searches for you.

* It is capable of opening websites like Google, Youtube, etc., in a web browser.

* It is capable of opening your code editor or IDE with a single voice command.

### Defining Speak Function
The first and foremost thing for an A.I. assistant is that it should be able to speak. To make our J.A.R.V.I.S. talk, we will make a function called <i><b>speak()</b></i>. This function will take audio as an argument, and then it will pronounce it.

```python
def speak(audio):
       pass      #For now, we will write the conditions later.
```
Now, the next thing we need is audio. We must supply audio so that we can pronounce it using the speak() function we made. We are going to install a module called pyttsx3.

### What is pyttsx3?
* A python library that will help us to convert text to speech. In short, it is a text-to-speech library.
* It works offline, and it is compatible with Python 2 as well as Python 3.

### Installation:
```
pip install pyttsx3
```
In case you receive such errors:
* No module named win32com.client
* No module named win32
* No module named win32api

Then, install pypiwin32 by typing the below command in the terminal :
```
pip install pypiwin32.
```
After successfully installing pyttsx3, import this module into your program

### Usage:
```python
import pyttsx3

engine = pyttsx3.init('sapi5')

voices= engine.getProperty('voices') #getting details of current voice

engine.setProperty('voice', voice[0].id)
```
### What is sapi5?
* Microsoft developed speech API.
* Helps in synthesis and recognition of voice.

### What Is VoiceId?
* Voice id helps us to select different voices.
* voice[0].id = Male voice 
* voice[1].id = Female voice

### Writing Our speak() Function :
We made a function called <b>speak()</b> at the starting of this tutorial. Now, we will write our <b>speak()</b> function to convert our text to speech.
```python
def speak(audio):

engine.say(audio) 

engine.runAndWait() #Without this command, speech will not be audible to us.
```
### Creating Our main() function: 
We will create a main() function, and inside this main() Function, we will call our speak function.
#### Code:
```python
if __name__=="__main__" :
     speak("Hello Sir")
```
Whatever you will write inside this speak() function will be converted into speech. Congratulations! With this, our J.A.R.V.I.S. has its own voice, and it is ready to speak.

## Defining Wish me Function :
Now, we will make a <b>wishme()</b> function, that will make our J.A.R.V.I.S. wish or greet the user according to the time of computer or pc. To provide current or live time to A.I., we need to import a module called ```datetime```. Import this module to your program, by:
```python
import datetime
```
Now, let's start defining the <b>wishme()</b> function:
```python
def wishme():
hour = int(datetime.datetime.now().hour)
```
Here, we have stored the current hour or time integer value into a variable named hour. Now, we will use this hour value inside an if-else loop.

## Defining Take command Function :
The next most important thing for our A.I. assistant is that it should take command with the help of the microphone of the user's system. So, now we will make a <b>takeCommand()</b> function.  With the help of the <b>takeCommand()</b> function, our A.I. assistant will return a string output by taking microphone input from the user.

Before defining the <b>takeCommand()</b> function, we need to install a module called <b>speechRecognition</b>. Install this module by: 
```
pip install speechRecognition
```
After successfully installing this module, import this module into the program by writing an import statement.
```python
import speechRecognition as sr
```
Let's start coding the takeCommand() function :
```python
def takeCommand():
    #It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
 ```
 We have successfully created our takeCommand() function. Now we are going to add a try and except block to our program to handle errors effectively.
 ```python
   try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in') #Using google for voice recognition.
        print(f"User said: {query}\n")  #User query will be printed.

    except Exception as e:
        # print(e)    
        print("Say that again please...")   #Say that again will be printed in case of improper voice 
        return "None" #None string will be returned
    return query
```
## Coding logic of Jarvis
Now, we will develop logic for different commands such as Wikipedia searches, playing music, etc.
### Defining Task 1: To search something on Wikipedia
To do Wikipedia searches, we need to install and import the Wikipedia module into our program. Type the below command to install the Wikipedia module :
```
pip install wikipedia
```
After successfully installing the Wikipedia module, import it into the program by writing an import statement.
```python
if __name__ == "__main__":
    wishMe()
    while True:
    # if 1:
        query = takeCommand().lower() #Converting user query into lower case

        # Logic for executing tasks based on query
        if 'wikipedia' in query:  #if wikipedia found in the query then this block will be executed
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2) 
            speak("According to Wikipedia")
            print(results)
            speak(results)
```
In the above code, we have used an if statement to check whether Wikipedia is in the search query of the user or not. If Wikipedia is found in the user's search query, then two sentences from the summary of the Wikipedia page will be converted to speech with the speak function's help.
### Defining Task 2: To open YouTube site in a Web-browser
To open any website, we need to import a module called webbrowser. It is an in-built module, and we do not need to install it with a pip statement; we can directly import it into our program by writing an import statement.

#### Code: 
```python
     elif 'open youtube' in query:
            webbrowser.open("youtube.com")
 ```
Here, we are using the elif loop to check whether Youtube is in the user's query. Let' suppose the user gives a command as "J.A.R.V.I.S., open youtube." So, open youtube will be in the user's query, and the elif condition will be true.

### Defining Task 3: To open Google site in a web-browser
```python
     elif 'open google' in query:
            webbrowser.open("google.com")
``` 

We are opening Google in a web-browser by applying the same logic that we used to open youtube. 
### Defining Task 4: To play music
To play music, we need to import a module called os. Import this module directly with an import statement.
```python
elif 'play music' in query:
            music_dir = 'D:\\Non Critical\\songs\\Favorite Songs2'
            songs = os.listdir(music_dir)
            print(songs)    
            os.startfile(os.path.join(music_dir, songs[0]))
```
In the above code, we first opened our music directory and then listed all the songs present in the directory with the os module's help. With the help of ```os.startfile```, you can play any song of your choice. I am playing the first song in the directory. However, you can also play a random song with the help of a random module. Every time you command to play music, J.A.R.V.I.S. will play any random song from the song directory.
### Defining Task 5: To know the current time
```python
  elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f"Sir, the time is {strTime}")
```
In the above, code with using ```datetime()``` function and storing the current or live system into a variable called ```strTime```. After storing the time in ```strTime```, we are passing this variable as an argument in speak function. Now, the time string will be converted into speech.

### Defining Task 6: To open the VS Code Program
```python
 elif 'open code' in query:
            codePath = "C:\\Users\\HP\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)
 ```
 To open the VS Code or any other application, we need the code path of the application.

<b>Steps to get the code path of the application:</b>

<b>Step 1</b>: Open the file location.

<b>Step 2</b>: Right-click on the application and click on properties.

<b>Step 3</b>: Copy the target from the target section.

After copying the target of the application, save the target into a variable. Here, I am saving the target into a variable called codePath, and then we are using the os module to open the application.

### Defining Task 7: To send Email

To send an email, we need to import a module called smtplib.

#### What is smtplib?
* Simple Mail Transfer Protocol (SMTP) is a protocol that allows us to send emails and to route emails between mail servers. An instance method called <b>sendmail</b> is present in the SMTP module. This instance method allows us to send an email. 
##### It takes 3 parameters:
* <b>The sender</b>: Email address of the sender.
* <b>The receiver</b>: Email of the receiver.
* <b>The message</b>: A string message which needs to be sent to one or more than one recipient.

## Defining Send email function :
We will create a <b>sendEmail()</b> function, which will help us send emails to one or more than one recipient.
```python
def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('youremail@gmail.com', 'your-password')
    server.sendmail('youremail@gmail.com', to, content)
    server.close()
```
In the above code, we are using the SMTP module, which we have already discussed above.

* <b>Note</b>: Do not forget to 'enable the less secure apps' feature in your Gmail account. Otherwise, the sendEmail function will not work properly.
## Calling sendEmail() function inside the main() function: 
``` python
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
```
We are using the try and except block to handle any possible error while sending emails.
## Recapitulate
1. First of all, we have created a wishme() function that gives the greeting functionality according to our A.I system time.
2. After wishme() function, we have created a takeCommand() function, which helps our A.I to take command from the user. This function is also responsible for returning the user's query in a string format.
3. We developed the code logic for opening different websites like google, youtube, and stack overflow.
4. Developed code logic for opening VS Code or any other application.
5. At last, we added functionality to send emails.

## Is it an A.I.?
Many people will argue that the virtual assistant that we have created is not an A.I, but it is the output of a bunch of the statement. But, if we look at the fundamental level, the sole purpose of A.I develop machines that can perform human tasks with the same effectiveness or even more effectively than humans.

It is a fact that our virtual assistant is not a very good example of A.I., but it is an A.I. !
