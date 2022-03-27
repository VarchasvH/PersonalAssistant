# Imports
import speech_recognition as sr    # Speech Recognition
import pyttsx3                     # text to speech
import datetime                    # To get the time
import wikipedia                   #To get information from wikipedia
import webbrowser                  #To open websites
import os                          #To open files
import time                        #To calculate time
import subprocess                  #To open files
from tkinter import *              #For the graphics
import pyjokes                     #For jokes
from playsound import playsound    #To play sounds
import keyboard              
import pyaudio

# Converting Text to Speech
assistantName = "Monarch"                                       # Name of the Assistant
engine = pyttsx3.init('sapi5')                                  # Defining the Engine && sapi5 is used in Windows
voices = engine.getProperty('voices')                           # TO get the voice
engine.setProperty('voice', voices[0].id)                       # 1 is for female voice and 0 is for Male voice

#Speak Function
def speak(text):
    engine.say(text)                              # Assistant Speaks
    print(assistantName + " : " + text)           # Speech is printed 
    engine.runAndWait(1)                          # Inside the brackets is the wait time in ms


# Greet Function
def greet():
    hour = datetime.datetime.now().hour   
    if(hour >= 0 and hour < 12):
        speak("Good Morning, Master!")
        
    elif ( hour >= 12 and hour < 18):
        speak("Good After, Master!")
        
    else:
        speak("Good Evening, Master!")


# Date Function    
def date():
    now = datetime.datetime.now() # Current time
    myDate = datetime.datetime.today() # Current Date
    
    monthName = now.month  # Current Month
    dayName = now.day    # Current day
    monthNames = ['Janameary', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
    dates = [ '1st', '2nd', '3rd', ' 4th', '5th', '6th', '7th', '8th', '9th', '10th', '11th', '12th', '13th', '14th', '15th', '16th', '17th', '18th', '19th', '20th', '21st', '22nd', '23rd','24rd', '25th', '26th', '27th', '28th', '29th', '30th', '31st']
        
    # Speaking the Date
    speak("Date is : " + monthNames[monthName - 1] + " " + dates[dayName -1] + '-')


# Note Function 
def note(text):
    date = datetime.datetime.now() # Time at which the note was made
    fileName = str(date).replace(":", "-") + "-note.txt"   # Creating the File
    with open(fileName, "w") as f:
        f.write(text)  # Writing the text in the notepad

    subprocess.Popen(["notepad.exe", fileName]) # Opening the text file

# WIkipedia Function
def wikipediaScreen(text):
    wikipediaScreen = Toplevel(screen)
    wikipediaScreen.title(text)
    wikipediaScreen.iconbitmap('appIcon.ico')
    message = Message(wikipediaScreen, text = text)
    message.pack()
    

# Getting user Input 
def getAudio():
    r = sr.Recognizer
    audio = ''
    
    with sr.Microphone() as source:
        print("Listening")
        playsound(("assistant_on.wav"))
        audio = r.listen(source, phrase_time_limit= 3)
        playsound("assistant_on.wav")
        print("Stop")
        
    try:
        text = r.recognize_google(audio, language = 'en-US')
        print('You : ' + ': ' + text)
        return text  
    
    except:
        return "None"

def ProcessAudio():

    run = 1
    if __name__=='__main__':
        while run==1:

            app_string = ["open word", "open powerpoint", "open excel", "open zoom","open notepad",  "open chrome"]
            app_link = [r'\Microsoft Office Word 2007.lnk',r'\Microsoft Office PowerPoint 2007.lnk', r'\Microsoft Office Excel 2007.lnk', r'\Zoom.lnk', r'\Notepad.lnk', r'\Google Chrome.lnk' ]
            app_dest = r'C:\Users\shriraksha\AppData\Roaming\Microsoft\Windows\Start Menu\Programs' # Write the location of your file

            statement = getAudio().lower()
            results = ''
            run +=1

            if "hello" in statement or "hi" in statement:
                greet()               


            if "good bye" in statement or "ok bye" in statement or "stop" in statement:
                speak('Your personal assistant ' + assistantName +' is shutting down, Good bye')
                screen.destroy()
                break

            if 'wikipedia' in statement:
                try:


                    speak('Searching Wikipedia...')
                    statement = statement.replace("wikipedia", "")
                    results = wikipedia.summary(statement, sentences = 3)
                    speak("According to Wikipedia")
                    wikipediaScreen(results)
                except:
                    speak("Error")


            if 'joke' in statement:
                speak(pyjokes.get_joke())    


            if 'open youtube' in statement:
                webbrowser.open_new_tab("https://www.youtube.com")
                speak("youtube is open now")
                time.sleep(5)


            if 'open google' in statement:
                    webbrowser.open_new_tab("https://www.google.com")
                    speak("Google chrome is open now")
                    time.sleep(5)


            if 'open gmail' in statement:
                    webbrowser.open_new_tab("mail.google.com")
                    speak("Google Mail open now")
                    time.sleep(5)

            if 'open netflix' in statement:
                    webbrowser.open_new_tab("netflix.com/browse") 
                    speak("Netflix open now")


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
                speak("Zoom is opening now")


            if app_string[4] in statement:
                os.startfile(app_dest + app_link[4])
                speak("Notepad is opening now")
        
            if app_string[5] in statement:
                os.startfile(app_dest + app_link[5])
                speak("Google chrome is opening now")


            if 'news' in statement:
                news = webbrowser.open_new_tab("https://timesofindia.indiatimes.com/city/mangalore")
                speak('Here are some headlines from the Times of India, Happy reading')
                time.sleep(6)

            if 'cricket' in statement:
                news = webbrowser.open_new_tab("cricbuzz.com")
                speak('This is live news from cricbuzz')
                time.sleep(6)

            if 'corona' in statement:
                news = webbrowser.open_new_tab("https://www.worldometers.info/coronavirus/")
                speak('Here are the latest covid-19 numbers')
                time.sleep(6)

            if 'time' in statement:
                strTime=datetime.datetime.now().strftime("%H:%M:%S")
                speak(f"the time is {strTime}")

            if 'date' in statement:
                date()

            if 'who are you' in statement or 'what can you do' in statement:
                    speak('I am '+assistantName+' your personal assistant. I am programmed to minor tasks like opening youtube, google chrome, gmail and search wikipedia etcetra') 


            if "who made you" in statement or "who created you" in statement or "who discovered you" in statement:
                speak("I was created with love by VarchasvH")

            
            if 'make a note' in statement:
                statement = statement.replace("make a note", "")
                note(statement)


            if 'note this' in statement:    
                statement = statement.replace("note this", "")
                note(statement)         

            speak(results)


# Name change function
def nameChange():
    nameInfo = name.get()
    file = open("assistantName", "w") #Opening the name file
    file.write(nameInfo)  # Writing the name
    # windows Destroy
    settingsScreenDestroy()
    screen.destroy()
    
# Name Change window Function
def windowNameChange():
    global settingsScreen
    global name
    
    settingsScreen = Toplevel(screen)
    settingsScreen.title("Settings") 
    settingsScreen.geometry("300x300")
    settingsScreen.iconbitmap('appIcon.ico')
    
    name = StringVar()
    
    currentLabel = Label(settingsScreen, text = "Current name: "+ assistantName)
    currentLabel.pack()
    
    enterLabel = Label(settingsScreen, text = "Enter the Name for the Assistant")
    enterLabel.pack(pady = 10)
    
    nameLabel = Label(settingsScreen, text = "Name")
    nameLabel.pack(pady = 10)
    
    nameEntry = Entry(settingsScreen, textvariable = name)
    nameEntry.pack()
    
    changeNameButton = Button(settingsScreen, text = "Ok", width = 10, height = 1, command = nameChange)
    changeNameButton.pack(pady = 10)
    

# Info about me
def info():
    infoScreen = Toplevel(screen)
    infoScreen.title("Info")
    infoScreen.iconbitmap('appIcon.ico')

    createrLabel = Label(infoScreen, text = "Created by Varchasv Hoon")
    createrLabel.pack()

    ageLabel = Label(infoScreen, text = " at the age of 19")
    ageLabel.pack()
    
    forLabel = Label(infoScreen, text = "For Github Repository")
    forLabel.pack()

# GUI
def mainScreen():
    global screen
    screen = Tk()
    screen.title(assistantName) # Title of the Screen
    screen.geometry("250x100")
    screen.iconbitmap('appIcon.ico') #Icon of the Screen
    
    nameLabel = Label(text = assistantName, width = 300, bg = "black", fg = "white", font = ("Calibri", 13))

    # Designing the Microphone Button
    microphonePhoto = PhotoImage(file = "assistantLogo.png")
    microphoneButton = Button(image = microphonePhoto, command = ProcessAudio)
    microphoneButton.pack(pady = 10)
    
    # Designing the Settings button
    settingsPhoto = PhotoImage(file = "settings.png")
    settingsButton = Button(image = settingsPhoto, command = windowNameChange)
    settingsButton.pack(pady=10)

    info_button = Button(text ="Info", command = info)
    info_button.pack(pady=10)

    screen.mainloop()
    
    
mainScreen()

    
