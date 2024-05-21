from win32com.client import Dispatch
import datetime
import speech_recognition as sr
import sys
import wikipedia
import webbrowser
import os
import random
import AppOpener

def speak(str):                             #Takes an input string and speaks it as an output
    speaker = Dispatch("Sapi.SpVoice")
    speaker.Speak(str)

def wish_me():                              #Wishes you Good Morning, Good Evening or Good Night based on time time
    hr = int(datetime.datetime.now().hour)  #Separates the hour value from the current time
    if hr>=0 and hr<12:
        speak("Good Morning!")
    elif hr>=12 and hr<=17:
        speak("Good Afternoon!")
    else:
        speak("Good Night!")
    speak("I am Sam. How may I help you?")

def take_command():                         #Takes an audio input from the user and returns its string output
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")          
        r.pause_threshold = 1
        audio = r.listen(source)            #The variable "audio" is now contained with the audio input from the user
    try:                                    #If the system recognized what the user said, then the audio input is converted to string
        print("Recognizing...")
        query = r.recognize_google(audio,language = "en-in")  #The audio stored in "query" has now been converted to string
        print(f"User said: {query}")
    except Exception as e:                  #The exception is raised if system could not catch what the user said
        print("Sorry, Couldn't catch that, Say that again please")
        return "None"
    return query

if __name__ == '__main__': 

    wish_me()

    while True:

        query = take_command().lower()
        if "wikipedia" in query:            #Reads you the information about something from wikipedia
            print("Searching Wikipedia...")
            query = query.replace("wikipedia","")       #Replaces the word "wikipedia" from query so that information of Wikipedia is not searched on Wikipedia website
            results = wikipedia.summary(query,sentences=2)  #Stores into the variable "results" the first two lines of the information about the query from Wikipedia
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif "youtube" in query:            #Opens Youtube
            webbrowser.open("youtube.com")

        elif "google" in query:
            webbrowser.open("google.com")   #Opens Google

        elif "play music" in query:         #Plays music in random order given that the mp3 files are present in the same folder
            list1 = os.listdir("C://Users/Prem/OneDrive/Desktop/Programming Projects/Python Programming Language/Sam Desktop Assistant") #Names of all the files in the given location are present in this list now
            list2 = [item for item in list1 if "." in item] #Names of only those files which contain "." in their name is present in this list now
            music_dir = [songs for songs in list2 if songs.split(".")[1]=="mp3"] #Names of only mp3 files is present in this list now
            music = random.choice(music_dir)      #Random mp3 file name is chosen from the list "music_dir"
            os.startfile(os.path.join("C://Users/Prem/OneDrive/Desktop/Programming Projects/Python Programming Language/Sam Desktop Assistant",music))  #os.path.join() intelligently joins the path components

        elif "time" in query:               #Tells the current time
            str_time = datetime.datetime.now().strftime("%H:%M:%S")  #Stores the current time in "str_time" in the form of a string
            speak(f"The current time is {str_time}")
            print(str_time)

        elif "whatsapp" in query:           #Opens WhatsApp
            AppOpener.open("WhatsApp")      #Opens the app in the argument

        elif "exit" in query:               #Exits the program
            speak("Hope you liked my service.")
            sys.exit("Exiting")

"""
This voice assistant can do the following on your commands:
1. Read you information about something from wikipedia
2. Open Youtube
3. Open Google
4. Play music
5. Tell time
6. Open WhatsApp
7. Exit itself 
"""