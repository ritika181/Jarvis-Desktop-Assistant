import pyttsx3
import speech_recognition as sr
import datetime
import os
import subprocess
import threading
import pywhatkit as kit
import webbrowser
import wikipedia
from requests import get
import cv2
from groq import Groq
import smtplib
import pyjokes
import requests
import feedparser
import instadownloader
import PyPDF2
import psutil
import time
import pyaudio
import sys
import pythoncom
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import QTimer, QTime, QDate, Qt
from PyQt5.QtGui import QMovie
from PyQt5.QtCore import*
from PyQt5.QtGui import*
from PyQt5.QtWidgets import*
from PyQt5.uic import loadUiType
from frontened_jarvis import Ui_jarvisUi
import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def speak(text):
    print("Jarvis:", text)
    speaker.Speak(text)

# def speak(text):
#     print(text)
#     speaker.speak_signal.emit(text)

def takecommand():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("listening..")
        try:
            r.pause_threshold=1
            audio=r.listen(source,timeout=3,phrase_time_limit=5)
        except sr.WaitTimeoutError:
            return "none"
    try:
        print("recognizing.....")
        query=r.recognize_google(audio,language='en-in')
        print(f"user said : {query}")
        
    except Exception as e:
        speak("pardon!can you repeat please")
        return "none"
    return query
#-----------------------------wish-----------------------
def wish():
    hour=int(datetime.datetime.now().hour)
    if hour < 12:
        speak("Good morning , hey there im jarvis , how may i help you")
    elif hour < 18:
        speak("Good afternoon , hey there im jarvis , how may i help you")
    else:
        speak("Good evening , hey there im jarvis , how may i help you")

def open_adobe():
    os.startfile(r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe")

#-----------------------WEATHER-----------------------------
def get_weather(city):
    API_KEY = os.getenv("OPENWEATHER_API_KEY")
    url = "https://api.openweathermap.org/data/2.5/weather"
    params = {"q": city, "appid": API_KEY, "units": "metric"}

    try:
        data = requests.get(url, params=params).json()
        if data.get("cod") != 200:
            speak("City not found")
            return

        temp = data["main"]["temp"]
        desc = data["weather"][0]["description"]
        speak(f"It is {temp} degrees in {city} with {desc}")
    except:
        speak("Unable to fetch weather")

#---------------------------------RAIN ALERT----------------------------------
def rain_alert(city):
    api_key = os.getenv("OPENWEATHER_API_KEY")
    url = f"https://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"
    data = requests.get(url).json()

    if data["cod"] != 200:
        speak(f"Sorry, I couldn't find the weather for {city}.")
        return

    weather_main = data["weather"][0]["main"].lower()
    description = data["weather"][0]["description"]

    if "rain" in weather_main or "drizzle" in weather_main or "thunderstorm" in weather_main:
        speak(f"Rain alert! It is {description} in {city}. Please carry an umbrella.")
    else:
        speak(f"No rain expected in {city}. Weather is {description}.")

#-------------------------ip----------------------
def get_city_by_ip():
    try:
        response = requests.get('https://ipinfo.io')
        data = response.json()
        city = data.get('city', None)
        return city
    except Exception as e:
        speak("Sorry, I couldn't detect your location.")
        return None

def rain_alert_by_ip():
    city = get_city_by_ip()
    if not city:
        return

    api_key = os.getenv("OPENWEATHER_API_KEY")
    url = f"https://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"
    data = requests.get(url).json()

    if data["cod"] != 200:
        speak("I couldn't find the weather information for your location.")
        return

    weather_main = data["weather"][0]["main"].lower()
    description = data["weather"][0]["description"]

    if "rain" in weather_main or "drizzle" in weather_main or "thunderstorm" in weather_main:
        speak(f"Rain alert! It is {description} in {city}. Please carry an umbrella.")
    else:
        speak(f"No rain expected in {city}. Weather is {description}.")

#----------------------------------GROQ AI-------------------------------------
client = Groq(api_key=os.getenv("GROQ_API_KEY"))

SYSTEM_PROMPT = """
You are Jarvis, an intelligent, friendly, and human-like AI assistant.
Rules:
- Speak naturally like a human, not like a robot.
- If the question is vague, politely ask a follow-up question.
Personality:
- Polite
- Friendly
- Supportive
- Conversational
"""
def ai_response(prompt):
    try:
        completion = client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[
                {"role": "system", "content":SYSTEM_PROMPT },
                {"role": "user", "content": prompt}
            ],
            temperature = 0.7,
            max_tokens = 500
            )
        return completion.choices[0].message.content
    except Exception as e:
        print("Error in AI call:", e)
        return "Sorry, I couldn't process that."
#------------------mail--------------------------    
def sendmail(to,content):
    sender = os.getenv("JARVIS_EMAIL")
    password = os.getenv("JARVIS_PASS")

    if not sender or not password:
        speak("Email credentials not set!")
        return
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender, password)
        server.sendmail(sender, to, content)
        server.quit()
        speak("Mail sent successfully!")
    except:
        speak("Sorry, unable to send email")
#-------------------news-------------------------
def tell_news():
    rss_url = "http://feeds.bbci.co.uk/news/rss.xml"  # BBC News RSS
    feed = feedparser.parse(rss_url)
    
    # Get top 10 news entries
    entries = feed.entries[:10]
    
    # Day list for Jarvis
    day = ["first","second","third","fourth","fifth","sixth","seventh","eighth","ninth","tenth"]

    if not entries:
        speak("Sorry, I couldn't fetch the news right now.")
        return

    for i in range(len(entries)):
        # speak each headline
        speak(f"Today's {day[i]} news is: {entries[i].title}")
#-----------------------pdfread------------------------
def pdf_reader(start_page=1,end_page=2):
    filename = "C:\\Users\\Dell\\Downloads\\Class Notes.pdf"    
    if not os.path.exists(filename):
        speak("File not found")
        return

    book = open(filename, "rb")
    pdfReader = PyPDF2.PdfReader(book)
    total_pages = len(pdfReader.pages)
    speak(f"Total number of pages in this book: {total_pages}")

    if start_page < 1 or end_page > total_pages or start_page > end_page:
        speak("Invalid page range")
        book.close()
        return

    for i in range(start_page - 1, end_page):
        page = pdfReader.pages[i]
        text = page.extract_text()
        if text:
            speak(text)
        else:
            speak(f"Page {i+1} has no readable text")

    book.close()

def wake_up():
    while True:
        command = takecommand().lower()
        if "jarvis" in command:
            speak("Yes sir?")
            break

#threading
class MainThread(QThread):
    def __init__(self):
        super(MainThread, self).__init__()

    def run(self):
        pythoncom.CoInitialize()
        self.TaskExecution()

    def TaskExecution(self):
        wake_up()
        time.sleep(0.5)
        wish()

        while True:
            self.query = takecommand().lower()
            if self.query == "":
                continue

            print("QUERY =", self.query)

            if "open adobe Reader" in self.query: 
                speak("Opening Adobe Reader") 
                open_adobe()

            elif "open youtube" in self.query: 
                speak("What should I play?") 
                song = takecommand() 
                speak(f"playing {song} on youtube") 
                kit.playonyt(song) 

            elif "open google" in self.query: 
                speak("What should I search?") 
                search = takecommand() 
                webbrowser.open(f"https://www.google.com/search?q={search}") 

            elif "open whatsapp" in self.query: 
                webbrowser.open("https://web.whatsapp.com")

            elif "open command prompt" in self.query: 
                os.system("start cmd") 

            elif "time" in self.query: 
                speak("The time is " + datetime.datetime.now().strftime("%H:%M:%S"))

            elif "wikipedia" in self.query: 
                speak("Searching Wikipedia") 
                speak("what should i search on wikipedia") 
                search=takecommand() 
                try: 
                    result = wikipedia.summary(search, sentences=2) 
                    speak(result) 
                except: speak("No results found")

            elif "news" in self.query or "tell me the news" in self.query:
                speak("Fetching the latest news for you")
                threading.Thread(target=tell_news).start()


            elif "ip address" in self.query: 
                ip = get("https://api.ipify.org").text 
                speak(f"Your IP address is {ip}")

            elif "open camera" in self.query:
                cap = cv2.VideoCapture(0)
                while True:
                    ret, frame = cap.read()
                    cv2.imshow("Camera", frame)
                    if cv2.waitKey(10) == 27:
                        break
                cap.release()
                cv2.destroyAllWindows()   
                    
            elif "play music" in self.query:
                music_dir = "C:\\Users\\Dell\\Music"
                songs = os.listdir(music_dir)
                os.startfile(os.path.join(music_dir, songs[0]))

            elif "weather" in self.query:
                speak("Tell me the city name")
                city = takecommand()
                if city:
                    get_weather(city)
         
            elif "rain alert" in self.query:
                speak("Do you want me to detect your city automatically? Say yes or no.")
                answer = takecommand().lower()
                if "yes" in answer:
                    rain_alert_by_ip()
                else:
                    speak("Please tell me your city name")
                    city = takecommand()
                    if city != "" and city != "none":
                        rain_alert(city)

            elif "send message" in self.query:
                message = takecommand().lower()
                speak("At what hour should I send it? (0-23)")
                try:
                    hour = int(takecommand())
                except:
                    speak("Invalid hour, using current hour")
                    from datetime import datetime
                    hour = datetime.now().hour
                    speak("At what minute should I send it? (0-59)")
                    try:
                        minute = int(takecommand())
                    except:
                        speak("Invalid minute, using current minute")
                        from datetime import datetime
                        minute = datetime.now().minute
                        phone = os.getenv("JARVIS_PHONE")
                        if not phone:
                            speak("Phone number is not set!")
                        else:
                            speak(f"Sending message to {phone} at {hour}:{minute}")
                            kit.sendwhatmsg(phone, message, hour, minute)

            elif "ask" in self.query or "question" in self.query:
                speak("Ask me anything!")
                question = takecommand() 
                if question:
                    answer = ai_response(question)
                    print(answer)
                    speak(answer)

            elif "email to ritika" in self.query:
                speak("What should I say?")
                content = takecommand().lower()
                to = os.getenv("JARVIS_RECIPIENT")  # get recipient from environment
                if not to:
                    speak("Recipient email not set!")
                else:
                    sendmail(to, content)

#to close notepad 
            elif "close notepad" in self.query:
                speak("Closing notepad")
                subprocess.call(["taskkill", "/f", "/im", "notepad.exe"])

            elif "notepad" in self.query:
                subprocess.Popen(["notepad.exe"])

            elif "alarm" in self.query:
                while True:
                    now = datetime.datetime.now()
                    if now.hour == 18 and now.minute == 37:
                        music_dir = "C:\\Users\\Dell\\Music"
                        songs = os.listdir(music_dir)
                        os.startfile(os.path.join(music_dir, songs[0]))
                        break

            elif "jokes" in self.query:
                joke=pyjokes.get_joke()
                speak(joke)

            elif "give my location" in self.query:
                speak("wait sir, let me check")
                try:
                    ipAdd = requests.get("https://api.ipify.org/").text
                    print(ipAdd)
                    url = f"https://get.geojs.io/v1/ip/geo/{ipAdd}.json"
                    geo_requests = requests.get(url)
                    geo_data = geo_requests.json()
                    city = geo_data['city']
                    country = geo_data['country']
                    speak(f"Sir, I am not sure, but I think we are in {city} city of {country}")
                except Exception as e:
                    speak("sorry sir, due to some issue I am not able to find where we are")
                    print(e)

            elif "shutdown the system" in self.query:
                os.system("shutdown /s /t 5")

            elif "restart the system" in self.query:
                os.system("shutdown /r /t 5")
        
            elif "instagram profile" in self.query:
                speak("sir,please enter the user name correctly")
                name=input("enter the username")
                webbrowser.open(f"www.instagram.com/{name}")
                speak(f"sir here is the profile of the user {name}")
                condition=takecommand()
                if "yes" in condition:
                    mod=instaloader.Instaloader()
                    mod.download_profile(name,profile_pic_only=True)
                    speak("profile picissaved on your folder")
                else:
                    pass
            
            elif "read pdf" in self.query:
                pdf_reader()

            elif "sleep the system" in self.query:
                speak("Putting the system to sleep")
                os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")

            elif "how much power does the system have" in self.query:
                battery=psutil.sensors_battery()
                percentage=battery.percent
                speak(f"sir,our system have {percentage} percent battery")
                if percentage>=75:
                    speak("we have enough power to continue our work ")
                elif percentage>=40 or percentage<=75:
                    speak("we should connect our system to charging point shortly")
                elif percentage>=15 or percentage<=40:
                    speak("we don't haveenough power to work")
                elif percentage<=15:
                    speak("very low battery , please connect to the charging point else the system will shut down shortly")

            if "no thanks" in self.query:
                speak("thanks for using me sir,Goodbye")
                break


startExecution = MainThread()

class Main(QDialog):
    def __init__(self):
        super().__init__()
        self.ui = Ui_jarvisUi()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.startTask)
        self.ui.pushButton_2.clicked.connect(self.close)

    def startTask(self):
        startExecution.start()

        self.ui.movie = QtGui.QMovie("../../../Downloads/iron_jarvis.gif")
        self.ui.label.setMovie(self.ui.movie)
        self.ui.movie.start()

        self.ui.movie = QtGui.QMovie("../../../Downloads/logo_iron.gif")
        self.ui.label_2.setMovie(self.ui.movie)
        self.ui.movie.start()

        self.ui.movie = QtGui.QMovie("../../../Downloads/initiating system.gif")
        self.ui.label_4.setMovie(self.ui.movie)
        self.ui.movie.start()

        timer = QTimer(self)
        timer.timeout.connect(self.showTime)
        timer.start(1000)
    
    def showTime(self):
        current_Time = QTime.currentTime()
        current_date = QDate.currentDate()
        label_Time = current_Time.toString('hh:mm:ss')
        label_date = current_date.toString(Qt.ISODate)
        self.ui.textBrowser.setText(label_date)
        self.ui.textBrowser_2.setText(label_Time)

app = QApplication(sys.argv)
jarvis = Main()
jarvis.show()
sys.exit(app.exec_())