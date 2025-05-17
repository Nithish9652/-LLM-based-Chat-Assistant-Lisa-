import speech_recognition as sr
import win32com.client
import webbrowser
import time
import os
import datetime
import urllib.parse
import requests
import uuid
import re
import screen_brightness_control as sbc
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
apikey = "AIzaSyAe2L5HOOK8Spda8lt9bWl8-CCZshknN7Y"

# Setup voice
speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Voice = speaker.GetVoices().Item(1)

# Recognizer configuration
r = sr.Recognizer()
r.energy_threshold = 1000
r.dynamic_energy_threshold = True
r.pause_threshold = 0.5

def get_brightness():
    current_brightness = sbc.get_brightness()[0]
    return current_brightness

# Function to increase brightness by 10%
def increase_brightness():
    current_brightness = get_brightness()
    new_brightness = min(current_brightness + 10, 100)
    sbc.set_brightness(new_brightness)
    speaker.Speak(f"Brightness increased to {new_brightness}%")

# Function to decrease brightness by 10%
def decrease_brightness():
    current_brightness = get_brightness()
    new_brightness = max(current_brightness - 10, 0)
    sbc.set_brightness(new_brightness)
    speaker.Speak(f"Brightness decreased to {new_brightness}%")


# Function to get the current volume
def get_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, 0, None)
    volume_interface = cast(interface, POINTER(IAudioEndpointVolume))
    volume = volume_interface.GetMasterVolumeLevelScalar() * 100
    return round(volume)


# Function to increase the volume by 10%
def increase_volume():
    current_volume = get_volume()
    new_volume = min(current_volume + 10, 100)
    set_volume(new_volume)
    speaker.Speak(f"Volume increased to {new_volume}%")


# Function to decrease the volume by 10%
def decrease_volume():
    current_volume = get_volume()
    new_volume = max(current_volume - 10, 0)
    set_volume(new_volume)
    speaker.Speak(f"Volume decreased to {new_volume}%")


# Function to set the volume to a specific value (0 to 100)
def set_volume(volume_level):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, 0, None)
    volume_interface = cast(interface, POINTER(IAudioEndpointVolume))
    volume_interface.SetMasterVolumeLevelScalar(volume_level / 100, None)

def ai(prompt):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={apikey}"

    headers = {
        "Content-Type": "application/json"
    }

    data = {
        "contents": [
            {
                "parts": [
                    {"text": prompt}
                ]
            }
        ]
    }

    try:
        response = requests.post(url, headers=headers, json=data)
        response.raise_for_status()
        result = response.json()

        # Extract the response text
        output = result["candidates"][0]["content"]["parts"][0]["text"]
        full_text = f"Gemini API response for Prompt: {prompt}\n*************************\n\n{output}"
        print(full_text)

        # Save response to file
        output_dir = "GeminiResponses"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        safe_prompt = re.sub(r'[<>:"/\\|?*]', '_', prompt[:50])
        filename = f"{safe_prompt}_{uuid.uuid4().hex[:8]}.txt"
        file_path = os.path.join(output_dir, filename)

        with open(file_path, "w", encoding="utf-8") as f:
            f.write(full_text)

        print(f"Saved response to {file_path}")
        speaker.Speak("I have fetched and saved the response.")

    except Exception as e:
        print(f"Error fetching Gemini response: {e}")
        speaker.Speak("There was a problem fetching the AI response.")


def greet_user():
    current_hour = time.localtime().tm_hour
    tf = "morning" if 5 <= current_hour < 12 else "afternoon" if 12 <= current_hour <= 18 else "evening"
    speaker.Speak(f"Good {tf}, this is Lisa. How can I help you?")

def handle_youtube():
    speaker.Speak("Opening YouTube. Do you want to play any Video?")
    while True:
        with sr.Microphone() as source2:
            print("Listening for Video request...")
            r.adjust_for_ambient_noise(source2, duration=0.5)
            try:
                audio2 = r.listen(source2, timeout=5, phrase_time_limit=5)
                response = r.recognize_google(audio2, language="en-in").lower()
                print(f"User said: {response}")

                if response == "no thanks":
                    webbrowser.open("https://youtube.com")
                    speaker.Speak("Opening YouTube.")
                    break
                elif response in ["exit", "quit"]:
                    speaker.Speak("Okay, canceling YouTube.")
                    break
                else:
                    from pytube import Search
                    s = Search(response)
                    if s.results:
                        first_video = s.results[0]
                        webbrowser.open(first_video.watch_url)
                        speaker.Speak(f"Playing {response} on YouTube")
                    else:
                        speaker.Speak("Sorry, couldn't find the video")

                    time.sleep(1)
                    speaker.Speak("Would you like to play another Video?")
                    try:
                        audio3 = r.listen(source2, timeout=5, phrase_time_limit=5)
                        another_response = r.recognize_google(audio3, language="en-in").lower()
                        if "no thanks" in another_response:
                            speaker.Speak("Okay, enjoy your Video!")
                            break
                        elif "yes" in another_response:
                            speaker.Speak("What Video would you like to play?")
                            continue
                    except:
                        speaker.Speak("I didn't understand. Exiting YouTube mode.")
                        break
            except sr.WaitTimeoutError:
                speaker.Speak("No input received. Please try again.")
                continue
            except sr.UnknownValueError:
                speaker.Speak("I didn't understand. Could you repeat?")
                continue
            except sr.RequestError:
                speaker.Speak("I am unable to process your request. Check your internet connection.")
                break

def sendmail():
    while True:
        try:
            speaker.Speak("Please say the username: ")
            with sr.Microphone() as source:
                audio = r.listen(source, timeout=5, phrase_time_limit=5)
                user_part = r.recognize_google(audio, language="en-in").lower().replace(" ", "")
                print(f"Username part: {user_part}")
                break
        except Exception as e:
            print(f"Error getting username: {e}")
            speaker.Speak("Sorry, I didn't catch that.")
    while True:
        try:
            speaker.Speak("Please say the domain of the recipient email. For example, gmail.com or iiitl.ac.in")
            with sr.Microphone() as source:
                audio = r.listen(source, timeout=5, phrase_time_limit=5)
                domain = r.recognize_google(audio, language="en-in").lower().replace(" ", "")
                print(f"Domain: {domain}")
                break
        except Exception as e:
            print(f"Error getting domain: {e}")
            speaker.Speak("Sorry, I didn't catch that. Please say the domain again.")

    recipient = f"{user_part}@{domain}"

    while True:
        try:
            speaker.Speak("What should be the subject?")
            with sr.Microphone() as source:
                audio = r.listen(source, timeout=5, phrase_time_limit=5)
                subject = r.recognize_google(audio, language="en-in")
                print(f"Subject: {subject}")
                break
        except Exception as e:
            print(f"Error getting subject: {e}")
            speaker.Speak("Sorry, I didn't catch that. Please say the subject again.")

    while True:
        try:
            speaker.Speak("What should I write in the email?")
            with sr.Microphone() as source:
                audio = r.listen(source, timeout=5, phrase_time_limit=5)
                content = r.recognize_google(audio, language="en-in")
                print(f"Content: {content}")
                break
        except Exception as e:
            print(f"Error getting content: {e}")
            speaker.Speak("Sorry, I didn't catch that. Please say the message again.")

    try:
        email_url = f"mailto:{recipient}?subject={urllib.parse.quote(subject)}&body={urllib.parse.quote(content)}"
        webbrowser.open(email_url)
        speaker.Speak("I've opened the email draft for you. Please review and send it.")

    except Exception as e:
        print(f"Error opening email draft: {e}")
        speaker.Speak("Something went wrong while opening the email draft.")


def open_website(query):
    websites = {
        "open google": "https://www.google.com",
        "open facebook": "https://www.facebook.com",
        "open twitter": "https://www.twitter.com",
        "open linkedin": "https://www.linkedin.com",
        "open instagram": "https://www.instagram.com",
        "open reddit": "https://www.reddit.com",
        "open amazon": "https://www.amazon.com",
        "open ebay": "https://www.ebay.com",
        "open wikipedia": "https://www.wikipedia.org",
        "open yahoo": "https://www.yahoo.com",
        "open bing": "https://www.bing.com",
        "open netflix": "https://www.netflix.com",
        "open spotify": "https://www.spotify.com",
        "open github": "https://www.github.com",
        "open stack overflow": "https://stackoverflow.com",
        "open medium": "https://www.medium.com",
        "open quora": "https://www.quora.com",
        "open pinterest": "https://www.pinterest.com",
        "open tiktok": "https://www.tiktok.com",
        "open tumblr": "https://www.tumblr.com",
        "open microsoft": "https://www.microsoft.com",
        "open apple": "https://www.apple.com",
        "open adobe": "https://www.adobe.com",
        "open coursera": "https://www.coursera.org",
        "open udemy": "https://www.udemy.com",
        "open edx": "https://www.edx.org",
        "open cnn": "https://www.cnn.com",
        "open bbc": "https://www.bbc.com",
        "open nytimes": "https://www.nytimes.com",
        "open forbes": "https://www.forbes.com"
    }

    applications = {
        "open notepad": "notepad.exe",
        "open calculator": "calc.exe",
        "open paint": "mspaint.exe",
        "open word": "winword.exe",
        "open excel": "excel.exe",
        "open powerpoint": "powerpnt.exe",
        "open chrome": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        "open firefox": r"C:\Program Files\Mozilla Firefox\firefox.exe",
        "open control panel": "control.exe",
        "open task manager": "taskmgr.exe",
        "open cmd": "cmd.exe",
        "open camera": "microsoft.windows.camera:",
        "open settings": "ms-settings:",
        "open clock": "ms-clock:",
        "open calendar": "outlookcal:",
        "open mail": "outlook.exe",
        "open photos": "ms-photos:",
        "open spotify": r"C:\Users\{username}\AppData\Roaming\Spotify\Spotify.exe",
        "open vlc": r"C:\Program Files\VideoLAN\VLC\vlc.exe",
        "open steam": r"C:\Program Files (x86)\Steam\Steam.exe",
        "open vs code": r"C:\Users\{username}\AppData\Local\Programs\Microsoft VS Code\Code.exe"
    }

    for key, url in websites.items():
        if query == key:
            webbrowser.open(url)
            speaker.Speak(f"Opening {key.replace('open ', '')}.")
            return True

    try:
        for app_name, app_path in applications.items():
            if query == app_name:
                os.startfile(app_path)
                speaker.Speak(f"Opening {app_name.replace('open ', '')}")
                return True
    except Exception:
        speaker.Speak(f"Sorry, I couldn't open {query.replace('open ', '')}. The application might not be installed.")
        return False

    return False

def main():
    greet_user()
    while True:
        with sr.Microphone() as source:
            print("Listening...")
            r.adjust_for_ambient_noise(source, duration=0.5)
            try:
                audio = r.listen(source, timeout=5, phrase_time_limit=5)
                query = r.recognize_google(audio, language="en-in").lower()
                print(f"User said: {query}")

                if query == "exit":
                    speaker.Speak("Thank you, have a nice day!")
                    break
                if "increase brightness" in query:
                    increase_brightness()
                if "decrease brightness" in query:
                    decrease_brightness()

                if "send email" in query or "send mail" in query:
                    sendmail()
                    continue

                if "the time" in query:
                    now = datetime.datetime.now()
                    speaker.Speak(f"The time is {now.strftime('%H')} hours and {now.strftime('%M')} minutes")
                    continue

                if "increase volume" in query:
                    increase_volume()
                    continue

                if "decrease volume" in query:
                    decrease_volume()
                    continue
                if query == "open youtube":
                    handle_youtube()
                    continue
                if "using artificial intelligence" in query:
                    ai(prompt=query)
                if not open_website(query):
                    speaker.Speak(query)

            except sr.WaitTimeoutError:
                continue
            except sr.UnknownValueError:
                speaker.Speak("I couldn't understand. Please say that again.")
                continue
            except sr.RequestError:
                speaker.Speak("There was a problem connecting to the internet. Please check your connection.")
                continue

# Start the program
if __name__ == "__main__":
    main()
