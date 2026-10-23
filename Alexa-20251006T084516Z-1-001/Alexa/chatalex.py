# ======================================
# 🧠 ALEXA PERSONAL DESKTOP ASSISTANT
# ======================================
# Features:
# - Voice recognition (press Enter to interrupt)
# - WikipediaAPI (clean summaries)
# - Daily quote (same one per day)
# - Requires saying "Alexa" before commands
# - Route finder via Google Maps
# - Talking responses (Windows voice)
# - Plays actual YouTube videos directly (pywhatkit)

import datetime
import time
import webbrowser
import threading
import requests
import speech_recognition as sr
import wikipediaapi
import win32com.client  # Windows COM (SAPI voice)
import os
import json
import pywhatkit  # ✅ your play feature

# ======================================
# VOICE SETUP
# ======================================
speaker = win32com.client.Dispatch("SAPI.SpVoice")  # Text-to-speech voice

def speak(text):
    """Make Alexa speak and print text."""
    print(f"Alexa 🎤: {text}")
    speaker.Speak(text)


# ======================================
# LISTENING FUNCTION (INTERRUPTIBLE)
# ======================================
stop_listening = False

def listen_for_command():
    """Listen for user's voice input, allow To Enter to interrupt."""
    global stop_listening
    r = sr.Recognizer()
    try:
        with sr.Microphone() as source:
            print("🎧 Listening... (Press Enter to interrupt)")
            stop_listening = False

            # Thread to allow Enter key to stop listening
            def wait_for_enter():
                global stop_listening
                input()
                stop_listening = True

            threading.Thread(target=wait_for_enter, daemon=True).start()

            audio = r.listen(source, timeout=5, phrase_time_limit=10)
            if stop_listening:
                print("⏹ Listening interrupted.")
                return ""
            command = r.recognize_google(audio).lower()
            print(f"You said: {command}")
            return command
    except Exception:
        return ""


# ======================================
# DAILY QUOTE FEATURE (One per day)
# ======================================
def daily_quote():
    """Fetch and speak one quote per day (cached)."""
    file_path = "daily_quote.json"
    today = datetime.date.today().isoformat()

    # Check if a quote is already saved for today
    if os.path.exists(file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            if data.get("date") == today:
                quote = data.get("quote")
                author = data.get("author")
                speak(f"Here’s your motivational quote for today: {quote} — by {author}.")
                return

    # Otherwise fetch a new one
    try:
        response = requests.get("https://api.quotable.io/random")
        data = response.json()
        quote = data["content"]
        author = data["author"]

        # Save the new quote
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump({"date": today, "quote": quote, "author": author}, f, indent=4)

        speak(f"Here’s your motivational quote for today: {quote} — by {author}.")
    except:
        speak("Sorry, I couldn’t get your quote today.")


# ======================================
# MAPS ROUTE FEATURE (TALKING)
# ======================================
def get_fastest_route(destination: str):
    """Open Google Maps with directions to destination."""
    try:
        speak(f"Let me check the fastest route to {destination}.")
        maps_url = f"https://www.google.com/maps/dir//{destination.replace(' ', '+')}"
        speak(f"Opening Google Maps for {destination}.")
        webbrowser.open(maps_url)
        time.sleep(2)
        speak("You can now follow the route on Google Maps. Drive safely!")
    except Exception as e:
        print("Map error:", e)
        speak("I couldn’t open the route right now.")



# ======================================
# WIKIPEDIA (using wikipediaapi)
# ======================================
def get_definition(topic):
    """Use WikipediaAPI to fetch clean summary info."""
    wiki = wikipediaapi.Wikipedia(user_agent='AlexaAssistant/1.0', language='en')
    page = wiki.page(topic)

    if page.exists():
        summary = page.summary[:600]
        if "may refer to" in summary.lower():
            speak(f"There are several meanings for {topic}. Can you be more specific?")
        else:
            speak(summary)
    else:
        speak(f"Sorry, I couldn’t find detailed information for {topic}.")


# ======================================
# PLAY FEATURE (using pywhatkit)
# ======================================
def play_on_youtube(query):
    """Play the requested song or video using pywhatkit."""
    try:
        speak(f"Playing {query} on YouTube.")
        pywhatkit.playonyt(query)
    except Exception as e:
        print("YouTube error:", e)
        speak("Sorry, I couldn’t play that video right now.")


# ======================================
# MAIN LOOP
# ======================================
def main():
    speak("Good day! I’m Alexa, your personal desktop assistant.")
    time.sleep(1)
    daily_quote()
    time.sleep(1)
    speak("To start a command, say 'Alexa'.")

    while True:
        command = listen_for_command()

        if not command:
            continue

        # Require "Alexa" wake word
        if not command.startswith("alexa"):
            speak("Please start your command by saying 'Alexa'.")
            continue

        # Remove wake word
        command = command.replace("alexa", "").strip()

        if "stop" in command or "bye" in command or "exit" in command:
            speak("Goodbye! Take care.")
            break

        elif "time" in command:
            now = datetime.datetime.now().strftime("%I:%M %p")
            speak(f"The current time is {now}.")

        elif "quote" in command:
            daily_quote()

        elif "route" in command or "directions" in command:
            speak("Where do you want to go?")
            destination = listen_for_command()
            get_fastest_route(destination)

        elif "what does" in command or "who is" in command or "define" in command:
            topic = command.replace("what is", "").replace("who is", "").replace("define", "").strip()
            if topic:
                get_definition(topic)
            else:
                speak("Please tell me what you want to know.")

        elif "play" in command:
            song = command.replace("play", "").strip()
            if song:
                play_on_youtube(song)
            else:
                speak("What would you like me to play?")

        else:
            speak("Sorry, I didn’t understand that. Try again.")


# ======================================
# RUN PROGRAM
# ======================================
if __name__ == "__main__":
    main()
