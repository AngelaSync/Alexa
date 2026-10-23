import speech_recognition as sr
import threading
import datetime
import requests
import wikipediaapi
import win32com.client
import os
import json
import pywhatkit
import time
import webbrowser
import winsound
import urllib.parse
import subprocess

# ======================================
# SETTINGS
# ======================================

USE_VOICEVOX = False
VOICEVOX_SPEAKER = 3  # Default speaker (Zundamon)
VOICEVOX_SPEAKER_LIST = []  # List of all speakers from VoiceVox
speaker = win32com.client.Dispatch("SAPI.SpVoice")
stop_listening = False
interrupt_triggered = False  # used when pressing Enter mid-speech



# ======================================
# VOICE HELPERS
# ======================================

def list_microsoft_voices():
    voices = speaker.GetVoices()
    print("\n🗣️ Available Microsoft Voices:")
    for i, v in enumerate(voices):
        print(f"{i + 1}: {v.GetDescription()}")
    return voices


def choose_microsoft_voice():
    voices = list_microsoft_voices()
    choice = input("\nChoose Microsoft voice number: ").strip()
    if choice.isdigit() and 1 <= int(choice) <= len(voices):
        speaker.Voice = voices[int(choice) - 1]
        print(f"✅ Using: {speaker.Voice.GetDescription()}")
    else:
        print("⚠️ Invalid choice, keeping current voice.")


def list_voicevox_voices():
    """List VoiceVox voices using simple numbers 1–N"""
    global VOICEVOX_SPEAKER_LIST
    try:
        response = requests.get("http://127.0.0.1:50021/speakers", timeout=3)
        speakers = response.json()
        flat_list = []
        print("\n🎤 Available VoiceVox Voices:")
        idx = 1
        for group in speakers:
            for style in group["styles"]:
                flat_list.append(style["id"])
                print(f"{idx:>2}: {group['name']} ({style['name']})")
                idx += 1
        VOICEVOX_SPEAKER_LIST = flat_list
        return flat_list
    except Exception as e:
        print(f"⚠️ VoiceVox not reachable ({e})")
        VOICEVOX_SPEAKER_LIST = []
        return []


def choose_voicevox_voice():
    """Let user select a VoiceVox voice using 1–N numbers"""
    global VOICEVOX_SPEAKER
    flat_list = list_voicevox_voices()
    if not flat_list:
        print("❌ No VoiceVox voices available.")
        return
    choice = input("\nEnter VoiceVox voice number (1–{0}): ".format(len(flat_list))).strip()
    if choice.isdigit() and 1 <= int(choice) <= len(flat_list):
        VOICEVOX_SPEAKER = flat_list[int(choice) - 1]
        print(f"✅ VoiceVox voice set to #{choice}")
    else:
        print("⚠️ Invalid input, keeping current voice.")


# ======================================
# SPEAK FUNCTION
# ======================================
def watch_for_enter():
    """Runs in background — stops speech and signals interruption when Enter is pressed."""
    global interrupt_triggered
    while True:
        try:
            input()
            interrupt_triggered = True
            try:
                speaker.Speak("", 3)  # stop speech
            except Exception:
                pass
            print("\n⏹ Interrupted — returning to active mode...")
        except EOFError:
            break

def speak(text):
    """Speak text using VoiceVox or Microsoft SAPI seamlessly"""
    global USE_VOICEVOX, VOICEVOX_SPEAKER
    print(f"\nAlexa 🎤: {text}\n")

    if USE_VOICEVOX:
        try:
            requests.get("http://127.0.0.1:50021/speakers", timeout=2)
            encoded = urllib.parse.quote(text)

            subprocess.run([
                "curl", "-s", "-X", "POST",
                f"http://127.0.0.1:50021/audio_query?text={encoded}&speaker={VOICEVOX_SPEAKER}",
                "-o", "query.json"
            ], check=True)

            subprocess.run([
                "curl", "-s", "-H", "Content-Type: application/json",
                "-X", "POST", "-d", "@query.json",
                f"http://127.0.0.1:50021/synthesis?speaker={VOICEVOX_SPEAKER}",
                "-o", "voicevox.wav"
            ], check=True)

            winsound.PlaySound("voicevox.wav", winsound.SND_FILENAME)
        except Exception as e:
            print(f"⚠️ VoiceVox error ({e}) — using Microsoft voice instead.")
            USE_VOICEVOX = False
            speaker.Speak(text)
    else:
        speaker.Speak(text)


# ======================================
# LISTENING FUNCTION
# ======================================2




def listen_for_command():
    global stop_listening
    r = sr.Recognizer()
    try:
        with sr.Microphone() as source:
            print("🎧 Listening... (Press Enter to interrupt)")
            stop_listening = False

            def wait_for_enter():
                global stop_listening
                input()
                stop_listening = True
                try:
                    speaker.Speak("", 3)  # stop any current speech instantly
                    print("⏹ Interrupted — returning to listening mode...")
                except Exception:
                    pass

            threading.Thread(target=wait_for_enter, daemon=True).start()
            audio = r.listen(source, timeout=5, phrase_time_limit=8)

            if stop_listening:
                return ""  # signal to main loop to just continue

            command = r.recognize_google(audio).lower()
            print(f"You said: {command}")
            return command
    except Exception:
        return ""


# ======================================
# DAILY QUOTE
# ======================================

def daily_quote():
    file_path = "daily_quote.json"
    today = datetime.date.today().isoformat()

    if os.path.exists(file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            if data.get("date") == today:
                quote = data["quote"]
                author = data["author"]
                speak(f"Here’s your motivational quote for today: {quote} — by {author}.")
                return

    try:
        response = requests.get("https://api.quotable.io/random", timeout=5)
        data = response.json()
        quote = data.get("content", "Keep going, you're doing great!")
        author = data.get("author", "Unknown")
    except Exception:
        quote, author = ("Keep pushing forward, even small steps count.", "Unknown")

    with open(file_path, "w", encoding="utf-8") as f:
        json.dump({"date": today, "quote": quote, "author": author}, f, indent=4)

    speak(f"Here’s your motivational quote for today: {quote} — by {author}.")


# ======================================
# COMMAND HELP LIST
# ======================================

def show_commands():
    commands = [
        "Alexa quote — Get a motivational quote.",
        "Alexa time — Ask for the current time.",
        "Alexa route to [place] — Get directions on Google Maps.",
        "Alexa what is [topic] — Learn from Wikipedia.",
        "Alexa who is [person] - Learn from Wikipedia"
        "Alexa play [song or video] — Play on YouTube.",
        "Alexa use voicevox — Switch to VoiceVox TTS.",
        "Alexa use microsoft — Switch to Microsoft TTS.",
        "Alexa change voice — Choose a different speaker.",
        "Alexa help — Hear the list of commands.",
        "Alexa stop — Exit the assistant."
    ]
    print("\n🧠 Available Alexa Commands:")
    for c in commands:
        print(" - " + c)
    speak("Here are some things you can ask me to do. Say 'help' anytime to hear this again.")


# ======================================
# OTHER FEATURES
# ======================================

def get_fastest_route(destination):
    try:
        speak(f"Let me check the fastest route to {destination}.")
        maps_url = f"https://www.google.com/maps/dir/?api=1&destination={destination.replace(' ', '+')}"
        webbrowser.open(maps_url)
        speak("Google Maps opened. Drive safely!")
    except:
        speak("I couldn’t open the route right now.")

def get_definition(topic):
    wiki = wikipediaapi.Wikipedia(user_agent='AlexaAssistant/1.0', language='en')
    page = wiki.page(topic)
    if page.exists():
        speak(page.summary[:600])
    else:
        speak(f"Sorry, I couldn’t find information about {topic}.")

def play_on_youtube(query):
    try:
        speak(f"Playing {query} on YouTube.")
        pywhatkit.playonyt(query)
    except:
        speak("Sorry, I couldn’t play that video right now.")


# ======================================
# MAIN PROGRAM
# ======================================

def main():

    global USE_VOICEVOX

    try:
        speak("Good day! I’m Alexa, your personal desktop assistant.")
        time.sleep(1)
        daily_quote()
        time.sleep(1)

        # Choose initial voice system
        voice_choice = input("Which voice would you like to use?\n1 - Microsoft SAPI\n2 - VoiceVox\nChoice: ").strip()
        USE_VOICEVOX = (voice_choice == "2")
        if USE_VOICEVOX:
            choose_voicevox_voice()
        else:
            choose_microsoft_voice()
        speak(f"{'VoiceVox' if USE_VOICEVOX else 'Microsoft'} voice selected.")

        # Choose input mode
        mode = input("Would you like to talk by 'voice' or by 'typing'? ").strip().lower()
        if mode not in ["voice", "type"]:
            mode = "type"
        speak(f"{'Voice' if mode == 'voice' else 'Typing'} mode activated.")

        show_commands()

        while True:
            try:
                command = listen_for_command() if mode == "voice" else input("You: ").lower().strip()
            except KeyboardInterrupt:
                print("\n⏹ Interrupted while listening.")
                speak("Okay, stopping now.")
                break

            if not command:
                continue
            if not command.startswith("alexa"):
                speak("Please start your command by saying 'Alexa'.")
                continue

            command = command.replace("alexa", "").strip()

            if any(kw in command for kw in ["stop", "bye", "exit", "quit"]):
                speak("Goodbye! Take care.")
                break

            elif "use voicevox" in command:
                USE_VOICEVOX = True
                choose_voicevox_voice()
                speak("VoiceVox voice activated.")

            elif "use microsoft" in command:
                USE_VOICEVOX = False
                choose_microsoft_voice()
                speak("Microsoft voice activated.")

            elif "change voice" in command:
                if USE_VOICEVOX:
                    choose_voicevox_voice()
                else:
                    choose_microsoft_voice()

            elif "time" in command:
                now = datetime.datetime.now().strftime("%I:%M %p")
                speak(f"The current time is {now}.")

            elif "quote" in command:
                daily_quote()

            elif "help" in command:
                show_commands()

            elif "route" in command or "directions" in command:
                speak("Where do you want to go?")
                destination = listen_for_command() if mode == "voice" else input("Destination: ")
                if destination:
                    get_fastest_route(destination)
                else:
                    speak("I didn’t catch the destination.")

            elif "what is" in command or "who is" in command or "define" in command:
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
                speak("Sorry, I didn’t understand that. Try again or say 'Alexa help'.")

    except KeyboardInterrupt:
        print("\n⏹ Program interrupted by user.")
        speak("Goodbye for now.")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
    finally:
        print("💤 Exiting Alexa assistant.")


# ======================================
# RUN
# ======================================

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n⏹ Exiting program safely.")

