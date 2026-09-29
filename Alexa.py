import tkinter as tk
from tkinter import scrolledtext
import speech_recognition as sr
import win32com.client
import datetime
import wikipedia
import webbrowser
import os
import random

speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Rate = 0  

def speak(text):
    try:
        speaker.Speak(text)
    except Exception as e:
        print("Voice error:", e)

root = tk.Tk()
root.title("Jarvis")
root.geometry("540x660")
root.configure(bg="#1e1e2e")

title_label = tk.Label(
    root, 
    text="Jarvis", 
    font=("Segoe UI", 14, "bold"), 
    bg="#1e1e2e", 
    fg="#89b4fa"
)
title_label.pack(pady=(12, 0))

center_frame = tk.Frame(root, bg="#1e1e2e")
center_frame.pack(expand=True, fill=tk.BOTH, pady=40)

hero_title = tk.Label(
    center_frame,
    text="Hey there!\nWhat can I do for you right now?",
    font=("Segoe UI", 18, "bold"),
    fg="#cdd6f4",
    bg="#1e1e2e",
    justify=tk.CENTER
)
hero_title.pack(expand=True)


chat_area = scrolledtext.ScrolledText(
    root, 
    wrap=tk.WORD, 
    font=("Segoe UI", 11), 
    bg="#181825", 
    fg="#cdd6f4",
    insertbackground="white",
    bd=0,
    padx=10,
    pady=10
)

PLACEHOLDER_TEXT = "Ask Jarvis..."
COLOR_PLACEHOLDER = "#6c7086"
COLOR_TEXT = "#ffffff"

def ensure_chat_view():
    
    if center_frame.winfo_ismapped():
        center_frame.pack_forget()
        chat_area.pack(padx=15, pady=10, fill=tk.BOTH, expand=True)

def update_chat(sender, message):
    ensure_chat_view()
    chat_area.configure(state=tk.NORMAL)
    chat_area.insert(tk.END, f"{sender}: {message}\n\n")
    chat_area.see(tk.END)
    chat_area.configure(state=tk.DISABLED)
    root.update()

def reset_ui():
    mic_btn.config(state=tk.NORMAL, bg="#313244", text="🎤")
    send_btn.config(state=tk.NORMAL)
    entry_box.config(state=tk.NORMAL)
    if not entry_box.get().strip():
        entry_box.delete(0, tk.END)
        entry_box.insert(0, PLACEHOLDER_TEXT)
        entry_box.config(fg=COLOR_PLACEHOLDER)
    root.update()

def get_wiki_summary(query):
    try:
        search_list = wikipedia.search(query, results=5)
        if not search_list:
            return None
        
        for title in search_list:
            try:
                summary = wikipedia.summary(title, sentences=2, auto_suggest=False)
                return summary
            except (wikipedia.exceptions.DisambiguationError, wikipedia.exceptions.PageError):
                continue
        return None
    except Exception:
        return None

def process_command(query):
    query_lower = query.lower().strip()

    if any(word in query_lower for word in ["hello","hello jarvis","hi", "hi jarvis"])  :
        reply = "Hello! How can I help you today?"
        update_chat("Bot", reply)
        speak(reply)

    elif "time" in query_lower:
        current_time = datetime.datetime.now().strftime("%I:%M %p")
        reply = f"The current time is {current_time}"
        update_chat("Bot", reply)
        speak(reply)

    elif "date" in query_lower or "day" in query_lower:
        now = datetime.datetime.now()
        reply = f"Today is {now.strftime('%A, %B %d, %Y')}"
        update_chat("Bot", reply)
        speak(reply)

    elif "joke" in query_lower:
        jokes_list = [
            "Why do programmers prefer dark mode? Because light attracts bugs!",
            "Why did the computer go to the doctor? Because it had a virus!",
            "There are 10 types of people in the world: those who understand binary, and those who don't.",
            "Why was the JavaScript developer sad? Because he didn't know how to null his feelings!"
        ]
        joke = random.choice(jokes_list)
        update_chat("Bot", joke)
        speak(joke)

    elif "word" in query_lower:
        reply = "Opening Microsoft Word..."
        update_chat("Bot", reply)
        speak(reply)
        os.system("start winword")

    elif "excel" in query_lower:
        reply = "Opening Microsoft Excel..."
        update_chat("Bot", reply)
        speak(reply)
        os.system("start excel")

    elif "powerpoint" in query_lower or "ppt" in query_lower:
        reply = "Opening Microsoft PowerPoint..."
        update_chat("Bot", reply)
        speak(reply)
        os.system("start powerpnt")

    elif "files" in query_lower or "folder" in query_lower or "explorer" in query_lower:
        reply = "Opening File Explorer..."
        update_chat("Bot", reply)
        speak(reply)
        os.system("explorer")

    elif "calculator" in query_lower or "calc" in query_lower:
        reply = "Opening Calculator..."
        update_chat("Bot", reply)
        speak(reply)
        os.system("calc")

    elif "notepad" in query_lower:
        reply = "Opening Notepad..."
        update_chat("Bot", reply)
        speak(reply)
        os.system("notepad")

    elif query_lower.startswith("play"):
        song_name = query_lower.replace("play", "").strip()
        reply = f"Playing {song_name} on YouTube..."
        update_chat("Bot", reply)
        speak(reply)
        webbrowser.open(f"https://www.youtube.com/results?search_query={song_name}")

    elif "open youtube" in query_lower:
        reply = "Opening YouTube..."
        update_chat("Bot", reply)
        speak(reply)
        webbrowser.open("https://youtube.com")

    elif "open google" in query_lower:
        reply = "Opening Google..."
        update_chat("Bot", reply)
        speak(reply)
        webbrowser.open("https://google.com")

    elif "how are you" in query_lower:
        reply = "I am doing great and ready to help you! How can I assist you today?"
        update_chat("Bot", reply)
        speak(reply)

    elif any(word in query_lower for word in ["bye", "goodbye", "exit", "close assistant", "stop"]):
        goodbye_responses = [
            "Goodbye! Have a great day ahead!",
            "Bye! Let me know if you need anything else later.",
            "Signing off now. Have a wonderful time!"
        ]
        reply = random.choice(goodbye_responses)
        update_chat("Bot", reply)
        speak(reply)
        root.after(2000, root.destroy)

    elif "who is" in query_lower or "wikipedia" in query_lower or "what is" in query_lower:
        search_topic = query_lower.replace("who is", "").replace("wikipedia", "").replace("what is", "").strip()
        
        if search_topic == "mgr" or search_topic == "m.g.r":
            search_topic = "M. G. Ramachandran"
        elif search_topic == "vijay":
            search_topic = "Vijay (actor)"

        update_chat("Bot", f"Searching Wikipedia for {search_topic}...")
        
        result_text = get_wiki_summary(search_topic)
        if result_text:
            update_chat("Bot", result_text)
            speak(result_text)
        else:
            reply = f"மன்னிக்கவும், {search_topic} பற்றி விக்கிப்பீடியாவில் தகவல் கிடைக்கவில்லை."
            update_chat("Bot", reply)
            speak(reply)

    else:
        reply = f"Searching for: {query}"
        update_chat("Bot", reply)
        speak(reply)
        webbrowser.open(f"https://www.google.com/search?q={query}")

    reset_ui()

def on_entry_click(event):
    if entry_box.get() == PLACEHOLDER_TEXT:
        entry_box.delete(0, tk.END)
        entry_box.config(fg=COLOR_TEXT)

def on_focusout(event):
    if not entry_box.get().strip():
        entry_box.insert(0, PLACEHOLDER_TEXT)
        entry_box.config(fg=COLOR_PLACEHOLDER)

def send_text(event=None):
    text = entry_box.get().strip()
    if text and text != PLACEHOLDER_TEXT:
        ensure_chat_view()
        update_chat("You", text)
        entry_box.delete(0, tk.END)
        send_btn.config(state=tk.DISABLED)
        mic_btn.config(state=tk.DISABLED)
        entry_box.config(state=tk.DISABLED)
        root.update()
        process_command(text)

def listen_voice():
    recognizer = sr.Recognizer()
    mic_btn.config(state=tk.DISABLED, bg="#f38ba8", text="🔴")
    send_btn.config(state=tk.DISABLED)
    entry_box.config(state=tk.DISABLED)
    root.update()

    with sr.Microphone() as source:
        try:
            recognizer.adjust_for_ambient_noise(source, duration=0.6)
            audio = recognizer.listen(source, timeout=5, phrase_time_limit=8)
            query = recognizer.recognize_google(audio, language='en-in')
            ensure_chat_view()
            update_chat("You", query)
            process_command(query)
        except sr.WaitTimeoutError:
            update_chat("System", "குரல் எதுவும் கேட்கவில்லை.")
            reset_ui()
        except Exception:
            update_chat("System", "குரலைப் புரிந்துகொள்ள முடியவில்லை.")
            reset_ui()

bottom_frame = tk.Frame(root, bg="#1e1e2e")
bottom_frame.pack(fill=tk.X, padx=15, pady=15)

entry_box = tk.Entry(
    bottom_frame, 
    font=("Segoe UI", 12), 
    bg="#313244", 
    fg=COLOR_PLACEHOLDER, 
    insertbackground="white", 
    relief=tk.FLAT
)
entry_box.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8, padx=(0, 6))

entry_box.insert(0, PLACEHOLDER_TEXT)
entry_box.bind("<FocusIn>", on_entry_click)
entry_box.bind("<FocusOut>", on_focusout)
entry_box.bind("<Return>", send_text)

mic_btn = tk.Button(
    bottom_frame, 
    text="🎤", 
    font=("Segoe UI", 12), 
    bg="#313244", 
    fg="#ffffff", 
    relief=tk.FLAT, 
    padx=10, 
    pady=3,
    cursor="hand2",
    command=listen_voice
)
mic_btn.pack(side=tk.LEFT, padx=(0, 6))

send_btn = tk.Button(
    bottom_frame, 
    text="➤", 
    font=("Segoe UI", 12, "bold"), 
    bg="#89b4fa", 
    fg="#11111b", 
    relief=tk.FLAT, 
    padx=12, 
    pady=3,
    cursor="hand2",
    command=send_text
)
send_btn.pack(side=tk.LEFT)

root.mainloop()