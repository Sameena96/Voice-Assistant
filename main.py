import pyttsx3
import wikipedia
import webbrowser
import docx
import PyPDF2
import threading
import os
import subprocess

# Initialize pyttsx3
engine = pyttsx3.init()
engine.setProperty('rate', 150)
engine.setProperty('volume', 1.0)

# Event to control speaking
speak_event = threading.Event()

# Lock to prevent simultaneous access to the TTS engine
engine_lock = threading.Lock()

# Function to speak text with the ability to stop
def speak_text(text):
    speak_event.set()  # Allow speaking
    with engine_lock:
        for line in text.splitlines():
            if not speak_event.is_set():  # Stop if speaking is disabled
                break
            engine.say(line)
            engine.runAndWait()
        engine.stop()  # Ensure engine stops completely

# Function to stop speaking
def stop_speaking():
    speak_event.clear()  # Disable speaking
    engine.stop()  # Immediately stop the TTS engine

# Function to handle interruptible tasks (reading documents, etc.)
def run_with_stop(func, *args):
    # Create and start a thread to run the task
    task_thread = threading.Thread(target=func, args=args, daemon=True)
    task_thread.start()
    return task_thread

# Function to search Google
def search_google(query):
    print(f"Searching Google for: {query}")
    speak_text(f"Searching Google for {query}")
    webbrowser.open(f"https://www.google.com/search?q={query}")

# Function to search Wikipedia
def search_wikipedia(query):
    try:
        print(f"Searching Wikipedia for: {query}")
        summary = wikipedia.summary(query, sentences=2)
        speak_text(summary)
    except wikipedia.exceptions.DisambiguationError:
        speak_text("The topic is ambiguous, please be more specific.")
    except wikipedia.exceptions.PageError:
        speak_text("I couldn't find information on Wikipedia for that topic.")
    except Exception as e:
        print(f"Error searching Wikipedia: {e}")
        speak_text("There was an error trying to access Wikipedia.")

# Function to play video on YouTube
def play_youtube_video(query):
    print(f"Playing {query} on YouTube.")
    speak_text(f"Playing {query} on YouTube.")
    webbrowser.open(f"https://www.youtube.com/results?search_query={query}")

# Function to read MS Word document
def read_word_doc(file_path):
    try:
        if not os.path.isfile(file_path):
            print("The file path provided is incorrect.")
            speak_text("The file path provided is incorrect.")
            return

        doc = docx.Document(file_path)
        full_text = [para.text for para in doc.paragraphs if para.text.strip()]
        text = '\n'.join(full_text)

        if not text:
            print("The document is empty or unreadable.")
            speak_text("The document is empty or unreadable.")
            return

        # Start speaking the document text in a new thread
        run_with_stop(speak_text, text)
    except Exception as e:
        print(f"Error reading Word document: {e}")
        speak_text("Unable to read the Word document.")

# Function to read PDF file
def read_pdf(file_path):
    try:
        if not os.path.isfile(file_path):
            print("The file path provided is incorrect.")
            speak_text("The file path provided is incorrect.")
            return

        # Open the PDF in the system's default viewer
        print(f"Opening PDF: {file_path}")
        speak_text("Opening the PDF in your default viewer.")

        if os.name == 'nt':  # For Windows
            os.startfile(file_path)
        elif os.name == 'posix':  # For Linux and macOS
            if "darwin" in os.uname().sysname.lower():  # macOS
                subprocess.Popen(["open", file_path], shell=False)
            else:  # Linux
                subprocess.Popen(["xdg-open", file_path], shell=False)

        # Extract and read the text content
        with open(file_path, "rb") as file:
            reader = PyPDF2.PdfReader(file)
            text = ""
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"

        if not text.strip():
            print("The PDF is empty or unreadable.")
            speak_text("The PDF is empty or unreadable.")
            return

        # Start speaking the PDF text in a new thread
        run_with_stop(speak_text, text)
    except Exception as e:
        print(f"Error reading PDF document: {e}")
        speak_text("Unable to read the PDF document.")

# Function to handle user input (non-blocking)
def handle_input():
    while True:
        stop_input = input("Type 'stop' to stop task or 'exit' to quit: ").strip().lower()
        if stop_input == "stop":
            stop_speaking()
            print("Task interrupted!")
            break
        elif stop_input == "exit":
            print("Goodbye!")
            os._exit(0)  # Immediately exit the program

# Main function to handle commands
def main():
    speak_text("Hello! How can I assist you today?")

    while True:
        print("\nOptions:")
        print("1. Speak text")
        print("2. Google Search")
        print("3. Wikipedia Search")
        print("4. Play YouTube Video")
        print("5. Read Word Document")
        print("6. Read PDF")
        print("7. Exit")

        user_input = input("Enter the option number or type your command: ").lower()

        if user_input == "1":
            text = input("Enter the text you want me to speak: ")
            run_with_stop(speak_text, text)

        elif user_input == "2":
            query = input("Enter the Google search query: ")
            search_google(query)

        elif user_input == "3":
            query = input("Enter the Wikipedia search query: ")
            search_wikipedia(query)

        elif user_input == "4":
            query = input("Enter the video you want to play on YouTube: ")
            play_youtube_video(query)

        elif user_input == "5":
            file_path = input("Enter the path to the Word document (.docx): ")
            run_with_stop(read_word_doc, file_path)

        elif user_input == "6":
            file_path = input("Enter the path to the PDF file: ")
            run_with_stop(read_pdf, file_path)

        elif user_input == "7":
            speak_text("Goodbye, Have a nice day")
            break

        else:
            speak_text("I'm sorry, I didn't understand that command.")

        # Start a separate thread for handling stop input
        input_thread = threading.Thread(target=handle_input, daemon=True)
        input_thread.start()

if __name__ == "__main__":
    main()
