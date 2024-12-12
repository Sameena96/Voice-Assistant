**Interactive Voice Assistant**

This Python-based project is an interactive voice assistant capable of performing various tasks such as reading documents, searching the web, and playing YouTube videos. The assistant utilizes text-to-speech (TTS) and speech interruption capabilities for seamless interactions.
**Features**

1. Speak Text: Input text for the assistant to read aloud.
2. Google Search: Perform a Google search for a given query.
3. Wikipedia Search: Retrieve summarized information from Wikipedia.
4. Play YouTube Video: Search and play YouTube videos.
5. Read Word Documents: Open and read .docx files aloud.
6. Read PDF Files: Open and read PDF documents.
7. Stop Tasks: Interrupt any ongoing task, such as speech or reading.
8. Exit Program: Terminate the program gracefully.

**Requirements Dependencies**

Install the required libraries by running:

      _pip install pyttsx3 wikipedia python-docx PyPDF2_

**Additional Requirements**

1. pyttsx3: Text-to-speech conversion.
2. wikipedia: Access Wikipedia summaries.
3. webbrowser: Open web pages.
4. python-docx: Read Microsoft Word documents.
5. PyPDF2: Extract text from PDF files.
6. threading: Handle simultaneous tasks.
7. subprocess: Open files in the default system viewer.

**How to Use**

Clone the repository or download the script.
Ensure all dependencies are installed.
Run the script:
           _  python script_name.py_

Follow the menu to perform desired tasks.

**Program Flow**

1. On launch, the assistant greets the user.
2. Users can choose tasks from the menu:

    - Option 1: Enter text to be read aloud.
    - Option 2: Input a query to search on Google.
    - Option 3: Input a query to search Wikipedia.
    - Option 4: Enter a video name to play on YouTube.
    - Option 5: Provide the path to a .docx file to read.
    - Option 6: Provide the path to a PDF file to read.
    - Option 7: Exit the program.

3. To interrupt any task (e.g., ongoing speech), type stop.
4. To terminate the program, type exit or select the exit option.

**Limitations**

- PDF Parsing: Some PDFs with complex formatting may not extract text properly.
- Wikipedia Search: May fail for ambiguous or specific queries.

**Contributing**

Feel free to fork the repository, create a branch, and submit pull requests for improvements or bug fixes.

