# GiggleGenius — Desktop Voice Assistant

**GiggleGenius** is a Python-based desktop voice assistant for Windows that listens for natural-language commands and performs tasks such as web search, playing YouTube, file-explorer automation, OCR-based clicking, system control (lock, shutdown), and reporting system stats. It combines speech recognition, text-to-speech, GUI controls, and automation libraries to provide hands-free control of common desktop tasks.

---

## Key Features

- Speech-to-text and text-to-speech interaction (continuous listening).
- Wikipedia lookups and web searches.
- Play music on YouTube via voice commands.
- File Explorer automation: navigate drives/folders, list files, rename, search.
- OCR-driven clicking: locate a word on-screen and click it.
- System controls: lock, shutdown, restart, privacy mode (minimize + mute).
- System information: battery status, CPU and RAM usage.
- Simple Tkinter GUI to start/stop the assistant and set user name.

---

## Quick Start

### Requirements

- Windows (project uses Windows-specific APIs: SAPI5, win32com, Explorer automation)
- Python 3.8+
- Tesseract OCR installed (for OCR features):
  - Default path used in code: `C:\Program Files\Tesseract-OCR\tesseract.exe`
- A working microphone and speakers.

### Install Python dependencies

Create a virtual environment (recommended) and install dependencies:


> If you don't have a `requirements.txt`, typical packages used by this project include:
> `pyttsx3`, `SpeechRecognition`, `pytesseract`, `Pillow`, `pyautogui`, `psutil`, `pywhatkit`, `pyjokes`, `pypiwin32` (pywin32), `wikipedia`.

### Configure Tesseract

If Tesseract is installed to a non-default location, update `pytesseract.pytesseract.tesseract_cmd` in the code to point to the correct path.

### Run the Assistant

From the project folder:

```bash
python main.py
```

Use the GUI to enter your name and click **Start Voice Assistant**. The assistant will greet you and begin listening.

---

## Common Voice Commands

- `wikipedia <topic>` — speak a short summary from Wikipedia.
- `play <song name>` — plays the song on YouTube.
- `open youtube` / `open google` — open websites.
- `search <query>` — perform a Google search.
- `where is <location>` — opens the location in Google Maps.
- `the time` — tells the current time.
- `battery status` — reports battery percentage and charger state.
- `system performance` — reports CPU and RAM usage.
- `lock the computer` — locks Windows session.
- `privacy mode` / `turn off privacy mode` — minimize/unmute actions.
- `shutdown system` / `restart system` — shut down or restart (use carefully).
- `open file explorer`, `change directory to <drive>`, `go to <folder>`, `list files`, `rename file <old> to <new>`, `find file <term>` — explorer automation.
- `click <word>` — OCR-based click on a visible word.
- `type` — assistant prompts for text then types it.
- `clear` — select-all + backspace.
- `write that` — write an admin log entry to `admin_log.txt`.
- `exit` — stop the assistant.

---

## Configuration & Notes

- The assistant uses Google’s free speech recognition via the `speech_recognition` package; network access is required.
- The project is **Windows-first**. Some features call Windows APIs (ctypes, win32com) and won't work on macOS/Linux without changes.
- OCR accuracy depends on screen resolution, contrast, and Tesseract quality.
- Some commands (shutdown/restart) are destructive — consider adding a confirmation flow before executing.

---

## Troubleshooting

- **Microphone not detected / recognition fails**: ensure microphone is configured in Windows and accessible by Python; update permissions.
- **Tesseract OCR errors**: ensure `tesseract.exe` path is correct and Tesseract is installed.
- **pyautogui actions mis-click**: UI scaling or multiple monitors can change coordinates; test carefully.
- **Slow response**: speech recognition calls Google’s API which may introduce latency.

---

## Contributing

Contributions are welcome. If you'd like to contribute:

1. Fork the repository.
2. Create a feature branch: `git checkout -b feature-name`.
3. Commit your changes and open a pull request.

Please follow the code style, add tests where appropriate, and document new features.

---

## License & Credits

- This repository template was prepared by the project author. Add your desired license (e.g., MIT) here.
- Core libraries used: pyttsx3, SpeechRecognition, pytesseract, pyautogui, psutil, pywhatkit, pyjokes, wikipedia, Pillow.

---

## Contact

Project author: Mursal Bajwa  
GitHub: https://github.com/MursalBajwa/Giggle-genius-Desktop-voice-assistant
