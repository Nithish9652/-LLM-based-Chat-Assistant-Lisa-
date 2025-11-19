# Lisa – LLM-Based Python Voice Assistant

Lisa is a smart desktop voice assistant built using Python. It performs system operations, reads emails, takes screenshots, extracts on-screen text using OCR, plays YouTube videos, opens apps/websites, and responds to AI prompts using the Gemini API.

---

## 🚀 Features

### 🎤 Voice Input
- Real-time Google Speech Recognition

### 🔊 Text-to-Speech
- Windows SAPI voice output

### 🌐 Open Websites & Applications
- “Open Google”
- “Open chrome”
- “Open facebook.com”

### 🌞 Brightness Control
- “Brightness up”
- “Brightness down”

### 🔊 Volume Control
- “Volume up”
- “Volume down”

### 📧 Read Gmail Inbox
- Reads latest email using IMAP
- Speaks sender, subject & body

### 🖼 Screenshot Capture
- “Lisa take screenshot”
- Saves screenshot as `screen.png`

### 👁 Screen Reader (OCR)
- “Lisa read screen”
- Takes screenshot → Extracts text with Tesseract OCR → Speaks text

### 🎥 YouTube Video Player
- Search & play videos using pytube

### 🤖 AI Integration (Gemini API)
- “AI explain quantum physics”
- “AI write a poem”
- Uses Google Gemini API for responses

---

## 📦 Installation

### 1️⃣ Install Python Dependencies

```bash
pip install SpeechRecognition pywin32 screen_brightness_control pycaw requests pytube3 pillow pytesseract pyautogui imaplib2
