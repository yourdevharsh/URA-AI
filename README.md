# URA AI

URA is a desktop assistant designed to help users perform tasks in **Microsoft Word** using a combination of:
- Real-time screen capture  
- YOLO object detection  
- A local LLM (DeepSeek or Gemma via Ollama)  
- Overlay-based visual guidance  

This tool identifies features inside Word, interprets user queries, and highlights the corresponding UI buttons directly on the screen — acting as an intelligent, interactive Office tutor.

---

## ✨ Features

- **Ask Anything**: “How to bold text?”, “Insert a chart”, “Check spelling”, etc.
- **Local AI Reasoning** (DeepSeek R1 or Gemma via Ollama)
- **Smart Label Matching** for Office features
- **Real-Time YOLO Detection** of Word Ribbon UI
- **Live Overlay Highlighting** on the real application window
- **Smooth Chat UI** with animated scrolling
- **Automatic Word Window Detection**
- **Minimize/Close buttons fully functional**
- **Chatbot auto-hides itself during screenshots**

---

## 🧠 How It Works (Short Summary)

1. You type a question into the chatbot  
2. The system:
   - Detects which Office feature you want  
   - Determines the correct button label  
   - Generates action steps using DeepSeek/Gemma  
3. Word is brought to the front  
4. The software screenshots the Word window  
5. YOLO detects the UI buttons inside the screenshot  
6. A transparent overlay highlights the detected feature  
7. Chat shows the AI-generated steps

Everything runs **locally** — no cloud processing.

---

## 📦 Installation

### **Prerequisites**
- Windows 10/11  
- Python 3.10+  
- Microsoft Word installed  
- Ollama installed → https://ollama.com/  

    
