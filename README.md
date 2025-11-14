📘 URA AI
Your personal Microsoft Word & Excel assistant — powered by YOLO + Gemma.
⭐ Overview

URA is an intelligent desktop assistant that helps you perform tasks in Microsoft Word and Excel instantly.

Just ask:
“How do I bold text?”
“Insert a table.”
“Add a chart.”

The AI detects toolbar buttons using YOLO, understands your query using LLMs ( Gemma ), and shows step-by-step instructions — along with an on-screen highlight pointing to the exact feature.

🚀 Features

🧠 Natural Language Understanding
Uses Gemma:2b (via Ollama) to understand your request.

👁️ Real-time UI Detection
YOLOv11 detects Word/Excel toolbar icons on your screen.

🖼️ Smart Screen Capture
Automatically captures your screen without capturing the chatbot window.

🔍 Live Overlay Highlight
Draws a bounding box over the exact button the user needs.

✔️ Automatic Step Generation
AI returns clean JSON steps like:

{ "intent": "bold", "label": "icon_bold", "steps": ["Home → Font → Bold"] }


📌 Works for Word and Excel
Supports 140+ features (bold, italic, underline, tables, borders, charts, formulas, spelling, etc.)

🔌 Offline Support
Works locally using Ollama models.
