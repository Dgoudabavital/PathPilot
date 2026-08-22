# PathPilot Advanced — AI Learning + Study Material Summarizer

## Features
- Personalized learning plans with 10-question quizzes per day topic.
- Immediate quiz score, correct answers, mastery status and progress tracking.
- Free / Moderate / Premium learning resource recommendations.
- User-entered course or skill names.
- Planned learning PDF download (JSON download removed).
- Contextual local study chatbot.
- **Upload PDF or image → extract text → source-grounded simple summary → simple explanation.**
- Previous uploaded materials are saved per user in SQLite.

## Run in VS Code
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```
Open http://127.0.0.1:5000

## Image OCR on Windows
Image summarization uses Tesseract OCR through `pytesseract`. Install the Tesseract OCR application separately and PathPilot can also find the common Windows install path automatically; if needed, set `TESSERACT_CMD` to the full path of `tesseract.exe`. PDF text extraction does not need Tesseract. If you only need PDF summarization, the PDF feature works without OCR.


## OpenCV image preprocessing

This version includes `opencv_processor.py`. Uploaded study images can be
preprocessed before OCR using grayscale conversion, CLAHE contrast enhancement,
denoising, adaptive thresholding, morphology, resizing, and deskewing.

Install dependencies with:

```bash
pip install -r requirements.txt
```

A test API endpoint is available at `POST /api/opencv-process` with form field
`image`. The frontend can call this endpoint before sending an image to OCR.


## Improved study-material understanding

Uploaded PDFs/images are now processed with a source-grounded pipeline. Text PDFs use direct extraction; scanned PDFs and images use PyMuPDF/OpenCV preprocessing + Tesseract OCR. The result page gives: what the material is, a beginner-friendly explanation, main ideas, key terms, quick revision points, and a source-grounded summary. The explanation deliberately avoids adding unsupported outside facts.

## Latest fixes: upload preview + tutor chatbot

- Selected PDF/image is previewed before submission.
- The result page displays the original uploaded image or an embedded PDF viewer.
- `+ Add another` lets the learner immediately upload another material.
- A secure `/materials/file/<id>` route serves the user's own uploaded file.
- The tutor chatbot now handles Enter-to-send, loading/error states, progress, quiz, resources, current topic, motivation, and source-grounded questions about recent uploaded material.
- Image OCR automatically checks common Windows Tesseract installation paths.

### Windows image OCR
Install Tesseract OCR separately if it is not installed. If it is installed somewhere else, set `TESSERACT_CMD` to the full path of `tesseract.exe` before starting Flask.

## Latest fixes
- Study Material uploader now supports multiple PDF/image files in one submission.
- Users can click **＋ Add another file** before or after processing.
- Selected images show thumbnails and selected PDFs show file cards before upload.
- The original uploaded image/PDF is displayed on the result page.
- Chatbot now has an offline fallback for common questions such as “What is Python?”, JavaScript, HTML, CSS, SQL, React, Flask, APIs, Machine Learning and OpenCV, while still using uploaded material when relevant.
- Upload file serving is configured for browser preview.


## AI tutor and document understanding

This version supports PDF, images, TXT, Markdown, CSV, DOCX and PPTX study materials. Images and scanned PDFs use OpenCV preprocessing + Tesseract OCR.

For a true general-purpose AI tutor and higher-quality source-grounded summaries, set these optional environment variables:

```text
OPENAI_API_KEY=your_key_here
OPENAI_MODEL=gpt-4o-mini
```

Without an AI API key, the site still works with offline extraction, summarization, quizzes, ML predictions and local topic explanations, but it cannot honestly provide unlimited general-world chatbot answers.


## General-purpose AI Tutor
The chatbot is a true general-purpose AI tutor when `OPENAI_API_KEY`, `GEMINI_API_KEY`, or optional local Ollama is configured. It sends the exact user question to the selected model, keeps recent conversation history, and adds learning-plan and uploaded-material context when relevant. If the provider is configured but fails, the UI shows the real provider error instead of silently falling back to a tiny hard-coded answer list.

### Local Ollama option
If you want the chatbot to run without a cloud API key, install Ollama and pull a model such as `llama3.2:3b`, then set:

```text
OLLAMA_ENABLED=true
OLLAMA_URL=http://127.0.0.1:11434
OLLAMA_MODEL=llama3.2:3b
```

The model must be installed and running locally.

### Gemini option
1. Create a Gemini API key.
2. Put it in `.env` as `GEMINI_API_KEY=...`.
3. Restart `python app.py`.
No extra Gemini Python package is required; the app calls the REST API with `requests`.


## General AI chatbot
The chatbot is a real LLM client. Configure Gemini, OpenAI, or local Ollama in `.env`. It no longer uses the old keyword-definition dictionary as a general-answer fallback. If no model is configured, it clearly reports that configuration is required.

### Gemini
Set `GEMINI_API_KEY` and restart the app.

### OpenAI
Set `OPENAI_API_KEY` and restart the app.

### Local Ollama
Install Ollama, download a model, set `OLLAMA_ENABLED=true`, and restart.
# PathPilot - Final Render Deployment

## Render
1. Upload/push this entire folder to the GitHub repository used by Render.
2. Make sure `app.py`, `ml_engine.py`, `opencv_processor.py`, `templates/`, `static/`, and `requirements.txt` are in the repository root.
3. Render build command: `pip install -r requirements.txt`
4. Render start command: `gunicorn --bind 0.0.0.0:$PORT app:app`
5. Deploy.

The application initializes the SQLite schema when Gunicorn imports `app.py`, so the `users` table is created before registration/login requests.

Do not upload a local `data/learning.db` file. The app creates it automatically.
