# 📑 PaperExtract AI — Scientific PDF to Structured Excel Research Extractor

> **Automated Literature Extraction Studio** • Transform research paper PDFs into deep, structured Excel datasets with AI-powered entity recognition.

Drag your PDFs in, click **Analyze**, and the app reads each paper and extracts:

| Title | DOI | Year | Abstract | Method | Research gap | Key findings | Limitations |

plus a **File name** and a **Status** column so you always know every row's health at a glance.

---

## ✨ Features

- ✅ **Drag & drop** PDFs (or use the **Add PDFs** button)
- ✅ **AI-powered extraction** of 8 fields per paper using **free OpenRouter models**
- ✅ **Batch processing** — 1 to 10+ papers in one go
- ✅ **Smart status flags**:
  - 🟢 **Done** — successfully analyzed
  - 🟡 **Scanned – review manually** — image-only PDF (no text), skipped automatically
  - 🔴 **AI error** / **Failed** — could not be read or processed
- ✅ **Automatic retry** when free models are busy (avoids rate-limit failures)
- ✅ **Clean Excel export** — styled headers, wrap text, frozen top row, color-coded status
- ✅ **Your AI key stays private** — stored only on your machine, never in the project

---

## 🧠 How it works

```
You drag PDFs in
        │
        ▼
────────────── ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─
  Paper Extractor (runs on your PC)
  1. Reads the text out of each PDF  (pdfplumber)
  2. Detects scanned/unreadable files → marks them, skips AI
  3. Sends the paper text to a free AI model (OpenRouter)
  4. AI returns the fields as clean JSON
  5. Shows progress in the window
────────────── ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─
        │
        ▼
  One Excel file (.xlsx) with one row per paper
```

---

## 🛠 Requirements

- **Windows 10 / 11**
- **Python 3.10+** — download from [python.org](https://www.python.org/downloads/) and tick **"Add Python to PATH"** during install

---

## ⚙️ Installation

### Option A — one click (recommended)

1. Double-click **`Start App.bat`**
2. It installs everything automatically, then opens the app

### Option B — manual

```bash
pip install -r requirements.txt
python app.py
```

---

## 🔑 First-time setup: get a free AI key

1. Go to **[openrouter.ai/settings/keys](https://openrouter.ai/settings/keys)**
2. Create a free account (if needed), then click **Create Key**
3. Copy the key (`sk-or-v1-...`)
4. In the app, paste it into the **OpenRouter key** box
5. Choose an **AI model** (see below)
6. Your key is saved securely on **your PC only**

> 🔒 **Privacy:** Your key is stored in a private folder on your computer and is **never** committed to the repository, uploaded, or shared.

---

## 🚀 Usage

1. **Open the app** — double-click `Start App.bat`
2. **Add PDFs** — drag and drop them into the window, or click **+ Add PDFs**
3. Click **Analyze (AI)** — watch each paper get processed live
4. Once finished, click **Export to Excel**
5. Choose a save location — the file opens automatically

> **Tip:** One batch takes roughly 1–2 minutes for 10 papers.

---

## 📊 Excel output

One row per paper, with these columns:

| Column | Description |
|---|---|
| **File name** | Name of the PDF |
| **Title** | Title of the paper |
| **DOI** | DOI link (if present) |
| **Year** | Publication year (if shown) |
| **Abstract** | 1–3 sentence summary |
| **Method** | How the study was conducted |
| **Research gap** | The problem the paper addresses |
| **Key findings** | 3–6 main results |
| **Limitations** | 2–4 limitations |
| **Status** | Done / Scanned / AI error / Failed (color-coded) |

---

## 🤖 AI models included

All are **free** tiers on OpenRouter (statuses change over time):

- `nvidia/nemotron-3-ultra-550b-a55b:free` — large, high-quality, robust JSON output *(default)*
- `google/gemma-4-31b-it:free`
- `google/gemma-4-26b-a4b-it:free`
- `openai/gpt-oss-20b:free`

> 💡 If the free tier is busy, the app waits and retries automatically. You can switch models anytime from the dropdown.

---

## 🩺 Troubleshooting

| Problem | Solution |
|---|---|
| "AI error – check manually" | Free tier may be busy; wait a minute and re-run, or switch model |
| "Scanned – review manually" | The PDF is an image scan with no selectable text — review it by hand |
| App won't open | Make sure Python 3.10+ is installed and "Add Python to PATH" was ticked |
| Key not working | Check it starts with `sk-or-v1-` and has no extra spaces |

---

## 📁 Project structure

```
automation tool/
├── Start App.bat      # One-click launcher (installs + opens the app)
├── app.py             # The full application (UI + logic)
├── PLAN.md            # Original planning document
├── requirements.txt   # Python dependencies
└── .gitignore         # Excludes secrets, cache, and temp files
```

---

## 🧪 Testing

The app was tested for:
- PDF text extraction
- Excel export correctness
- Real AI calls with free OpenRouter models
- Drag & drop and window behavior

---

## 📄 License

You're free to use, modify, and share this project for personal and educational use.
