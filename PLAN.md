# PLAN: PDF → AI → Excel Tool

## 1. What we're building
A small **desktop app** (a window on your screen) named **"Paper Extractor"**.
You drag PDF research papers into it, click one button, and it creates a **single Excel file** with one row per paper containing:

| Title | DOI | Abstract | Method | Research gap | Key findings | Limitations |

Plus useful helper columns: **File name**, **Year** (guessed), **Status** (ok / scanned / failed).

## 2. Is an interface needed?
**Yes.** You chose a desktop app, so there will be a simple window you open and click in. No internet hosting needed — it runs on your own computer.

## 3. How it works (the flow)

```
You drag PDFs in
      │
      ▼
 Desktop window
      │  1. reads the text out of each PDF  (using "pdfplumber" library)
      │
      ├──► Text too short (scanned/image PDF)?
      │         └──► Excel row marked "Scanned – needs manual review", skip AI
      │
      ▼  2. sends the text to Google AI (Gemini, FREE tier)
      │         prompt: "Find Title, DOI, Abstract, Method, Research gap,
      │                  Key findings, Limitations. Answer as JSON."
      ▼  3. AI answer comes back, app fills in the row
      │
      ▼  4. you review on screen, click "Export to Excel"
      ▼
 Excel file (.xlsx)  ←  created by "openpyxl" library
```

## 4. Parts/tools we need (and cost)

| Part | What it is | Cost |
|---|---|---|
| **Python** | The engine that runs everything | Free |
| **pdfplumber** library | Extracts text from PDFs | Free |
| **openpyxl** library | Creates Excel files | Free |
| **google-generativeai** library | Talks to Gemini | Free |
| **Google Gemini** | The AI brain | **Free tier** (free API key, limited to ~a few thousand requests/day — plenty for 1–10 PDFs) |
| The app code | Written by the assistant | — |

**Total cost: $0.** (Only manual step for you: getting the free Google key, ~5 minutes.)

## 5. What YOU need to do (start to finish)

1. **Install Python** — 1-click installer from python.org (tick "Add Python to PATH").
2. **Get the free Gemini key** — go to `https://aistudio.google.com/apikey`, click "Create API key", copy it, paste it into the app's settings box. (No credit card needed.)
3. **Open the app** — double-click it.
4. **Drag PDFs / click "Add"** — 1 to 10 at a time.
5. **Preview results** → click **"Export to Excel"**.
6. Save the Excel file anywhere. Done.

## 6. Build steps (implementation)

- **Phase 1 – App window:** drag-drop area, list of files, Add/Remove/Export buttons (Python + "customtkinter" for a clean modern look).
- **Phase 2 – PDF reading:** open each PDF, pull the text, detect scanned/unreadable files (text near zero) and flag them.
- **Phase 3 – AI connection:** send text (trimmed if a paper is very long) to Gemini, ask it ONLY for the 7 fields, returned as structured JSON so it's reliable.
- **Phase 4 – Safe parsing:** handle missing DOI, wrong JSON, AI refusals — never crash the app; mark problem rows as "Check manually".
- **Phase 5 – Excel export:** neat headers, auto column widths, one row per paper, freeze top row, status colors.
- **Phase 6 – Settings:** box to paste your API key, and a small delay (1 sec) between papers so the free tier never blocks you.

## 7. Practical expectations

- **No internet limits:** 10 PDFs ≈ takes ~1–2 minutes total.
- **Very long papers:** app sends the beginning (≈ first ~14,000 words) so it fits the free AI limit — abstracts/intro/method always come first, so results stay good.
- **Scanned/image PDFs:** skipped automatically, marked in Excel as "Scanned – review later".
- **Different layouts:** Gemini handles different paper formats well (IEEE, Springer, MDPI, etc.).

## 8. If you ever want to learn to modify it yourself

- **Week 1:** Python basics (variables, functions, running a script) — free at `freecodecamp.org` or `python.org`.
- **Week 2:** A simple GUI tutorial (search "Python tkinter first app"); then read how `main.py` calls the AI (2–3 functions).
- **Week 3:** Play with the AI prompt in the Settings box — you can change what fields get extracted without touching code.

You **don't need to learn any of this to use the tool** — the app will be ready to double-click.

## 9. AI model choice

- Recommended: **Gemini flash model** (gemini-2.0-flash) — free tier, fast, excellent at extraction.
- Alternative: newest Gemini model (still free tier) if you want absolute highest quality even if slightly slower.