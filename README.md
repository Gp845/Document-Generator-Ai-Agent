# 📄 Document Generation Agent

An AI-powered document generation agent that creates professionally styled **PowerPoint presentations**, **Excel spreadsheets**, and **PDF reports** from natural language prompts. Built with a multi-model architecture: Gemini orchestrates intent, while secondary LLMs dynamically write the rendering code at runtime.

---

## ✨ Features

- **Natural language input** — describe what you want, the agent figures out the right document type and structure
- **Three output formats** — `.pptx`, `.xlsx`, `.pdf`
- **Dynamic formatting** — a Format Agent (Groq / OpenRouter) generates context-aware Python rendering code tailored to each topic's visual identity (finance, health, tech, etc.)
- **Pexels image integration** — automatically fetches relevant stock photos for slides and reports
- **Revision support** — refine content *or* formatting after the first generation without starting over
- **Standalone toolset** — `toolset.py` provides `generate_pptx_file`, `generate_xlsx_file`, and `generate_pdf_file` as independent functions you can call from any project

---

## 🏗️ Architecture

```
User prompt
    │
    ▼
Gemini Orchestrator (gemini-3.1-flash-lite)
    │  Decides tool + builds structured data
    ├──► create_presentation()  ──► Format Agent (Groq / LLaMA-3.3-70b)  ──► output_presentation.pptx
    ├──► create_spreadsheet()   ──► Format Agent (OpenRouter / Nemotron)  ──► output_data.xlsx
    └──► create_report()        ──► Format Agent (OpenRouter / Nemotron)  ──► output_report.pdf
    │
    └──► revise_document()  ──► reuses or regenerates Format Agent code
```

| Component | Model | Role |
|---|---|---|
| Orchestrator | `gemini-3.1-flash-lite` | Parses intent, selects tool, structures data |
| Format Agent (PPTX) | `llama-3.3-70b-versatile` via Groq | Generates Python rendering code for presentations |
| Format Agent (PDF/XLSX) | `nemotron-3-super-120b` via OpenRouter | Generates Python rendering code for reports & spreadsheets |
| Image Fetcher | Pexels API | Retrieves relevant stock images |

---

## 📦 Installation

```bash
pip install -r requirements.txt
```

---

## ⚙️ Configuration

The project requires four API keys. **Do not hardcode keys in source files** — use environment variables or a `.env` file instead:

```bash
export GOOGLE_API_KEY="your_google_api_key"
export PEXELS_API_KEY="your_pexels_api_key"
export OPENROUTER_API_KEY="your_openrouter_api_key"
export GROQ_API_KEY="your_groq_api_key"
```

| Key | Where to get it |
|---|---|
| `GOOGLE_API_KEY` | [Google AI Studio](https://aistudio.google.com) |
| `PEXELS_API_KEY` | [Pexels API](https://www.pexels.com/api/) |
| `OPENROUTER_API_KEY` | [OpenRouter](https://openrouter.ai) |
| `GROQ_API_KEY` | [Groq Console](https://console.groq.com) |

---

## 🚀 Usage

### Interactive agent (`main.py`)

```bash
python main.py
```

The agent runs an interactive chat loop. Example prompts:

```
You: make a 5-slide deck about renewable energy
You: create a spreadsheet comparing the top 10 programming languages by salary
You: write a report on the state of AI in healthcare
```

#### Revising a document

After generation, you can refine it in the same session:

```
You: make the accent colour green              # formatting change
You: add a slide about offshore wind           # content change
You: switch to a dark minimalist theme         # formatting change
You: update the revenue figure to $4.2M        # content change
```

### Standalone workers (`toolset.py`)

Use the three worker functions directly without the agent:

```python
from toolset import generate_pptx_file, generate_xlsx_file, generate_pdf_file

# PowerPoint
with open("deck.pptx", "wb") as f:
    f.write(generate_pptx_file({
        "title": "Q1 Strategy",
        "subtitle": "2026 Kickoff",
        "slides": [
            {
                "title": "Goals",
                "content": ["Grow revenue 20%", "Launch v2", "Expand to EU"],
                "accent_color": "#1E90FF",
                "icon_emoji": "🎯",
                "callout_stat": "20% growth = $2M additional ARR"
            }
        ]
    }).read())

# Excel
with open("report.xlsx", "wb") as f:
    f.write(generate_xlsx_file({
        "sheet_name": "Sales",
        "columns": ["Product", "Cost", "Quantity"],
        "rows": [["Widget A", 10.50, 100], ["Widget B", 25.00, 50]],
        "highlight_top_n": 2,
        "number_format_cols": ["Cost", "Quantity"],
        "freeze_header": True
    }).read())

# PDF
with open("report.pdf", "wb") as f:
    f.write(generate_pdf_file({
        "title": "Project Status",
        "sections": [
            {
                "heading": "Overview",
                "content": "The project is on track for Q2 delivery.",
                "callout": "80% complete — on schedule",
                "section_type": "highlight"
            }
        ]
    }).read())
```

---

## 📁 Output Files

| File | Description |
|---|---|
| `output_presentation.pptx` | Generated PowerPoint presentation |
| `output_data.xlsx` | Generated Excel spreadsheet |
| `output_report.pdf` | Generated PDF report |

---

## 📋 Data Schemas

### Presentation slide object
| Field | Type | Description |
|---|---|---|
| `title` | `str` | Slide heading |
| `content` | `list[str]` | 3–5 bullet points |
| `image_keyword` | `str` | 2–4 word Pexels search phrase |
| `accent_color` | `str` | Hex color e.g. `"#1E90FF"` |
| `icon_emoji` | `str` | Single emoji prefix |
| `callout_stat` | `str` | Bold highlight stat (≤ 12 words) |

### Report section object
| Field | Type | Description |
|---|---|---|
| `heading` | `str` | Section title |
| `content` | `str` | 3–5 sentence paragraph |
| `image_keyword` | `str` | 2–4 word Pexels search phrase |
| `callout` | `str` | Key stat or insight (≤ 15 words) |
| `section_type` | `str` | `"normal"` \| `"highlight"` \| `"warning"` \| `"tip"` |
| `table_data` | `list[list]` | 2D list; first row = headers |

### Spreadsheet payload
| Field | Type | Description |
|---|---|---|
| `sheet_name` | `str` | Worksheet tab name |
| `columns` | `list[str]` | Column headers |
| `rows` | `list[list]` | Data rows |
| `highlight_top_n` | `int` | Number of top rows to highlight |
| `number_format_cols` | `list[str]` | Column names to apply number formatting |
| `freeze_header` | `bool` | Whether to freeze the header row |

---

## 🔒 Security Note

The repository should **never** contain hardcoded API keys. Move all keys to environment variables before committing to version control. Consider adding a `.env` file (with `python-dotenv`) and adding `.env` to your `.gitignore`.
