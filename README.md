# SolveDesk

The main idea behind SolveDesk is to help users find solutions to previously encountered technical problems without having to search through multiple documents or solve the same problem repeatedly.

It maintains a knowledge base using Excel to store problem details and Word documents to store solutions. When an existing solution is unavailable, the application can generate an AI-based suggestion for the user to review and verify before adding it to the knowledge base.

## Demo##  Demo Video

[**Watch the Demo on Loom**](https://www.loom.com/share/92b08655eba843e7ade02a7dfe9c01b8)

## Screenshots

### Screenshot 1
![Application Screenshot 1](https://github.com/user-attachments/assets/1090f5b3-522e-4db9-9388-c415e0f3fec6)

### Screenshot 2
![Application Screenshot 2](https://github.com/user-attachments/assets/64ce9123-0ab9-44ba-b9ef-3ec56b15d033)

### Screenshot 3
![Application Screenshot 3](https://github.com/user-attachments/assets/7a074ceb-6c40-45ca-a68a-d84fae8663c2)


## What it does

- Searches saved problems by meaning when the embedding model is available; otherwise, it falls back to keyword overlap.
- Offers the Word document attached to a matching knowledge-base entry.
- Generates a draft with Ollama when no suitable saved answer is found.
- Lets you edit, retry, or discard a draft.
- Requires verification notes before a new solution can be saved with the `Verified` status.
- Saves new entries to Excel and creates a Word document in the `solutions` folder.
- Checks for possible duplicates when saving an AI-generated solution.
- Accepts optional PDF, DOCX, and TXT supporting documents.

This is a single-user local prototype. It does not prove that a person tested a fix; the `Verified` status records the user's confirmation and notes.


## Requirements

- Windows 10 or 11
- Python 3.10 or 3.11 recommended
- Ollama installed locally
- A compatible Ollama model downloaded
- Internet access for the first installation of Python packages and the sentence-transformer embedding model

Once the required packages and embedding model are cached and the Ollama model is available locally, the app can run without sending prompts to an external AI service.

## Setup (PowerShell)

Open PowerShell in the `SolveDesk` folder and run:

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

If `py -3.11` is not installed, create the environment with an available supported Python version instead.

Check that Ollama can load the configured model:

```powershell
ollama list
ollama run llama3.2:3b
```

If the model cannot start, fix that issue before testing AI generation in the app.

Start the app:

```powershell
streamlit run app.py
```

Open the local address printed in the terminal, usually `http://localhost:8501`.

## Tech stack
- Frontend: Streamlit
- Backend and Application Logic: Python
- Knowledge Base: Excel (.xlsx)
- Solution Documents: Microsoft Word (.docx)
- Semantic Search: Sentence Transformers model
- Fallback Search: Keyword-overlap matching
- Local LLM: Ollama with llama3.2:3b
- Document Processing: Python libraries for handling Excel, Word, PDF and TXT files

## How it works

- Problem Input: The user describes a technical problem in the application.
- Knowledge Base Search: SolveDesk searches existing problem records using semantic similarity when the embedding model is available, with keyword matching as a fallback.
- Solution Retrieval: If a suitable match is found, the application provides the associated solution document.
- AI-Assisted Resolution: If no suitable saved solution is found, Ollama generates a draft solution for the user to review.
- Verification: The user can edit, retry, or discard the generated draft. Verification notes are required before saving a solution with the Verified status.
- Knowledge Base Update: The problem record is saved to Excel, and its solution is stored as a Word document for future retrieval.

## Project files

- `app.py` — Streamlit interface, search, Ollama calls, and save workflow.
- `knowledge_base/knowledge_base.xlsx` — searchable problem records; this is the source of truth.
- `solutions/` — Word solution documents.
- `supporting_docs/` — optional documents uploaded through the app.

## Workbook columns

The workbook uses `problem_id`, `problem_title`, `problem_description`, `document_path`, `keywords`, `status`, `verification_notes`, and `created_at`.


## Future enhancements

- Replace the existing Streamlit interface with a React-based frontend to provide a more interactive, responsive, and user-friendly experience while retaining the existing backend functionality.
- Store and reuse the FAISS vector index to avoid rebuilding it for every search and improve search efficiency as the knowledge base grows.
- Implement backup and recovery mechanisms to protect Excel records and Word solution documents.
