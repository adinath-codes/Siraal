# ⚙️ Siraal: Your AUTOCAD assitant

**The Deterministic AI-to-CAD Manufacturing Engine.**

Siraal is an enterprise-grade automation platform that bridges the gap between business data, Large Language Models, and deterministic CAD software. It enables mechanical engineers and procurement teams to generate complex 3D geometry from Excel BOMs or natural language, validate it against strict factory rules, and run real-time cost and ESG audits before the metal is ever purchased.

## 🌄 APP DEMO
**FULL DEMO VIDEOS**:
https://drive.google.com/file/d/16tRVIIonR_pIUBL50XCIIGI54h4Izpdr/view?usp=sharing

<img width="1919" height="1010" alt="Screenshot 2026-03-29 184759" src="https://github.com/user-attachments/assets/9569e144-5be2-4966-b7b0-82f9cc0038d6" />
<img width="1916" height="1025" alt="Screenshot 2026-03-29 184752" src="https://github.com/user-attachments/assets/9ed4db89-d1e5-4296-a3a3-47db49e8fb5e" />
<img width="1919" height="1018" alt="Screenshot 2026-03-29 184744" src="https://github.com/user-attachments/assets/f5e97b28-87a6-463a-b548-584f35b34e8c" />
<img width="1919" height="1010" alt="Screenshot 2026-03-29 184716" src="https://github.com/user-attachments/assets/f18b9503-c97a-47b8-83db-1b5a23225098" />
<img width="1919" height="1016" alt="Screenshot 2026-03-29 184703" src="https://github.com/user-attachments/assets/94e8f321-537d-4366-b5fa-a159f493de77" />
<img width="1919" height="1024" alt="Screenshot 2026-03-29 184807" src="https://github.com/user-attachments/assets/45292d2c-9fc2-45fb-9832-76af07d38e28" />

## 🌄 RESULTS SAMPLE
<img width="748" height="511" alt="Screenshot 2026-03-21 200240" src="https://github.com/user-attachments/assets/6978b483-de25-4806-bd53-5223a3654f61" />
<img width="998" height="616" alt="Screenshot 2026-03-14 065044" src="https://github.com/user-attachments/assets/9e5d0806-98d4-41f3-89d6-9e9ddcc63249" />
<img width="903" height="472" alt="Screenshot 2026-03-13 091107" src="https://github.com/user-attachments/assets/74751893-4a30-4706-97c4-a1761faeb2e2" />
<img width="1919" height="747" alt="Screenshot 2026-03-12 100840" src="https://github.com/user-attachments/assets/4e54d491-0dee-48d2-8bfa-ca605c1d7045" />

-----

## 🚀 Key Features

  * **📊 Excel-to-CAD Automation:** Read raw Bill of Materials (BOM) data and programmatically drive the CAD COM interface to generate batch geometry with 0% human transcription error.
  * **🧊 Patentable CSG JSON Architecture:** Decouples 3D intelligence from heavy B-Rep files (STEP/IGES). Siraal compiles lightweight, mathematical "executable algebraic blueprints" (JSON) that guarantee deterministic, crash-free CAD generation.
  * **💬 Text-to-CAD & AI Mass Edit:** Use natural language to design parts (e.g., "Design a heavy-duty NEMA motor mount") or mass-edit hundreds of parts instantly prior to rendering.
  * **💰 Autonomous AI Value Engineer:** Real-time cost estimation using live global metal pricing. Siraal calculates raw billet volume, CNC machining time, CO2 footprint, and flags terrible "Buy-to-Fly" waste ratios for immediate financial recovery.
  * **🛡️ Enterprise Safety & Compliance:** A digital gatekeeper with a real-time Verification Rules Engine. Blocks impossible physical geometry and logs all managerial overrides into an immutable Audit Trail for strict ISO-9001 compliance.
  * **🌐 The Siraal Library (B2B SaaS):** A cloud-based marketplace of mathematically perfect, AI-verified JSON manufacturing templates tailored for Tier-2 and Tier-3 contract manufacturers.

-----

## 🛠️ Tech Stack

### Backend & Manufacturing Infrastructure

  * **FastAPI**: Asynchronous API gateway acting as the "headless" orchestrator.
  * **Python `win32com.client`**: Direct algorithmic driver for the AutoCAD COM API (The core CAD execution layer).
  * **Gemini API**: Multimodal LLM engine utilized for Text-to-JSON mathematical translation and Value Engineering insights.
  * **Live Commodity APIs**: Real-time fetching of global steel, aluminum, and titanium market prices.

-----

## 💻 Local Installation Guide

Follow these steps to run Siraal locally on your machine.

### Prerequisites

  * **OS:** Windows 10 or Windows 11 (Mandatory for AutoCAD COM interface compatibility).
  * **CAD Software:** Autodesk AutoCAD installed and licensed on the host machine.
  * **Python:** Version 3.10 or higher.

### 1\. Clone the Repository

```bash
git clone https://github.com/yourusername/siraal-engine.git
cd siraal-engine
```
**HOW TO RUN**:
INSTALL AUTOCAD first
1) https://drive.google.com/file/d/1WT5WqvbVaZNYHPzeh439sUWJqf2F2jaM/view?usp=sharing

## 📂 Project Structure

Here is an overview of the repository's architecture:

```text
├── siraal/             # 🧠 Backend & CAD Infrastructure (Python)
│   ├── apis/                # Main FastAPI Application
│   │   ├── main.py          # API Gateway & Logic Orchestrator
│   │   ├── cad_driver.py    # Bridge to AutoCAD COM Interface (Critical)
│   │   ├── rules_engine.py  # ISO-9001 Validation & Safety Logic
│   │   └── requirements.txt # Python dependencies (FastAPI, pywin32, requests)
│   │
│   ├── templates/           # 🧊 CSG JSON Library
│   │   └── standard_parts/  # Executable algebraic blueprints (e.g., V8_Engine.json)
│   │
│   └── logs/                # 🛡️ Enterprise Audit Trails
│       └── audit_trail.log  # Immutable record of rule overrides
│
├── reports/                 # 📊 Generated AI Value Engineering PDFs
└── README.md                # Project Documentation
```

## ⚙️ Setup Guidelines

To run the Siraal Manufacturing Engine locally, ensure AutoCAD is installed and follow these steps:

### 1\. System Requirements

  * **OS:** Windows 10/11 (Required for `pywin32` COM interactions).
  * **RAM:** 16GB System RAM recommended for smooth CAD compilation.
  * **Software:** AutoCAD must be installed. It is recommended to have AutoCAD open before launching the backend.

### 2\. Environment Setup (Backend)

Siraal requires a Python environment capable of communicating with Windows COM objects.

```bash
# Navigate to backend
cd siraal

# Create & Activate Virtual Environment
python -m venv .venv
.venv\Scripts\activate

# Install Dependencies (FastAPI, pywin32, etc.)
cd apis
pip install -r requirements.txt

#Setting up the environemnetal variables
## for windows:
set METALPRICE_API_KEY=<YOUR API KEY>
set GEMINI_API_KEY=<YOUR API KEY> 
## or mac/linux:
export METALPRICE_API_KEY=<YOUR API KEY>
export GEMINI_API_KEY=<YOUR API KEY> 

# Running the Headless Engine
python main.py
```

