# ⚙️ Siraal: Your AUTOCAD assitant

**The Deterministic AI-to-CAD Manufacturing Engine.**

Siraal is an enterprise-grade automation platform that bridges the gap between business data, Large Language Models, and deterministic CAD software. It enables mechanical engineers and procurement teams to generate complex 3D geometry from Excel BOMs or natural language, validate it against strict factory rules, and run real-time cost and ESG audits before the metal is ever purchased.

## 🌄 Demo

[](https://www.google.com/search?q=%5BYOUR_LINKEDIN_LINK_HERE%5D\(https://www.linkedin.com/\)) [](https://www.google.com/search?q=%5BYOUR_YOUTUBE_VIDEO_LINK_HERE%5D\(https://www.youtube.com/\))

*(Note: Replace placeholder image links with your actual Siraal UI screenshots)*

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

### Frontend

  * **ReactJS** + **Vite**: High-performance, modern UI rendering.
  * **Tailwind CSS**: For enterprise-grade, glassmorphism-inspired styling and dashboards.
  * **React Three Fiber (Three.js):** For lightweight, web-based 3D previews of the JSON CSG models.
  * **Axios**: For asynchronous API communication with the local CAD engine.

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
  * **Node.js:** Version 18+.

### 1\. Clone the Repository

```bash
git clone https://github.com/yourusername/siraal-engine.git
cd siraal-engine
```

## 📂 Project Structure

Here is an overview of the repository's architecture:

```text
siraal-engine/
├── siraal-ui/               # 🎨 Frontend Source Code (React + Vite)
│   ├── src/                 # React components, Dashboards, and 3D Web Viewer
│   ├── public/              # Static assets and icons
│   └── package.json         # Frontend dependencies (Tailwind, Three.js, etc.)
│
├── siraal-core/             # 🧠 Backend & CAD Infrastructure (Python)
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
for windows:
set METALPRICE_API_KEY=<YOUR API KEY>
set GEMINI_API_KEY=<YOUR API KEY> 
for mac/linux:
export METALPRICE_API_KEY=<YOUR API KEY>
export GEMINI_API_KEY=<YOUR API KEY> 

# Running the Headless Engine
python main.py
```

### 3\. Environment Setup (Frontend)

Siraal requires the ReactJS + Vite frontend dashboard to interact with the backend engine.

```bash
# Open a new terminal and navigate to frontend
cd siraal-ui

# Install Dependencies (Tailwind, React Three Fiber, Axios)
npm install

# Running the Dashboard
npm run dev
```
