# Career_Suite_AI_v5
# Master Job Manager: An AI-Powered Job Analytics Platform **REVISED NAME**

[![Project Status: Active](https://img.shields.io/badge/status-active-success.svg)](https://github.com/your-username/your-repo-name)
[![Built with: Google Apps Script](https://img.shields.io/badge/Built_with-Google_Apps_Script-blue.svg)](https://www.google.com/script/start/)
[![Powered by: Gemini AI](https://img.shields.io/badge/Powered_by-Gemini_AI-purple.svg)](https://ai.google/gemini/)

An intelligent, dual-module system built within the Google Workspace ecosystem to automate the entire job search. It transforms a Google Sheet into a command center by intelligently parsing both active application updates and new job leads from your inbox, and culminates in a live, interactive analytics dashboard.

[<img src="(https://imgur.com/a/RsJD7Jw)" alt="Master Job Manager Dashboard Screenshot">](https://docs.google.com/spreadsheets/d/1QdqVZITapArTKGnBlpRGvAmpPPqCAvxng3SWYvNPZL0/edit#gid=0)
*(Click the image to see a live example of the dashboard)*

---

### 🎯 About The Project

The modern job search is fragmented across countless platforms, leading to a chaotic mess of emails and manual data entry. I built the Master Job Manager to solve this by creating a unified, "set it and forget it" system that automates the entire workflow.

This project is architected with a modular design, separating the logic for tracking active applications from the logic for capturing new job leads. It programmatically creates hierarchical labels and dedicated filters in Gmail to automatically route emails to the correct processing engine. The system's intelligence lies in its sophisticated AI and regex parsers, and its value is realized in the comprehensive, formula-driven dashboard that provides actionable insights on the job hunt.

---

### ✨ Key Features

#### 🏛️ Core System Architecture
*   **⚙️ Modular & Scalable Design:** Code is logically separated into distinct modules (`Main`, `Leads`, `Dashboard`, `GeminiService`, `ParsingUtils`, `Triggers`, `Leads_SheetUtils`, etc.) for superior organization, maintainability, and scalability.
*   **🧹 Automated Setup & Admin:** A custom Google Sheets menu allows for a **one-click full-system setup**, which creates all necessary spreadsheet tabs, hierarchical Gmail labels, and dedicated time-driven triggers for each module.
*   **🧩 Module-Specific Utilities:** Demonstrating a clean separation of concerns, the project utilizes both shared utility functions and module-specific helpers (e.g., `Leads_SheetUtils.gs`), ensuring that logic is encapsulated and reusable where appropriate.
*   **🔐 Secure API Key Management:** The Gemini API key is stored securely in `ScriptProperties`, not hardcoded.

#### 🧠 Sophisticated AI & Parsing Engine
*   **🤖 Advanced Prompt Engineering:** The system uses a highly detailed, multi-page prompt for the Google Gemini API, instructing it to act as a specialized data extractor. The prompt includes explicit instructions on handling edge cases, ignoring irrelevant emails, and adhering to a strict JSON output format.
*   **🔄 Resilient API Integration:** The `UrlFetchApp` calls to the Gemini API are wrapped in a robust error-handling and retry mechanism, which specifically manages rate-limiting (HTTP 429) and other transient network errors.
*   **🛡️ Hybrid Parser with Graceful Degradation:** For the application tracker, the system first attempts to parse with the Gemini API. If the API call fails, it **gracefully degrades** to a custom-built, multi-layered regex engine. This fallback engine meticulously deconstructs emails to ensure data quality.

#### 📊 Dynamic Analytics Dashboard
*   **📈 Formula-Driven & Efficient:** The dashboard is powered by dynamic array formulas (`QUERY`, `FILTER`, `COUNTIFS`) located in a hidden helper sheet. This efficient architecture ensures the dashboard updates instantly as new data arrives without requiring slow, script-based calculations.
*   ** KPIs:** The dashboard calculates and displays crucial job search metrics, including Total vs. Active Applications, Interview & Offer Rates, and a custom **Direct Reject Rate**.
*   **🖼️ Programmatic Chart Management:** The script automatically creates, modifies, and even removes charts (Pie, Line, Column) based on the availability of data in the helper sheet, ensuring the dashboard never displays broken or empty visuals.

#### ⚡ Robust Automation Engine
*   **🕰️ Centralized Trigger Management:** A dedicated `Triggers.gs` file manages the creation and verification of all time-driven automation.
*   **🎯 Purpose-Built Triggers:** The system sets up multiple, distinct triggers for different tasks: an hourly trigger for processing frequent application updates, and daily triggers for less frequent tasks like lead processing and stale application cleanup.

---

### 🛠️ Built With

*   **Core Logic:** Google Apps Script (JavaScript)
*   **Database/UI:** Google Sheets
*   **AI Service:** Google Gemini API (with Advanced Prompt Engineering)
*   **Email Integration:** Gmail API (Standard `GmailApp` & Advanced Gmail Service)
*   **Automation:** Google Apps Script Time-Driven Triggers

---

### 🚀 Getting Started

1.  **Create a new Google Sheet.**
2.  Open the Apps Script editor via `Extensions` > `Apps Script`.
3.  Create new script files and name them correctly (e.g., `Main.gs`, `Dashboard.gs`, `Leads_Main.gs`, `Triggers.gs`, etc.). Copy the code from this repository into the corresponding files.
4.  **Enable Advanced Services:** In the script editor, click the `+` next to "Services" and add the **Gmail API**.
5.  **Set up your Gemini API Key:** From the sheet menu, run **⚙️ Master Job Manager > Admin & Config > 🔑 Set Shared Gemini API Key**.
6.  **Run the Master Setup:** From the sheet menu, run **⚙️ Master Job Manager > ▶️ RUN FULL PROJECT SETUP** and authorize the script.
