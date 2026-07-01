# InAmigos Foundation Platform & Analytics Suite

A comprehensive, premium web portal and automated analytics reporting suite designed for the **InAmigos Foundation**. This repository bridges modern web development with data-driven advocacy, combining a high-performance interactive website with automated Python engines that generate research reports (PDF/DOCX) and compile national volunteer opportunity directories.

---

## 🌟 Features & Architecture

### 1. Interactive NGO Web Portal (`index.html`, `styles.css`)
A state-of-the-art, responsive landing page and volunteer hub featuring:
*   **Dual-Theme Engine:** Toggle between sleek Dark Mode and elegant Light Mode with a single click.
*   **Modern Fluid Design:** Leverages modern typography (Outfit & Inter), glassmorphism, accent glow blobs, and interactive CSS hover transitions.
*   **Scroll Progress Indicator:** A custom top-pinned bar that tracks reader progress in real-time.
*   **Comprehensive Sections:**
    *   *Hero & Welcome:* Direct Call-to-Action for donations and volunteering.
    *   *About Us:* Mission, Vision, and Core Pillars of the Foundation.
    *   *Focus Areas:* Showcases Education, Environmental Protection, and Health Support initiatives.
    *   *Impact Tracker:* Real-time statistical counters of achievements.
    *   *Responsive Media Gallery:* Grid structure displaying NGO activities.
    *   *In-App Modals:* Built-in overlay modal forms for direct sign-ups and donations.

### 2. Automated Opportunity Compiler (`generate_sheet.py`)
A custom Python script utilizing `openpyxl` that generates `Volunteer_Opportunities_India.xlsx`.
*   Compiles **17 premier volunteer programs, fellowships, and internships in India** (e.g., Teach For India, Gandhi Fellowship, NITI Aayog Internship, CRY, WWF).
*   **Premium Formatting:** Applied high-quality professional layout:
    *   *Navy Header Fill* (`#1A365D`) with white bold text.
    *   *Zebra-Striping* (`#F7FAFC`) for improved data readability.
    *   *Auto-Fitting Columns* to prevent text truncation.
    *   *Active Hyperlinks* styled with a clear teal accent.
    *   Visible gridlines and custom border styling.

### 3. Report & PDF Generation Suite (`generate_report.py`, `generate_pdf_report.py`, `generate_ai_growth_insights.py`)
Generates high-impact publications detailing the role of **Artificial Intelligence in NGO Operations**.
*   **Generated Documents:**
    *   `AI_NGO_Fundraising_Report.docx` / `AI_NGO_Fundraising_Report.pdf`: Analyzes sector adoption stats, predictive churn modeling, and donor trust mechanisms.
    *   `AI_NGO_Growth_Insights.docx` / `AI_NGO_Growth_Insights.pdf`: Outlines 5 core AI growth strategies for digital presence, hyper-personalized outreach, and efficiency.
*   **Branded Typography & Layouts:**
    *   Color palette: Navy (`#1A365D`) and Teal (`#319795`) styling.
    *   Executive summaries enclosed in custom styled callout boxes with solid left border accents.
    *   Automatic header and footer page indexing ("Page X of Y").

---

## 📂 Directory Structure

```text
inamigos-ngo/
│
├── assets/                    # Media assets and program icons
│   ├── gallery_activity.png
│   ├── gallery_campaign.png
│   ├── gallery_outreach.png
│   ├── gallery_volunteer.png
│   ├── hero_community.png
│   ├── project_education.png
│   ├── project_environment.png
│   └── project_health.png
│
├── index.html                 # Core web interface (HTML5)
├── styles.css                 # Custom CSS variables, responsive design, dark mode
│
├── generate_sheet.py          # Script generating Volunteer_Opportunities_India.xlsx
├── generate_report.py         # Script generating AI_NGO_Fundraising_Report.docx
├── generate_pdf_report.py     # Script generating AI_NGO_Fundraising_Report.pdf
├── generate_ai_growth_insights.py # Script generating AI Growth Insights (PDF & DOCX)
│
├── Volunteer_Opportunities_India.xlsx  # Compiled spreadsheet of fellowships
├── AI_NGO_Fundraising_Report.docx      # Fundraising DOCX report
├── AI_NGO_Fundraising_Report.pdf       # Fundraising PDF report
├── AI_NGO_Growth_Insights.docx         # Growth Insights DOCX report
├── AI_NGO_Growth_Insights.pdf          # Growth Insights PDF report
│
├── .gitignore                 # Standard Python/IDE ignore list
└── README.md                  # Project documentation (this file)
```

---

## 🚀 Setup & Execution

### 1. Web Portal
Since the interface is built with native HTML5 and Vanilla CSS, you do not need a compile step. Simply open the main file in your browser:
*   Double-click `index.html` or host it using a local server:
    ```bash
    python -m http.server 8000
    ```
    Then visit [http://localhost:8000](http://localhost:8000).

### 2. Python Report & Document Generation
To run the automated document compilers, you'll need Python 3.8+ and a few external formatting packages.

#### Installation:
```bash
pip install python-docx fpdf2 openpyxl
```
*(Note: `fpdf2` or `fpdf` is required depending on your local environment layout; the scripts are optimized to leverage basic fpdf features).*

#### Run Document Generators:
*   **Compile the Spreadsheet:**
    ```bash
    python generate_sheet.py
    ```
*   **Compile the Fundraising DOCX Report:**
    ```bash
    python generate_report.py
    ```
*   **Compile the Fundraising PDF Report:**
    ```bash
    python generate_pdf_report.py
    ```
*   **Compile the Growth Insights Suite:**
    ```bash
    python generate_ai_growth_insights.py
    ```

---

## 🎨 Theme Configuration & Customization
You can easily adjust the accent branding colors by modifying the CSS variables located at the top of `styles.css`:

```css
:root {
  --primary-color: #1a365d;   /* Main Navy Branding */
  --accent-color: #319795;    /* Teal Highlights */
  --bg-light: #ffffff;        /* Light Mode Background */
  --bg-dark: #121212;         /* Dark Mode Background */
}
```

---

## 📄 License
This repository is configured for the **InAmigos Foundation**. All rights reserved.
