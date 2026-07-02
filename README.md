# Gleim Study Session Auto-Scraper (Lead-and-Scrape)

An intelligent, robust web scraper built with Python, **Playwright**, and **openpyxl** to extract questions, answer options, correct choices, and detailed explanations from Gleim Study Sessions directly into a premium-designed Excel workbook.

---

## 🚀 The "Lead-and-Scrape" Workflow

Unlike traditional fully-automated scrapers that easily break due to CAPTCHAs, MFA, session timeouts, or complex dashboard changes, this scraper is built around a robust **hybrid workflow**:

1. **You Lead**: Run the script. A secure Chromium browser opens. You log in manually, set up your study session (e.g., choose units, set max questions, choose order), and navigate until you are looking at **Question 1**.
2. **It Scrapes**: Press `Enter` in your terminal. The scraper instantly scans all open browser tabs, detects the active study session window, and takes over to automatically click, extract, and document all questions.

This hybrid approach ensures **100% reliability**, handles all multi-tab navigation, and keeps your login credentials safe.

---

## 🔒 Security First

* **Zero Credentials Leaked**: Your login credentials can be saved locally in a `.env` file for auto-fill convenience, but the script supports complete manual login.
* **Ignored Files**: The `.gitignore` is pre-configured to ensure your `.env` configuration, logs, and scraped `.xlsx` spreadsheets are never pushed to GitHub.

---

## 🛠️ Features

* **Smart Tab Detection**: Automatically finds the active study session tab even if you open multiple pages or tabs during setup.
* **Premium Formatting**: Automatically duplicates the styling of `my_template.xlsx` (column widths, border structures, and navy header highlights) when generating a new workbook.
* **Duplicate Prevention**: Reads any existing workbook and automatically skips previously scraped questions based on their `(Subunit, Question ID)` keys. This allows you to scrape a large unit across multiple overlapping sessions without duplicating data.
* **HTML Table parsing**: Automatically converts HTML tables in questions and explanations to Markdown tables in Excel for readability.

---

## 📋 Setup & Usage

### 1. Install Dependencies
Make sure you have Python 3.8+ installed, then run:
```bash
# Install required Python packages
pip install -r requirements.txt

# Install Playwright browser binaries
playwright install chromium
```

### 2. Configure Environment (Optional)
Copy `.env.example` to `.env` and fill in your credentials to auto-fill the login form:
```bash
cp .env.example .env
```

### 3. Run the Scraper
Start the script:
```bash
python scraper.py
```

1. In the Chromium window that opens, log in and navigate to the dashboard.
2. Select your study units and set your question options (e.g., select max session size of 125).
3. Once you see **Question 1** on the page, return to your terminal and press **Enter**.
4. Sit back and watch the scraper extract the entire session!
