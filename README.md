# Web Scraper with Secure Login and Excel Integration

This repository contains a secure, locally runnable web scraper built using Python, **Playwright**, and **openpyxl**. It logs into a password-protected website, scrapes data paginated across multiple pages, and writes or appends the data directly into your pre-designed Excel workbook.

---

## 🔒 Security Best Practice

**Never share your password or username in the AI chat window or check them into version control.**

Your credentials are loaded securely at runtime on your local machine using a `.env` file, which is listed in `.gitignore` so it is never shared or uploaded.

---

## 🚀 Setup Instructions

Follow these steps to set up and run the scraper on your machine.

### 1. Install Dependencies

Ensure you have Python 3.8+ installed on your Mac. Open your terminal in this directory and run:

```bash
# Install Python packages
pip install -r requirements.txt

# Install Playwright browser binaries
playwright install chromium
```

### 2. Configure Environment Variables

1. Copy the `.env.example` file to a new file named `.env`:
   ```bash
   cp .env.example .env
   ```
2. Open `.env` in a text editor and fill in the values:
   - `SCRAPER_USERNAME` and `SCRAPER_PASSWORD` with your actual login credentials.
   - `LOGIN_URL` with the URL of the login form.
   - `SELECTOR_...` selectors for locating elements on the page (see below on how to find them).
   - `EXCEL_TEMPLATE_NAME` with the name of your Excel file (e.g., `template.xlsx`). Make sure you put the Excel file in this directory!

### 3. Customize Selectors (Optional)

If the website uses different elements, open your browser, right-click on the inputs/buttons, choose **Inspect**, and find their CSS Selectors (e.g. `#username`, `.email-input`, `button[type="submit"]`). Update these values in your `.env` file.

### 4. Customize Scraped Columns (Optional)

Open `scraper.py` and modify the `parse_row_data` function if you need to extract specific elements (like cell values from a table or text from cards). By default, it parses standard HTML tables or lists.

### 5. Run the Scraper

To run the scraper:
```bash
python scraper.py
```

A visible Chromium browser window will open, and you can watch the script fill in the credentials, log in, navigate the pages, extract the data, and append it to your Excel file.
