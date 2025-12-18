# 🏠 Universal AI House Hunter (V5.2)

A professional, AI-powered real estate dashboard built on Google Sheets. This tool automates the process of scraping, translating, scoring, and organizing property listings from any website globally.

---

## 🌟 Key Features

* **Global Extraction:** Seamlessly scrape data from any real estate portal (Immobiliare.it, ImmobilienScout24, Zillow, etc.).
* **Anti-Block Technology:** Integrated with **Jina AI Reader** to bypass 403 Forbidden errors and CAPTCHAs.
* **Manual Fallback:** If a site is heavily protected, simply paste the text content; the AI handles the rest.
* **AI Scoring System:** Ranks every listing from 1-100 based on your personal requirements in the `Preferences` sheet.
* **Multilingual Output:** Automatically translates listings from any language into your chosen target language (Italian, Chinese, German, etc.).
* **Dual Inquiry Drafter:** Generates two outreach messages for landlords: one in the **local language** (for higher response rates) and one in your **target language** (for your understanding).
* **Interactive Dashboard:** High-density UI with clickable Google Maps links, status tracking, and Zebra-stripe formatting.
* **Auto-Archive:** Keep your workspace clean by moving "Bad" or "Applied" listings to an archive with one click.

---

## 🚀 Quick Start Guide

### 1. Setup the Spreadsheet

1. Open a new [Google Sheet](https://sheets.new).
2. Go to **Extensions > Apps Script**.
3. Paste the content of `Code.gs` into the editor and save.
4. Reload your Sheet. You will see a new menu: **🏠 AI House Hunter**.
5. Click **1. Initialize Sheets**.

### 2. Configure API Keys

1. **AI Provider:** Go to the menu > **3. API Management**. Set your **Gemini** or **OpenAI** API Key.
2. **Jina AI (Optional):** Get a free key at [jina.ai](https://jina.ai/reader/) to bypass website blocks. Set it via the menu.
3. **Language:** Go to **4. Set Output Language** and type your preferred language (e.g., "Italian").

### 3. Set Your Preferences

* Go to the `Preferences` sheet.
* Fill in your requirements (Budget, Location, Must-haves, Inquiry style). This data is used by the AI to calculate the **Score**.

---

## 🛠️ Usage

1. **Automated:** Paste a listing URL in the `Links` sheet (Column A).
2. **Manual:** If a site blocks scraping, copy the text from the ad and paste it into Column B (Double-click the cell first!).
3. **Run:** Click **5. Process All New Items**.
4. **Review:** Watch your `Database` fill with translated data, scores, and message drafts!

---

## 🧪 Troubleshooting

### Formula Parse Error (Maps)

This script uses the `;` separator for Google Maps formulas (standard for European/Italian locales).
If your Google Sheets uses the US locale (`,` separator), find this line in the code:
`const mapsHyperlink = "=HYPERLINK(\"" + mapsUrl + "\"; \"📍 View Map\")";`
And change the `;` to a `,`.

---

## 🛡️ License & Disclaimer

Distributed under the MIT License. This tool is for personal use. Ensure you comply with the Terms of Service of any website you interact with.
