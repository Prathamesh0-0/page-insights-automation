# Website Performance Report Automation

## 🚀 Overview

This Python script is a powerful automation tool designed to streamline the process of lead outreach by generating and sending personalized website performance reports. It reads a list of leads from an Excel file, fetches real-time data from the Google PageSpeed Insights API, and automatically creates and emails a detailed PDF report for each website.

## ✨ Features

* **Automated Data Extraction**: Fetches performance scores, Core Web Vitals, and other key metrics using the Google PageSpeed Insights API.

* **Graphical PDF Generation**: Uses Playwright to navigate to the PageSpeed report URL and generate a high-fidelity PDF of the full report.

* **Personalized Emailing**: Sends a custom, branded email with the PDF report attached to each lead.

* **Robust Error Handling**: Skips invalid leads, retries failed API calls, and logs all activity for easy debugging.

* **Scalable**: Designed to process a large number of leads efficiently.

## 🛠️ Prerequisites

Before you run the script, ensure you have the following installed:

* **Python 3.x**

* The required Python libraries. You can install them using `pip`:

    ```bash
    pip install -r requirements.txt
    ```

* A Google PageSpeed Insights API key.

## ⚙️ Setup

1.  **Clone the repository:**

    ```bash
    git clone https://github.com/Prathamesh0-0/page-insights-automation.git
    cd page-insights-automation
    ```

2.  **Create your configuration:**

    * Fill in your **Google PageSpeed API Key**, **SMTP server details**, and **email credentials** in the script.

3.  **Prepare your leads:**

    * Create an Excel file named `leads.xlsx` with columns for `name`, `website`, and `email`.

    * The script will automatically try to find the best-matching columns, but using these exact names is recommended.

## 🏃 Usage

Simply run the script from your terminal:

```bash
python send_reports.py
