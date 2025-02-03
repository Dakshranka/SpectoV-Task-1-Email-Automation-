# Internship Offer Letter Automation

This Google Apps Script automates sending internship offer letters from a Google Sheet. It reads names and emails from the sheet and sends personalized emails using Gmail.

## 🚀 Features
- Sends personalized internship offer emails.
- Prevents duplicate emails by tracking the `Status` column.
- Logs and alerts errors (missing sheet, headers, or failed emails).
- Adds a custom menu (**Send Mail**) in Google Sheets for easy access.

## 📌 Setup Instructions

### 1️⃣ **Prepare Google Sheets**
1. Create a Google Sheet and name it **`Automation`**.
2. Add the following column headers in **Row 1**:
   - `Name` (Recipient's Name)
   - `Email` (Recipient's Email Address)
   - `Status` (Leave this blank initially)

### 2️⃣ **Add Apps Script**
1. Open the Google Sheet.
2. Click **Extensions > Apps Script**.
3. Delete any existing code and paste the **automation.gs** script.
4. Click **Save** (💾) and **Run** the `onOpen` function to activate the menu.

### 3️⃣ **Grant Permissions**
1. When running the script for the first time, it will ask for Gmail and Spreadsheet permissions.
2. Click **Authorize**.

### 4️⃣ **Run the Script**
1. Go to **Google Sheets**.
2. Click **Send Mail** (from the menu) > **Send Emails**.
3. The script will:
   - Read the email list.
   - Send emails to recipients without a "Status".
   - Update `Status` to **"Email Sent"**.

## 📜 Logging & Debugging
- **View Logs**: Open Apps Script editor > Click `View` > `Logs`.
- If emails fail, `Status` will show `"Failed to Send"`.

## ⚠️ Email Quotas
Google limits emails per day:
- **Free Gmail Account**: 100 emails/day.
- **Google Workspace (Paid)**: Up to 1,500 emails/day.

## 🎯 Use Cases
- Sending internship/job offer letters.
- Automated email notifications.
- Bulk email communication from Google Sheets.

---

### ✨ Contributing
Feel free to modify the script to fit your needs! 😊

---
Made with ❤️ by SpectoV Team.
