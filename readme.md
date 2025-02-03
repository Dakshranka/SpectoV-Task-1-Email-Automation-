# Google Sheets API Trigger Script

This Google Apps Script listens for edits made in a Google Sheet and triggers an HTTP request to a specified API endpoint with data from the edited row.

## Features

- Triggers on edits in a specific sheet (`Sheet1`).
- Sends a POST request to a custom Flask API with data from the edited row.
- Only processes rows below the header (ignores header row edits).
- Logs API responses and errors for debugging.

## Requirements

- A Google Sheets document with at least two columns:
  - Column 1: `Name`
  - Column 2: `Email Address`
- A Flask API (or similar HTTP server) capable of accepting POST requests with JSON payloads at a specified URL.
- **Ngrok** (if you're running the Flask API locally for testing purposes).

## Setup

### 1. **Create a New Google Sheets Script:**
   - Open your Google Sheets document.
   - Go to `Extensions` → `Apps Script`.
   - Delete any default code in the editor and paste the provided script.

### 2. **Deploy the Script:**
   - Save the script in the Apps Script editor.
   - Go to `Triggers` (clock icon) in the Apps Script editor and set up a trigger to run the function `myFunction` on the event `On edit`.

### 3. **Install Ngrok (For Local Development):**
   If you're running your Flask API locally and want to expose it to the internet for testing, you can use Ngrok to create a tunnel. Follow these steps:
   
   - [Download Ngrok](https://ngrok.com/download) and install it on your machine.
   - After installation, start your Flask API on your local machine (e.g., `python app.py`).
   - In another terminal window, run Ngrok to tunnel your local Flask server:
     ```bash
     ngrok http 5000
     ```
     This will give you a public URL, such as `https://1234abcd.ngrok.io`, that you can use for testing.

### 4. **Edit the API URL:**
   - Replace the example URL (`https://14f6-122-187-117-179.ngrok-free.app/trigger-email`) in the script with the Ngrok URL (or your production URL if not using Ngrok).

### 5. **API Setup:**
   - Ensure that your Flask API is capable of handling POST requests at the `/trigger-email` endpoint.
   - The API should expect a JSON payload with the following structure:
     ```json
     {
       "Name": "John Doe",
       "Email Address": "john.doe@example.com"
     }
     ```

## Script Details

### Function: `myFunction(e)`

- **Trigger:** This function is triggered by an `onEdit` event in the Google Sheets document.
- **Validation:**
  - The function ensures that the edit occurs in `Sheet1` and that it is not a header row.
- **API Call:** 
  - Sends a POST request to the configured API endpoint with `Name` and `Email Address` data from the edited row.
- **Logging:** 
  - Logs success or error responses to the console for debugging.

### Example Payload:
When a row is edited in the sheet, the script sends a POST request with the following JSON payload:
```json
{
  "Name": "John Doe",
  "Email Address": "john.doe@example.com"
}
