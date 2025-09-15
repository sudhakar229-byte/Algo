# Daily Stock File Automation

This script automates the process of updating the daily stock Excel file. It handles data from multiple sources, including Google Sheets and other Excel files, to create a final, consolidated stock report.

## How It Works

The script performs the following steps:
1.  Determines the current date to work with the correct daily files.
2.  Connects to Google Sheets to fetch fresh "Arena" stock data.
3.  Reads the main stock file and replaces the old "Arena" data with the new data.
4.  Processes data from the dispatch report and adds it to the main stock file.
5.  Compares the DMS stock file and adds any new vehicles.
6.  Reads the sales register, finds all sold vehicles, and removes them from the main stock file.
7.  Saves the result as a new, timestamped Excel file, leaving all original files untouched.

## Setup Instructions (Phase 1)

This section will guide you through the one-time setup process required to run this script on your computer.

### 1. Install Python
- Go to the official Python website: [https://www.python.org/downloads/](https://www.python.org/downloads/)
- Download the latest version for your operating system (Windows or macOS).
- Run the installer. **Important:** On the first screen of the installer, make sure to check the box that says **"Add Python to PATH"** before you click "Install Now".

### 2. Set Up the Project Folder
- Create a new folder on your computer where you want to keep this project. For example, `C:\StockAutomation`.
- Download the files from this project (`main.py` and `requirements.txt`) and place them inside this new folder.

### 3. Install Required Libraries
- Open the command prompt or terminal on your computer.
  - On Windows, press the Windows key, type `cmd`, and press Enter.
- Navigate to your project folder using the `cd` command. For example:
  ```
  cd C:\StockAutomation
  ```
- Once you are in the correct folder, run the following command to install all the necessary libraries for the script. It will read the `requirements.txt` file and download them automatically.
  ```
  pip install -r requirements.txt
  ```

### 4. Google API Credentials

To allow the script to securely access your Google Sheets, you need to create a "Service Account" and give it permission. This is a one-time setup.

**Step 1: Go to the Google Cloud Console**
- Open your web browser and go to the [Google Cloud Console](https://console.cloud.google.com/). You may need to log in with your Google account.

**Step 2: Create a New Project**
- At the top of the page, click the project dropdown (it might say "Select a project").
- In the popup, click **"NEW PROJECT"**.
- Give your project a name, like "Stock Automation Script", and click **"CREATE"**.

**Step 3: Enable the Google Drive and Google Sheets APIs**
- In the search bar at the top, search for **"Google Drive API"** and press Enter.
- Click on the "Google Drive API" result and then click the **"ENABLE"** button.
- Go back to the search bar, search for **"Google Sheets API"**, and enable it as well.

**Step 4: Create a Service Account**
- In the search bar, search for **"Service Accounts"** and go to the Service Accounts page.
- Click **"+ CREATE SERVICE ACCOUNT"** at the top.
- Give the service account a name, like "stock-script-runner". The "Service account ID" will be filled in automatically.
- Click **"CREATE AND CONTINUE"**.
- For the "Role", select **"Project" -> "Viewer"**. This gives it read-only access, which is safer. Click **"CONTINUE"**.
- You can leave the last section blank. Click **"DONE"**.

**Step 5: Create and Download the Credentials Key**
- You should now see your new service account in the list. Click on the three dots under the "Actions" column and select **"Manage keys"**.
- Click **"ADD KEY" -> "Create new key"**.
- Choose **"JSON"** as the key type and click **"CREATE"**.
- A JSON file will be downloaded to your computer. This is your credentials file.

**Step 6: Prepare the Credentials File**
- Find the downloaded JSON file in your Downloads folder.
- **Rename the file to `credentials.json`**.
- Move this `credentials.json` file into the `stock_automation` folder, right next to the `main.py` script.

**Step 7: Share Your Google Sheets**
- Open your "arena updated stock" and "dispatch report converter" Google Sheets.
- Open the `credentials.json` file in a text editor. Find the line that says `"client_email"`. It will look something like `"client_email": "stock-script-runner@...iam.gserviceaccount.com"`.
- Copy that entire email address.
- In each Google Sheet, click the **"Share"** button (top right).
- Paste the service account's email address into the sharing dialog and give it **"Viewer"** access. **Do not give it "Editor" access.**
- Click **"Share"**.

That's it! Your script is now authorized to read the data from your Google Sheets.

## How to Run the Script

Once the setup is complete, you can run the script every day to automate your work.

1.  **Prepare Your Files:**
    *   Make sure you have downloaded all the required daily files into your project folder (`C:\StockAutomation` or wherever you created it). The script will look for files with today's date in the name:
        - `STOCK_AS_ON_DD-MM-YYYY.xlsx`
        - `DISPATCHES DD-MM-YYYY.html`
        - `DMS_STOCK_AS_ON_DD-MM-YYYY.xlsx`
        - `Wings stock as on DD-MM-YYYY.xlsx`
    *   Make sure your `credentials.json` file is also in this folder.

2.  **Run the Script:**
    *   Open the command prompt or terminal.
    *   Navigate to your project folder:
      ```
      cd C:\StockAutomation
      ```
    *   Run the script using the following command:
      ```
      python main.py
      ```

3.  **Get Your Output:**
    *   The script will print its progress in the window.
    *   When it is finished, a new file named `UPDATED_STOCK_DD-MM-YYYY.xlsx` will be created in your project folder. This file contains the final, processed data.
