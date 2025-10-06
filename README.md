# THIS DOES NOT WORK YET.

- It's not tested on Windows
- The monitor does not fetch all the necessary info on macos. There were rows without playerKeys. Probably a simple fix, but that goes into to backlog until next draft season.

# Yahoo Fantasy Hockey Draft Assistant

This tool creates a draft spreadsheet for your Yahoo Fantasy Hockey league and automatically tracks picks during your draft. It downloads player rankings and statistics from Yahoo, then monitors your live draft and records each pick in real-time.

## What You'll Get

A spreadsheet with:

- All available players with their ADPs and positions
- Your own custom player projections
- Real-time updates as picks are made during your draft
- League settings and team rosters
- Historical draft data

On Mac, you'll get an Apple Numbers file. On Windows, you'll get an Excel file.

## Setup

### Step 1: Get Yahoo API Credentials

1. Visit https://developer.yahoo.com/ and sign in
2. Click "My Apps" → "Create an App"
3. Fill in the form:
   - **Application Name:** Fantasy Draft Tool
   - **Application Type:** Web Application
   - **Home Page URL:** http://localhost
   - **Redirect URI(s):** https://developers.google.com/oauthplayground
   - **API Permissions:** Check "Fantasy Sports"
4. After creating the app, copy your **Client ID** and **Client Secret**

### Step 2: Find Your League ID

1. Go to your Yahoo Fantasy Hockey league page
2. Look at the web address - it will look like: `https://hockey.fantasysports.yahoo.com/hockey/12345/`
3. The number (`12345` in this example) is your League ID

### Step 3: Configure the Tool

1. Copy the file `.env.example` and rename it to `.env`
2. Open `.env` in a text editor
3. Paste in your Client ID, Client Secret, and League ID

### Step 4: Run Setup

**On Mac:** Double-click the file `run_setup.command`

**On Windows:** Double-click the file `run_setup.bat`

A browser window will open asking you to authorize the app. After you approve, your spreadsheet will be created automatically.

## Using During Your Draft

### On Mac

1. Open the file `fantasy_draft_data.numbers`
2. Double-click `run_monitor.command` to start the draft monitor
3. As picks happen in your draft, they'll automatically appear in your spreadsheet
4. Press Ctrl+C when your draft is complete

### On Windows

1. Open Command Prompt and navigate to the tool folder
2. Run: `python draft_monitor.py`
3. Open the file `fantasy_draft_data.xlsx`
4. As picks happen in your draft, they'll automatically appear in your spreadsheet
5. Press Ctrl+C when your draft is complete

## Need Help?

**"Python not recognized" or "Command not found"**

- Make sure Python is installed from https://python.org/downloads
- On Windows, check "Add Python to PATH" during installation
- Try using `python3` instead of `python`

**"ModuleNotFoundError"**

- Run: `pip install -r requirements.txt`

**Draft picks aren't showing up**

- Confirm your League ID is correct in the `.env` file
- Make sure your draft has actually started
- Check that the monitor script is still running
