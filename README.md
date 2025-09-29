# Yahoo Fantasy Hockey Draft Tool

A tool that automatically tracks your Yahoo Fantasy Hockey draft and creates an Excel spreadsheet with player data, rankings, and your custom projections.

## What It Does

- Downloads player information and draft rankings from Yahoo
- Creates an Excel file with multiple sheets for tracking players and projections
- Monitors your draft in real-time and records picks automatically
- Works on both Mac (with Apple Numbers support) and Windows

## Setup

### 1. Get Yahoo API Access

1. Go to https://developer.yahoo.com/ and sign in with your Yahoo account
2. Click "My Apps" → "Create an App"
3. Fill out the form:
   - **Application Name:** `Fantasy Draft Tool`
   - **Application Type:** `Web Application`
   - **Description:** `Personal fantasy hockey draft assistant`
   - **Home Page URL:** `http://localhost`
   - **Redirect URI(s):** `https://developers.google.com/oauthplayground`
   - **API Permissions:** Check `Fantasy Sports` (Read permission)
4. Click "Create App" and copy your **Client ID** and **Client Secret**

### 2. Find Your League ID

1. Go to your Yahoo Fantasy Hockey league
2. Look at the URL: `https://hockey.fantasysports.yahoo.com/hockey/12345/`
3. Copy the number (e.g., `12345`) - this is your League ID

### 3. Install Requirements

1. Install Python 3 from https://python.org/downloads
   - **Windows:** Check "Add Python to PATH" during installation
2. Download this project and open Terminal/Command Prompt in the project folder
3. Run: `pip install -r requirements.txt`

### 4. Configure Settings

1. Copy `.env.example` to `.env`
2. Open `.env` and fill in:
   ```
   YAHOO_CLIENT_ID=your_client_id_here
   YAHOO_CLIENT_SECRET=your_client_secret_here
   LEAGUE_ID=your_league_id_here
   ```

### 5. Initial Setup

Run: `python setup.py`

This will:
- Connect to Yahoo API (you'll authorize in your browser)
- Download league and player data
- Create your draft spreadsheet

## Using During Your Draft

1. **Before the draft starts:** Run `python draft_monitor.py`
2. **Open your Excel file** (default: `fantasy_draft_data.xlsx`)
3. **Leave the monitor running** - it checks for new picks every 10 seconds
4. **Watch picks appear automatically** in your spreadsheet

Press Ctrl+C to stop the monitor when your draft is finished.

## Your Spreadsheet

The Excel file contains these sheets:

- **Draft Board:** Main view showing all players, who drafted them, and your projections
- **Draft Results:** Log of every pick made during the draft
- **Skater/Goalie Projections:** Add your own player projections here
- **Teams:** All teams and managers in your league
- **League Settings:** Your league's scoring and roster settings
- **Pre-Draft Analysis:** Average draft position and rankings data

## Configuration Options

Optional settings in `.env`:

- `EXCEL_FILENAME=my_draft.xlsx` - Change the spreadsheet filename

## Troubleshooting

**"Command not found" or "Python not recognized"**
- Try `python3` instead of `python`
- Make sure Python is installed and in your system PATH

**"ModuleNotFoundError"**
- Run `pip install -r requirements.txt` (or `pip3 install -r requirements.txt`)

**Yahoo authentication issues**
- Double-check your Client ID and Client Secret in `.env`
- Make sure there are no extra spaces
- Try running `python setup.py` again

**Draft picks not updating**
- Make sure the monitor script is still running
- Check that your draft has actually started
- Verify your League ID is correct