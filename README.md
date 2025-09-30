# Yahoo Fantasy Hockey Draft Tool

A tool that automatically tracks your Yahoo Fantasy Hockey draft and creates a spreadsheet with player data, rankings, and your custom projections.

## What It Does

- Downloads player information and draft rankings from Yahoo
- Creates a spreadsheet with multiple sheets for tracking players and projections
- Monitors your draft in real-time and records picks automatically
- **macOS:** Creates native Apple Numbers documents with live formulas
- **Windows:** Creates Excel (.xlsx) files

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

**macOS:** Creates `fantasy_draft_data.numbers`
**Windows:** Creates `fantasy_draft_data.xlsx`

## Using During Your Draft

### macOS

1. **Open your Numbers file** (default: `fantasy_draft_data.numbers`)
2. **Keep the Draft Board sheet active**
3. **Run:** `bash run_monitor.command` (or double-click it in Finder)
4. **Watch picks appear automatically** in the Draft Results sheet in the background
5. Press Ctrl+C to stop when your draft is finished

The monitor updates the spreadsheet while it's open, without switching sheets!

### Windows

1. **Before the draft starts:** Run `python draft_monitor.py`
2. **Open your Excel file** (default: `fantasy_draft_data.xlsx`)
3. **Leave the monitor running** - it checks for new picks every 10 seconds
4. **Watch picks appear automatically** in your spreadsheet
5. Press Ctrl+C to stop when your draft is finished

## Your Spreadsheet

Your spreadsheet contains these sheets:

- **Draft Board:** Main view showing all players, who drafted them (by manager name), and projected points
- **Draft Results:** Log of every pick made during the draft (auto-populated by monitor)
- **Skater/Goalie Projections:** Add your own player projections here
- **Teams:** All teams and managers in your league
- **League Settings:** Your league's scoring and roster settings
- **Pre-Draft Analysis:** Average draft position and rankings data from Yahoo

## Configuration Options

Optional settings in `.env`:

- `FILENAME=my_draft.numbers` (macOS) or `FILENAME=my_draft.xlsx` (Windows) - Change the spreadsheet filename

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
