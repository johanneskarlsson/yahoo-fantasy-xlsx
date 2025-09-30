#!/usr/bin/env python3

import os
import logging
from dotenv import load_dotenv

from yahoo_api import YahooFantasyAPI
import platform
if platform.system() == 'Darwin':
    from macos.numbers_export import MacOSDraftExporter as DraftExporter  # type: ignore
else:
    from windows.xlsx_export import XlsxDraftExporter as DraftExporter  # type: ignore
"""Setup script: builds canonical XLSX; on macOS a .numbers companion is auto-created by exporter."""

# Load environment variables
load_dotenv()

# Configure logging (reduce verbosity for setup)
logging.basicConfig(
    level=logging.WARNING,
    format='%(levelname)s: %(message)s'
)

def setup_yahoo_authentication():
    """Setup Yahoo API authentication"""
    print("\n=== Yahoo API Authentication ===")
    api = YahooFantasyAPI()

    try:
        # Try to load existing token
        if api.load_token() and api.refresh_token_if_needed():
            print("✓ Yahoo authentication already configured")
            return True
        else:
            print("Yahoo authentication required...")
            api.authenticate()
            print("✓ Yahoo authentication completed")
            return True
    except Exception as e:
        print(f"✗ Yahoo authentication failed: {e}")
        return False

def initialize_data():
    """Initialize the Excel/Numbers file with required data (slim mode if exporter lacks advanced methods)."""
    try:
        yahoo_api = YahooFantasyAPI()

        # Decide filenames and whether to auto-convert to Numbers after XLSX generation
        requested_filename = os.getenv('FILENAME', 'fantasy_draft_data.xlsx')

        # Normalize to an .xlsx canonical filename
        if requested_filename.lower().endswith('.numbers'):
            base = requested_filename[:-8]  # strip .numbers
            if not base:
                base = 'fantasy_draft_data'
            xlsx_filename = base + '.xlsx'
            numbers_filename = base + '.numbers'
        elif requested_filename.lower().endswith('.xlsx'):
            xlsx_filename = requested_filename
            numbers_filename = requested_filename[:-5] + '.numbers'
        else:
            # Add .xlsx if no recognized extension
            xlsx_filename = requested_filename + '.xlsx'
            numbers_filename = requested_filename + '.numbers'

        numbers_exporter = DraftExporter(xlsx_filename)
        if platform.system() == 'Darwin':
            print(f"Creating Numbers file: {numbers_filename}")
        else:
            print(f"Creating Excel file: {xlsx_filename}")

        # Ensure authentication
        yahoo_api.ensure_authenticated()

        # Predeclare variables for later rebuild section
        league_settings = None
        draft_analysis = None

        # macOS optimization: import the large Pre-Draft Analysis sheet FIRST so later sheets aren't wiped
        if platform.system() == 'Darwin' and hasattr(numbers_exporter, 'update_draft_analysis_data'):
            print("Fetching draft analysis data (ADP, rankings, etc.) first for fast CSV base...")
            draft_analysis = yahoo_api.get_player_draft_analysis()
            if draft_analysis:
                numbers_exporter.update_draft_analysis_data(draft_analysis)  # type: ignore[attr-defined]
                print(f"✓ Added draft analysis for {len(draft_analysis)} players")
            else:
                print("⚠ No draft analysis data available")

        if hasattr(numbers_exporter, 'update_league_settings_data'):
            print("Fetching league settings...")
            league_settings = yahoo_api.get_league_settings()
            if league_settings:
                numbers_exporter.update_league_settings_data(league_settings)  # type: ignore[attr-defined]
                # Only build projection sheets here on non-macOS (macOS will do it after all data to avoid duplication)
                if platform.system() != 'Darwin' and hasattr(numbers_exporter, 'setup_projection_sheets'):
                    numbers_exporter.setup_projection_sheets(league_settings)  # type: ignore[attr-defined]
                print(f"✓ Added league settings for: {league_settings.get('league_name', 'Unknown League')}")
            else:
                print("⚠ No league settings returned from API")
        else:
            print("(Slim exporter: skipping league settings & projections)")

        if hasattr(numbers_exporter, 'update_teams_data'):
            print("Fetching teams data...")
            teams_data = yahoo_api.get_teams_data()
            if teams_data:
                numbers_exporter.update_teams_data(teams_data)  # type: ignore[attr-defined]
                print(f"✓ Added {len(teams_data)} teams to file")

        # Windows / non-macOS path still needs to fetch draft analysis if not already done
        if platform.system() != 'Darwin' and hasattr(numbers_exporter, 'update_draft_analysis_data'):
            print("Fetching draft analysis data (ADP, rankings, etc.)...")
            draft_analysis = yahoo_api.get_player_draft_analysis()
            if draft_analysis:
                numbers_exporter.update_draft_analysis_data(draft_analysis)  # type: ignore[attr-defined]
                print(f"✓ Added draft analysis for {len(draft_analysis)} players")
            else:
                print("⚠ No draft analysis data available")

        # Rebuild projection sheets & draft board if we have analysis
        if hasattr(numbers_exporter, 'setup_projection_sheets') and hasattr(numbers_exporter, 'create_draft_board'):
            try:
                # On macOS we deferred projection sheet creation until all base data (league settings + analysis) exist
                if platform.system() == 'Darwin' and league_settings:
                    print("Building projection sheets...")
                    numbers_exporter.setup_projection_sheets(league_settings)  # type: ignore[attr-defined]
                    print("✓ Projection sheets ready")
                elif platform.system() != 'Darwin':
                    print("Rebuilding projection sheets...")
                    numbers_exporter.setup_projection_sheets(league_settings)  # type: ignore[attr-defined]
                    print("✓ Projection sheets ready")
                print("Creating Draft Board...")
                numbers_exporter.create_draft_board()  # type: ignore[attr-defined]
                print("✓ Draft Board created")
            except Exception:
                print("⚠ Failed to build projection sheets or draft board")

        if hasattr(numbers_exporter, 'timestamp'):
            numbers_exporter.timestamp()  # type: ignore[attr-defined]

        if platform.system() == 'Darwin':
            if os.path.exists(numbers_filename):
                print(f"✓ Numbers file ready: {numbers_filename}")
            else:
                print("⚠ Numbers file not found")
        else:
            print(f"✓ Excel file ready: {xlsx_filename}")

        return True

    except Exception as e:
        print(f"✗ Data initialization failed: {e}")
        return False

def main():
    """Main setup function"""
    print("Yahoo Fantasy Draft Monitor")
    print("=" * 35)

    if not os.path.exists('.env'):
        print("✗ .env file not found!")
        print("Please copy .env.example to .env and fill in your values")
        return

    required_vars = ['YAHOO_CLIENT_ID', 'YAHOO_CLIENT_SECRET', 'LEAGUE_ID']
    missing_vars = []

    for var in required_vars:
        if not os.getenv(var):
            missing_vars.append(var)

    if missing_vars:
        print(f"✗ Missing required environment variables: {', '.join(missing_vars)}")
        print("Please check your .env file")
        return

    if not setup_yahoo_authentication():
        return

    # Check if draft file already exists
    requested_filename = os.getenv('FILENAME', 'fantasy_draft_data.xlsx')
    if platform.system() == 'Darwin':
        # macOS: check for .numbers file
        if not requested_filename.lower().endswith('.numbers'):
            check_filename = requested_filename.replace('.xlsx', '.numbers') if requested_filename.endswith('.xlsx') else requested_filename + '.numbers'
        else:
            check_filename = requested_filename
    else:
        # Windows: check for .xlsx file
        if not requested_filename.lower().endswith('.xlsx'):
            check_filename = requested_filename + '.xlsx'
        else:
            check_filename = requested_filename

    if os.path.exists(check_filename):
        response = input(f"\n⚠ File '{check_filename}' already exists. Overwrite it? (y/n): ").lower().strip()
        if response != 'y':
            print("\n✓ Setup cancelled. Your existing file was not modified.")
            return

    print("\n=== Initializing Draft Workbook ===")
    if initialize_data():
        print("\n✓ Setup completed successfully!")
        if platform.system() == 'Darwin':
            print("\nNext steps:")
            print("1. Open your Numbers file (fantasy_draft_data.numbers)")
            print("2. Run: bash run_monitor.command")
        else:
            print("\nNext steps:")
            print("1. Open your Excel file (fantasy_draft_data.xlsx)")
            print("2. Run: python draft_monitor.py")
    else:
        print("\n✗ Setup completed with errors")

if __name__ == "__main__":
    main()