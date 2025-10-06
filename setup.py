#!/usr/bin/env python3

import os
import logging
from typing import Tuple, Optional
from dotenv import load_dotenv

from yahoo_api import YahooFantasyAPI
import platform

# Determine platform once (avoid repeated expensive calls & branching noise)
IS_MACOS = platform.system() == 'Darwin'

if IS_MACOS:
        from macos.numbers_export import MacOSDraftExporter as DraftExporter  # type: ignore
else:
        from windows.xlsx_export import XlsxDraftExporter as DraftExporter  # type: ignore
"""Setup script: builds canonical XLSX; on macOS a .numbers companion is auto-created by exporter.

Refactored for clarity & maintainability:
    * Single platform flag (IS_MACOS)
    * Centralized filename derivation
    * Reduced duplication in initialization flow
    * Added optional FORCE_OVERWRITE env toggle to skip prompt
"""

# Load environment variables
load_dotenv()

# Configure logging (reduce verbosity for setup)
logging.basicConfig(
    level=logging.WARNING,
    format='%(levelname)s: %(message)s'
)

def setup_yahoo_authentication() -> bool:
    """Ensure Yahoo OAuth is configured (load/refresh or interactive authenticate)."""
    print("\n=== Yahoo API Authentication ===")
    api = YahooFantasyAPI()
    try:
        if api.load_token() and api.refresh_token_if_needed():
            print("✓ Yahoo authentication already configured")
            return True
        print("Yahoo authentication required...")
        api.authenticate()
        print("✓ Yahoo authentication completed")
        return True
    except Exception as e:  # pragma: no cover - defensive
        print(f"✗ Yahoo authentication failed: {e}")
        return False


def derive_filenames(requested: str) -> Tuple[str, str]:
    """Return (xlsx_filename, numbers_filename) from a requested base/filename.

    Accepts raw names with/without .xlsx / .numbers. Empty base ('.numbers') falls back.
    """
    requested = (requested or '').strip() or 'fantasy_draft_data'
    lower = requested.lower()
    if lower.endswith('.numbers'):
        base = requested[:-8] or 'fantasy_draft_data'
    elif lower.endswith('.xlsx'):
        base = requested[:-5]
    else:
        base = requested
    xlsx = base + '.xlsx'
    numbers = base + '.numbers'
    return xlsx, numbers


def resolve_existing_target_name(requested: str) -> str:
    """Determine which file we consider the primary artifact for overwrite check."""
    if IS_MACOS:
        if requested.lower().endswith('.numbers'):
            return requested
        if requested.lower().endswith('.xlsx'):
            return requested[:-5] + '.numbers'
        return requested + '.numbers'
    # Windows / non-mac: ensure .xlsx
    if requested.lower().endswith('.xlsx'):
        return requested
    return requested + '.xlsx'


def prompt_overwrite(path: str) -> bool:
    """Prompt before overwriting existing file unless FORCE_OVERWRITE=1 is set."""
    if not os.path.exists(path):
        return True
    if os.getenv('FORCE_OVERWRITE') == '1':
        print(f"⚠ Overwriting existing file '{path}' (FORCE_OVERWRITE=1)")
        return True
    response = input(f"\n⚠ File '{path}' already exists. Overwrite it? (y/n): ").lower().strip()
    if response == 'y':
        return True
    print("\n✓ Setup cancelled. Your existing file was not modified.")
    return False

def initialize_data() -> bool:
    """Initialize workbook with league + draft data.

    Unified order (all platforms): draft analysis -> league settings -> teams -> projections -> draft board -> timestamp.
    """
    try:
        api = YahooFantasyAPI()
        requested_filename = os.getenv('FILENAME', 'fantasy_draft_data.xlsx')
        xlsx_filename, numbers_filename = derive_filenames(requested_filename)

        exporter = DraftExporter(xlsx_filename)
        print(f"Creating {'Numbers' if IS_MACOS else 'Excel'} file: {numbers_filename if IS_MACOS else xlsx_filename}")

        api.ensure_authenticated()

        league_settings: Optional[dict] = None

        # Draft Board
        if hasattr(exporter, 'create_draft_board'):
            print("Fetching draft analysis data...")
            draft_analysis = api.get_player_draft_analysis()
            if draft_analysis:
                exporter.create_draft_board(draft_analysis)  # type: ignore[attr-defined]
                print(f"✓ Draft analysis: {len(draft_analysis)} players")
            else:
                print("⚠ No draft analysis data")

        # League settings
        if hasattr(exporter, 'update_league_settings_data'):
            print("Fetching league settings...")
            league_settings = api.get_league_settings()
            if league_settings:
                exporter.update_league_settings_data(league_settings)  # type: ignore[attr-defined]
                print(f"✓ League settings: {league_settings.get('league_name', 'Unknown League')}")
            else:
                print("⚠ No league settings returned from API")
        else:
            print("(Slim exporter: skipping league settings & projections)")

        # Teams
        if hasattr(exporter, 'update_teams_data'):
            print("Fetching teams data...")
            teams = api.get_teams_data()
            if teams:
                exporter.update_teams_data(teams)  # type: ignore[attr-defined]
                print(f"✓ Teams: {len(teams)}")
            else:
                print("⚠ No teams data")

        # Draft Results
        if hasattr(exporter, 'update_draft_results_data'):
            print("Fetching draft results data...")
            draft_results = api.get_draft_results()
            if draft_results:
                exporter.update_draft_results_data(draft_results)  # type: ignore[attr-defined]
                print(f"✓ Draft results: {len(draft_results)} picks")
            else:
                print("⚠ No draft results data")

        # Projections after base data (only if we have league settings)
        if league_settings and hasattr(exporter, 'setup_projection_sheets'):
            try:
                print("Building projection sheets...")
                exporter.setup_projection_sheets(league_settings)  # type: ignore[attr-defined]
                print("✓ Projection sheets ready")
            except Exception:  # pragma: no cover - defensive
                print("⚠ Failed to build projection sheets")

        if draft_analysis and hasattr(exporter, 'create_pos_sheets'):
            print("Building position sheets...")
            exporter.create_pos_sheets(draft_analysis)  # type: ignore[attr-defined]
            print("✓ Position sheets ready")

        if IS_MACOS:
            exporter.apply_draft_board_formulas()  # type: ignore[attr-defined]
            msg_symbol = '✓' if os.path.exists(numbers_filename) else '⚠'
            print(f"{msg_symbol} Numbers file: {numbers_filename}")
        else:
            print(f"✓ Excel file: {xlsx_filename}")

        return True
    except Exception as e:  # pragma: no cover - broad catch for CLI robustness
        print(f"✗ Data initialization failed: {e}")
        return False

def main() -> None:
    """CLI entrypoint for initial draft workbook creation."""
    print("Yahoo Fantasy XLSX")
    print("=" * 35)

    if not os.path.exists('.env'):
        print("✗ .env file not found!")
        print("Please copy .env.example to .env and fill in your values")
        return

    required_vars = ['YAHOO_CLIENT_ID', 'YAHOO_CLIENT_SECRET', 'LEAGUE_ID']
    missing = [v for v in required_vars if not os.getenv(v)]
    if missing:
        print(f"✗ Missing required environment variables: {', '.join(missing)}")
        print("Please check your .env file")
        return

    if not setup_yahoo_authentication():
        return

    requested_filename = os.getenv('FILENAME', 'fantasy_draft_data.xlsx')
    target = resolve_existing_target_name(requested_filename)
    if not prompt_overwrite(target):
        return

    print("\n=== Initializing Draft Workbook ===")
    success = initialize_data()
    if success:
        print("\n✓ Setup completed successfully!")
        print("\nNext steps:")
        if IS_MACOS:
            print("1. Open your Numbers file (fantasy_draft_data.numbers)")
            print("2. Run: bash run_monitor.command")
        else:
            print("1. Open your Excel file (fantasy_draft_data.xlsx)")
            print("2. Run: python draft_monitor.py")
    else:
        print("\n✗ Setup completed with errors")

if __name__ == "__main__":
    main()