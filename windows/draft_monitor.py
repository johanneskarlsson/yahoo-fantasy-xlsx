#!/usr/bin/env python3

"""Draft monitor that fetches results periodically and appends new picks to Draft Results."""
import os, sys, time, logging, platform
from dotenv import load_dotenv
from yahoo_api import YahooFantasyAPI

if platform.system() == 'Darwin':
    from macos.numbers_export import MacOSDraftExporter as DraftExporter  # type: ignore
else:
    from windows.xlsx_export import XlsxDraftExporter as DraftExporter  # type: ignore

load_dotenv()

logging.basicConfig(level=logging.WARNING, format='%(levelname)s: %(message)s')
LOG = logging.getLogger("draft_monitor")

INTERVAL = 10  # Check for new picks every 10 seconds

filename = os.getenv('FILENAME', 'fantasy_draft_data.xlsx')
if filename.lower().endswith('.numbers'):
    base = filename[:-8] or 'fantasy_draft_data'
    filename = base + '.xlsx'
    LOG.warning("FILENAME was a .numbers package; using canonical '%s'", filename)
elif not filename.lower().endswith('.xlsx'):
    filename += '.xlsx'

exporter = DraftExporter(filename)
api = YahooFantasyAPI()
seen_picks: set[int] = set()


def _scalar(v):
    if isinstance(v, dict):
        return v.get('#text') or v.get('full') or v.get('name') or ''
    return v

def _player_key(dr):
    # Get player_key directly from draft result
    return dr.get('player_key', '')

def collect_new(draft_results):
    rows = []
    for dr in draft_results:
        try:
            pick_raw = _scalar(dr.get('pick'))
            if pick_raw is None:
                continue
            pick_num = int(pick_raw)
            if pick_num in seen_picks:
                continue
            rnd = _scalar(dr.get('round')) or ''
            team_key = _scalar(dr.get('team_key')) or ''
            player_key = _player_key(dr) or ''
            rows.append([rnd, pick_num, player_key, team_key, ""])
            seen_picks.add(pick_num)
        except Exception:
            continue
    rows.sort(key=lambda r: r[1])
    return rows

def main():
    print(f"🏒 Yahoo Fantasy Draft Monitor")
    print(f"📊 Monitoring file: {exporter.filename}")
    print(f"⏱️  Checking every {INTERVAL} seconds")

    # Test API connection
    try:
        print("🔗 Testing Yahoo API connection...")
        api.ensure_authenticated()
        print("✅ Connected to Yahoo API")
    except Exception as e:
        print(f"❌ Failed to connect to Yahoo API: {e}")
        print("Please run 'python setup.py' first to authenticate")
        return

    print("🔄 Monitoring... (Press Ctrl+C to stop)")
    print()

    polls = 0
    try:
        while True:
            start = time.time()
            try:
                results = api.get_draft_results() or []
                polls += 1

                # Log draft status on first poll
                if polls == 1:
                    if results:
                        print(f"📋 Found {len(results)} total draft picks so far")
                    else:
                        print("📋 No draft picks found yet")

                new_rows = collect_new(results)

                if new_rows:
                    try:
                        exporter.append_draft_results(new_rows)
                        exporter.add_timestamp()
                        for r in new_rows:
                            print(f"✅ Pick {r[1]}: {r[2]} (Team: {r[3]})")
                    except Exception as e:
                        print(f"⚠️  Error saving draft picks: {e}")
                else:
                    # Show periodic status so user knows it's working
                    if polls % 6 == 1:  # Every ~60 seconds (6 polls × 10s)
                        print(f"⏳ Still monitoring... ({polls} checks completed)")

            except Exception as e:
                print(f"⚠️  Error during check #{polls}: {e}")
                # Continue monitoring even if one check fails

            remaining = INTERVAL - (time.time() - start)
            if remaining > 0:
                time.sleep(remaining)
    except KeyboardInterrupt:
        print(f"\n🛑 Stopped by user after {polls} checks.")

if __name__ == "__main__":
    main()