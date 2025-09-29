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

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', handlers=[logging.StreamHandler(sys.stdout)])
LOG = logging.getLogger("draft_monitor")

INTERVAL = int(os.getenv("DRAFT_MONITOR_INTERVAL", "10") or 10)

filename = os.getenv('EXCEL_FILENAME', 'fantasy_draft_data.xlsx')
if filename.lower().endswith('.numbers'):
    base = filename[:-8] or 'fantasy_draft_data'
    filename = base + '.xlsx'
    LOG.warning("EXCEL_FILENAME was a .numbers package; using canonical '%s'", filename)

exporter = DraftExporter(filename)
api = YahooFantasyAPI()
seen_picks: set[int] = set()


def _scalar(v):
    if isinstance(v, dict):
        return v.get('#text') or v.get('full') or v.get('name') or ''
    return v

def _player_name(dr):
    p = dr.get('player') if isinstance(dr, dict) else None
    if isinstance(p, dict):
        nm = p.get('name')
        if isinstance(nm, dict):
            return nm.get('full') or ''
    return dr.get('player_name') or ''

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
            rows.append([rnd, pick_num, _player_name(dr) or '', team_key, ""])
            seen_picks.add(pick_num)
        except Exception:
            continue
    rows.sort(key=lambda r: r[1])
    return rows

def main():
    LOG.info("Starting draft monitor (interval=%ss, file=%s)" % (INTERVAL, exporter.filename))
    try:
        while True:
            start = time.time()
            results = api.get_draft_results() or []
            new_rows = collect_new(results)
            if new_rows:
                exporter.append_draft_results(new_rows)
                exporter.add_timestamp()
                for r in new_rows:
                    LOG.info("Pick %s: %s", r[1], r[2])
            remaining = INTERVAL - (time.time() - start)
            if remaining > 0:
                time.sleep(remaining)
    except KeyboardInterrupt:
        LOG.info("Stopped by user.")

if __name__ == "__main__":
    main()