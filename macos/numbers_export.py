import os, logging, subprocess, json, shutil
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from base_exporter import BaseDraftExporter
from typing import List

class MacOSDraftExporter(BaseDraftExporter):
	"""Mac exporter with Numbers conversion and live sync capabilities."""

	def __init__(self, filename: str = "fantasy_draft_data.xlsx"):
		super().__init__(filename)
		# (Deferred) Numbers conversion now happens once at end of setup


	def convert_to_numbers(self, force: bool = False):
		"""Ensure a .numbers companion exists. If force=True, regenerate it (overwrites)."""
		numbers_target = self.filename[:-5] + '.numbers'
		xlsx_abs = os.path.abspath(self.filename)
		numbers_abs = os.path.abspath(numbers_target)
		if os.path.exists(numbers_abs) and not force:
			return numbers_abs
		if force and os.path.exists(numbers_abs):
			try:
				# .numbers is a package (directory); remove recursively
				if os.path.isdir(numbers_abs):
					shutil.rmtree(numbers_abs)
				else:
					os.remove(numbers_abs)
			except Exception as e:
				self.logger.debug(f"Failed to remove existing Numbers file for reconversion: {e}")
		applescript = f'''set sourcePath to POSIX file "{xlsx_abs}"
set targetPath to POSIX file "{numbers_abs}"
tell application "Numbers"
  set docRef to open sourcePath
  delay 1
  try
    save docRef in targetPath
  end try
  delay 0.3
  try
    close docRef saving no
  end try
end tell
return "OK"'''
		try:
			res = subprocess.run(["osascript", "-e", applescript], capture_output=True, text=True)
			if res.returncode == 0:
				self.logger.debug(f"Created Numbers file {numbers_abs}{' (recreated)' if force else ''}")
			else:
				self.logger.warning(f"Numbers conversion failed rc={res.returncode}: {res.stderr.strip()}")
		except Exception as e:
			self.logger.warning(f"Numbers conversion error: {e}")
		return numbers_abs

	def _post_append_hook(self, rows: List[List[str]], start_row: int):
		"""Mac-specific hook for Numbers live sync."""
		try:
			self._sync_numbers(rows, start_row)
		except Exception as e:
			self.logger.debug(f"Live sync failed: {e}")


	def _sync_numbers(self, rows: List[List[str]], start_row: int):
		numbers_path = os.path.abspath(self.filename[:-5] + '.numbers')
		if not os.path.exists(numbers_path):
			return
		base = os.path.splitext(os.path.basename(self.filename))[0]
		doc_hint = os.getenv("NUMBERS_DOC_HINT", base)
		cell_cmds = []
		for offset, row in enumerate(rows):
			rnum = start_row + offset
			for col_idx, val in enumerate(row, start=1):
				if val is None:
					v = '""'
				elif isinstance(val, (int,float)):
					v = str(val)
				else:
					s = str(val).replace('"','\\"')
					v = f'"{s}"'
				col_letter = chr(64+col_idx)
				cell_cmds.append(f'set value of cell "{col_letter}{rnum}" to {v}')
		updates = "\n            ".join(cell_cmds)
		last_required = start_row + len(rows) - 1
		script = f'''on run
set requiredRows to {last_required}
set hint to "{doc_hint}"
set numbersPath to POSIX file "{numbers_path}"
tell application "Numbers"
  set _target to missing value
  repeat with d in documents
	if name of d contains hint then
	  set _target to d
	  exit repeat
	end if
  end repeat
  if _target is missing value then return "NO_DOC"
  tell _target
	if not (exists sheet "Draft Results") then return "NO_SHEET"
	tell sheet "Draft Results"
	  if (count of tables) < 1 then return "NO_TABLE"
	  tell table 1
		set currentRows to row count
		repeat while currentRows < requiredRows
		  add row below last row
		  set currentRows to row count
		end repeat
		{updates}
	  end tell
	end tell
  end tell
end tell
return "OK"\nend run'''
		res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
		if res.returncode == 0 and res.stdout.strip().endswith("OK"):
			self.logger.debug("Numbers live sync applied")
		else:
			self.logger.debug(f"Numbers sync status rc={res.returncode} out={res.stdout.strip()} err={res.stderr.strip()}")

__all__ = ["MacOSDraftExporter"]