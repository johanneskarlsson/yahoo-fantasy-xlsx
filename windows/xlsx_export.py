import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from base_exporter import BaseDraftExporter
from typing import List

class XlsxDraftExporter(BaseDraftExporter):
	"""Windows exporter for Excel files only."""

	def __init__(self, filename: str = "fantasy_draft_data.xlsx"):
		super().__init__(filename)

	def _post_append_hook(self, rows: List[List[str]], start_row: int):
		"""Windows exporter has no additional actions after appending picks."""
		pass  # No additional actions needed for Windows


__all__ = ["XlsxDraftExporter"]