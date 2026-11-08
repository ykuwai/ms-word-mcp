"""Write Word's own Math AutoCorrect list to docs/reference/math-autocorrect.json.

    uv run python scripts/probes/windows_math_autocorrect.py

The list (\\alpha -> α, \\rightarrow -> →, ...) is what Word itself turns TeX
control words into while the user types an equation. word_insert_equation
converts LaTeX with it, so the table is Word's, not ours. Word for Mac
answers missing value for each entry's value, so it is read on Windows.
Reads only; touches no document.
"""

import json
from pathlib import Path

import pythoncom
import win32com.client

pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
entries = word.OMathAutoCorrect.Entries
table = {}
for i in range(1, entries.Count + 1):
    entry = entries.Item(i)
    table[entry.Name] = entry.Value
functions = word.OMathAutoCorrect.Functions
names = [functions.Item(i).Name for i in range(1, functions.Count + 1)]
out = Path(__file__).resolve().parents[2] / "docs" / "reference" / "math-autocorrect.json"
out.write_text(json.dumps({
    "source": f"Word {word.Version} (build {word.Build}) for Windows, "
              "Application.OMathAutoCorrect.Entries and .Functions",
    "entries": dict(sorted(table.items())),
    "functions": sorted(names),
}, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
print(f"{len(table)} entries and {len(names)} functions written to {out}")
