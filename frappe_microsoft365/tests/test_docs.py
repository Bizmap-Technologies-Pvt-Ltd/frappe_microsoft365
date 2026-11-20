"""Checks on the documentation that need no tenant and no PowerShell.

An admin pastes these commands as they are, so a wrong one fails in front of the person least able
to tell whether the fault is theirs. Parsing every block with pwsh belongs in CI; this catches the
mistake that has actually shipped, cheaply and on every run.
"""

import re
import unittest
from pathlib import Path

#: The repository root: docs/ and README.md sit next to the package directory.
ROOT = Path(__file__).resolve().parents[2]

FENCE = re.compile(r"^\s*```(\w*)\s*$")
POWERSHELL = {"powershell", "pwsh", "ps1", "ps"}

#: `-Name true` or `-Name:true`. PowerShell reads a bare word as a string, and a Boolean
#: parameter refuses the string "true" — it needs `$true`.
BARE_BOOLEAN = re.compile(r"(?<![\w$])-[A-Za-z][\w]*(?::|\s+)(true|false)\b", re.IGNORECASE)


def markdown_files():
	files = sorted((ROOT / "docs").glob("**/*.md"))
	readme = ROOT / "README.md"
	if readme.exists():
		files.append(readme)
	return files


def powershell_blocks(text):
	"""Yield (first line number, lines) for every fenced PowerShell block."""
	inside, language, start, lines = False, "", 0, []
	for number, line in enumerate(text.splitlines(), start=1):
		match = FENCE.match(line)
		if not match:
			if inside:
				lines.append(line)
			continue
		if not inside:
			inside, language, start, lines = True, match.group(1).lower(), number + 1, []
		else:
			if language in POWERSHELL:
				yield start, lines
			inside = False


def bare_booleans(text):
	"""Every `-Parameter true|false` inside a PowerShell block, as (line number, line)."""
	found = []
	for start, lines in powershell_blocks(text):
		for offset, line in enumerate(lines):
			code = line.split("#", 1)[0]
			if BARE_BOOLEAN.search(code):
				found.append((start + offset, line.strip()))
	return found


class TestPowerShellInDocs(unittest.TestCase):
	def test_no_boolean_parameter_is_given_a_bare_true_or_false(self):
		files = markdown_files()
		if not files:
			# An install without the repository around it (a wheel) ships no docs to check.
			self.skipTest("docs/ is not next to the package")
		problems = []
		for path in files:
			for number, line in bare_booleans(path.read_text(encoding="utf-8")):
				problems.append(f"{path.relative_to(ROOT)}:{number}: {line}")
		self.assertEqual(problems, [], "use $true / $false for Boolean parameters")

	def test_the_check_catches_what_shipped_in_1_1_0(self):
		shipped = (
			"```powershell\n"
			"Set-CsTeamsMeetingConfiguration -EnableGraphTranscriptAccess true "
			"-EnableAttributedTranscripts true -Identity Global\n"
			"```\n"
		)
		self.assertEqual(len(bare_booleans(shipped)), 1)
		self.assertEqual(bare_booleans(shipped.replace(" true", " $true")), [])
		self.assertEqual(len(bare_booleans(shipped.replace(" true", ":false"))), 1)

	def test_only_powershell_blocks_are_read(self):
		other = "```bash\nsomething --flag -x true\n```\nOutside -Enabled true\n"
		self.assertEqual(bare_booleans(other), [])
