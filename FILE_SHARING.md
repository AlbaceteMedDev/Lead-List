# Sharing Files with Claude Code

When working on this pipeline with Claude Code, you can hand over files (CSVs
from AcuityMD, scoring scripts, dashboard code, config overrides, etc.) in a
few different ways. Pick whichever fits your workflow.

## Options

1. **Drag-and-drop into the terminal.** Works in most terminals (macOS
   Terminal, iTerm2, Warp). Dropping a file pastes its path into the prompt.
2. **Type the path directly.** `/path/to/file.py` or `~/Downloads/file.py`.
   Claude can open it with the Read tool.
3. **Copy it into this repo.** Drop the file into `/home/user/Lead-List/`
   (via your file manager, `scp`, or `cp`), then tell Claude the filename.
   For AcuityMD CSVs specifically, place them in `data/input/` so the
   pipeline picks them up automatically.
4. **Paste the contents directly into chat.** Works for text files that are
   not too large.
5. **Reference by `@`.** Type `@` in Claude Code to open a file picker
   scoped to the current working directory.

## Recommendations for this project

- **Raw AcuityMD CSV exports** → copy into `data/input/` (option 3). The
  ingestion step reads every CSV in that directory.
- **Scoring scripts, dashboard code, or one-off utilities from another
  session** → copy into the repo root or `src/` (option 3), then tell
  Claude the filename so it can apply edits in place.
- **Config overrides** (`mac_jurisdictions.json`, `hospital_keywords.json`,
  `email_templates.json`) → drop into `config/`.
- **Short snippets or error logs** → paste directly into chat (option 4).
- **Files already on disk** → reference by `@` (option 5) or give the
  absolute path (option 2).

Once the file is in place, tell Claude the filename and what you want done
with it.
