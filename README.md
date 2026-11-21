# Office live agent

A Claude Managed Agent that works in your open Word, Excel and PowerPoint windows on Windows, as if you were doing the task yourself. The agent runs in Anthropic's cloud. A Python script on your PC runs its tool calls by controlling the Office apps you have installed.

## What it can do

- **Word:** read paragraphs, replace text, insert headings, bullets and tables, delete paragraphs, add comments, and toggle track changes.
- **Excel:** read cells and formulas, write values and formulas, add sheets, format cells, and create charts.
- **PowerPoint:** read slides, replace text, add and delete slides.
- **Files:** open and create files in a working folder you choose, save, save a copy, and export to PDF.
- **Outlook:** open an email draft for you to review. It never sends email.

Every edit asks for your approval (`Allow? [y/N]`) before it runs. Saving, exporting and drafting email always ask. Edits are not written to disk until you approve a save.

## Requirements

- Windows with desktop Microsoft Office (Word, Excel, PowerPoint). Classic Outlook is needed for email drafts.
- Python 3.10 or newer.
- An Anthropic API key.

## Setup

1. Install the dependencies:

   ```powershell
   pip install -r requirements.txt
   ```

2. Set your API key. This stores it for new terminal windows:

   ```powershell
   setx ANTHROPIC_API_KEY "your-api-key-here"
   ```

3. In the Claude Console, create a **cloud** environment and copy its ID (starts with `env_`).
4. Create the agent. Leave the placeholder `AGENT_ID` in `office_live.py` and run:

   ```powershell
   python office_live.py
   ```

   It creates an agent with the tools it needs and prints its ID (starts with `agent_`).
5. Create a file named `config_local.py` next to `office_live.py` with your IDs:

   ```python
   ENVIRONMENT_ID = "env_..."
   AGENT_ID = "agent_..."
   ```

   This file is listed in `.gitignore`, so it stays out of the repo.

## Usage

```powershell
python office_live.py
```

Or double-click `office_live.bat` after editing its folder path. Choose a working folder from the menu, open a document in Word, Excel or PowerPoint, and type a request, for example:

```
Add a heading called Summary and two bullet points to the open Word document.
```

Type `exit` to quit. To undo a change, press Ctrl+Z in the app, or close the document without saving.

## Notes
© 2026 Michael Antwi. All rights reserved.

- Opening, creating and saving only work inside the working folder. Documents you already have open can be edited wherever they are.
- If OneDrive AutoSave is on for a file, edits are saved as they happen. Keep `CONFIRM_EDITS = True` in that case.
- If Office shows a dialog box, such as a password prompt, the script waits and retries, then reports an error. Dismiss the dialog and ask again.
- Word's find-and-replace is limited to 255 characters per side.
- Run `python office_live.py --tools` to print the tool definitions as JSON.
