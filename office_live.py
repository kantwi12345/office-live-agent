"""Live Office bridge: a Claude Managed Agent drives your real Word, Excel, PowerPoint and Outlook.

The agent runs in Anthropic's cloud. It asks for tools; this script runs them on your PC by
controlling the Office apps you have installed, in visible windows, as if you were doing it.

Needs: Windows, desktop Microsoft Office (Word/Excel/PowerPoint), classic Outlook for email drafts.
Setup: pip install anthropic pywin32
Run:   python office_live.py             (menu to pick the working folder)
       python office_live.py "D:\\Work"   (use a folder directly)
       python office_live.py --tools     (print the tool definitions as JSON)
"""
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

try:
    from config_local import ENVIRONMENT_ID, AGENT_ID   # your private values, not committed
except ImportError:
    ENVIRONMENT_ID = "env_PASTE_YOUR_ENVIRONMENT_ID"
    AGENT_ID = "PASTE_YOUR_LIVE_OFFICE_AGENT_ID_HERE"
# True: you approve every edit before it happens. False: edits in open documents run without
# asking (saving, exporting and email drafts still ask). Keep True if OneDrive AutoSave is on.
CONFIRM_EDITS = True

SYSTEM = (
    "You control the user's real Microsoft Office apps (Word, Excel, PowerPoint, Outlook) on their "
    "Windows PC through tools, as if you were sitting at the keyboard. Documents open in visible "
    "windows, so the user can watch. Call list_open first to see what is already open; the user "
    "may want you to work on those. Use open_file for files in the working folder and create_file "
    "for new ones. Read a document (word_read, excel_read, ppt_read) before editing it, make "
    "small targeted edits (word_edit, excel_edit, ppt_edit), then read it back when the result "
    "matters. Edits happen in the open document and are NOT saved until you call save_document, "
    "which the user approves; the user can also undo in the app. Never claim something was saved "
    "unless save_document succeeded. outlook_draft only opens a draft for the user to review and "
    "send; you cannot send email. Keep replies short and say exactly what you changed."
)

MAX_CHARS = 20000
OFFICE_EXT = {".docx", ".doc", ".xlsx", ".xlsm", ".xls", ".pptx", ".ppt"}
APPS = {"word": "Word.Application", "excel": "Excel.Application", "powerpoint": "PowerPoint.Application"}
EXT_KIND = {
    ".docx": "word", ".doc": "word",
    ".xlsx": "excel", ".xlsm": "excel", ".xls": "excel",
    ".pptx": "powerpoint", ".ppt": "powerpoint",
}
NEW_EXT = {"word": ".docx", "excel": ".xlsx", "powerpoint": ".pptx"}

# Word built-in style constants (language independent)
W_NORMAL, W_BULLET = -1, -49
W_HEADING = {1: -2, 2: -3, 3: -4, 4: -5}
CHART_TYPES = {"column": 51, "bar": 57, "line": 4, "pie": 5, "scatter": -4169}

ROOT = None  # set in set_root()


def _obj(**props):
    return {"type": "object", "properties": props}


_FILE = {"type": "string", "description": "Open document's name or path. Omit for the active one."}

BLOCKS_SCHEMA = {
    "type": "array",
    "items": {
        "type": "object",
        "properties": {
            "type": {"type": "string", "enum": ["heading", "paragraph", "bullets", "table"]},
            "text": {"type": "string"},
            "level": {"type": "integer"},
            "items": {"type": "array", "items": {"type": "string"}},
            "rows": {"type": "array", "items": {"type": "array", "items": {"type": "string"}}},
        },
        "required": ["type"],
    },
}

TOOLS = [
    {
        "type": "custom",
        "name": "list_open",
        "description": "List the Word, Excel and PowerPoint documents that are open right now.",
        "input_schema": _obj(),
    },
    {
        "type": "custom",
        "name": "list_files",
        "description": "List Office files and subfolders in the working folder. 'path' is relative to it.",
        "input_schema": _obj(path={"type": "string"}),
    },
    {
        "type": "custom",
        "name": "open_file",
        "description": "Open an Office file from the working folder in its app (visible window).",
        "input_schema": {**_obj(path={"type": "string"}), "required": ["path"]},
    },
    {
        "type": "custom",
        "name": "create_file",
        "description": "Create a new blank Word, Excel or PowerPoint file in the working folder and leave "
                       "it open for editing. 'app' is word, excel or powerpoint. The user approves. "
                       "Never overwrites.",
        "input_schema": {
            **_obj(app={"type": "string", "enum": list(APPS)}, path={"type": "string"}),
            "required": ["app", "path"],
        },
    },
    {
        "type": "custom",
        "name": "word_read",
        "description": "Read the open Word document: numbered paragraphs with their style names.",
        "input_schema": _obj(file=_FILE),
    },
    {
        "type": "custom",
        "name": "word_edit",
        "description": "Edit the open Word document. 'ops' run in order. Each op has 'op' and: "
                       "replace (find, replace; all occurrences, text up to 255 characters); "
                       "insert (blocks, optional after_text: insert after the first paragraph "
                       "containing it, otherwise append at the end; tables only when appending); "
                       "delete_paragraph (containing; must match 3 or fewer); "
                       "comment (containing, text: adds a review comment); "
                       "track_changes (enabled: true or false). Not saved until save_document.",
        "input_schema": {
            **_obj(
                file=_FILE,
                ops={
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "op": {"type": "string", "enum": ["replace", "insert", "delete_paragraph", "comment", "track_changes"]},
                            "find": {"type": "string"},
                            "replace": {"type": "string"},
                            "after_text": {"type": "string"},
                            "containing": {"type": "string"},
                            "text": {"type": "string"},
                            "enabled": {"type": "boolean"},
                            "blocks": BLOCKS_SCHEMA,
                        },
                        "required": ["op"],
                    },
                },
            ),
            "required": ["ops"],
        },
    },
    {
        "type": "custom",
        "name": "excel_read",
        "description": "Read cells from the open Excel workbook (values, with formulas in brackets). "
                       "'sheet' defaults to the active sheet; 'range' (like A1:D20) defaults to the used range.",
        "input_schema": _obj(file=_FILE, sheet={"type": "string"}, range={"type": "string"}),
    },
    {
        "type": "custom",
        "name": "excel_edit",
        "description": "Edit the open Excel workbook. 'ops' run in order. Each op has 'op' and: "
                       "write (sheet, cell: top-left address like 'A1', values: list of rows; strings "
                       "starting with '=' are formulas); add_sheet (name); format (sheet, range, bold, "
                       "number_format like '0.00' or '#,##0', autofit: true to fit column widths); "
                       "chart (sheet, source: data range like 'A1:B6', type: column|bar|line|pie|scatter, "
                       "title, anchor: cell where the chart's top-left goes). Not saved until save_document.",
        "input_schema": {
            **_obj(
                file=_FILE,
                ops={
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "op": {"type": "string", "enum": ["write", "add_sheet", "format", "chart"]},
                            "sheet": {"type": "string"},
                            "name": {"type": "string"},
                            "cell": {"type": "string"},
                            "values": {"type": "array", "items": {"type": "array", "items": {}}},
                            "range": {"type": "string"},
                            "bold": {"type": "boolean"},
                            "number_format": {"type": "string"},
                            "autofit": {"type": "boolean"},
                            "source": {"type": "string"},
                            "type": {"type": "string", "enum": list(CHART_TYPES)},
                            "title": {"type": "string"},
                            "anchor": {"type": "string"},
                        },
                        "required": ["op"],
                    },
                },
            ),
            "required": ["ops"],
        },
    },
    {
        "type": "custom",
        "name": "ppt_read",
        "description": "Read the open PowerPoint presentation: text on each slide, numbered from 1.",
        "input_schema": _obj(file=_FILE),
    },
    {
        "type": "custom",
        "name": "ppt_edit",
        "description": "Edit the open PowerPoint presentation. 'ops' run in order. Each op has 'op' and: "
                       "replace_text (find, replace); add_slide (title, bullets: added at the end); "
                       "delete_slide (index, from 1). Not saved until save_document.",
        "input_schema": {
            **_obj(
                file=_FILE,
                ops={
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "op": {"type": "string", "enum": ["replace_text", "add_slide", "delete_slide"]},
                            "find": {"type": "string"},
                            "replace": {"type": "string"},
                            "title": {"type": "string"},
                            "bullets": {"type": "array", "items": {"type": "string"}},
                            "index": {"type": "integer"},
                        },
                        "required": ["op"],
                    },
                },
            ),
            "required": ["ops"],
        },
    },
    {
        "type": "custom",
        "name": "save_document",
        "description": "Save an open document. With no other argument it saves in place. 'save_as' saves it "
                       "under a new name in the working folder (never overwrites); 'export_pdf' writes a PDF "
                       "copy there. The user approves.",
        "input_schema": {
            **_obj(
                app={"type": "string", "enum": list(APPS)},
                file=_FILE,
                save_as={"type": "string"},
                export_pdf={"type": "string"},
            ),
            "required": ["app"],
        },
    },
    {
        "type": "custom",
        "name": "outlook_draft",
        "description": "Open a new email draft in Outlook for the user to review and send. It is never sent "
                       "automatically. 'attachments' are paths inside the working folder. The user approves.",
        "input_schema": {
            **_obj(
                to={"type": "string"},
                subject={"type": "string"},
                body={"type": "string"},
                attachments={"type": "array", "items": {"type": "string"}},
            ),
            "required": ["subject", "body"],
        },
    },
]


# ---------- helpers ----------
def safe(rel):
    p = (ROOT / rel).resolve()
    if p != ROOT and ROOT not in p.parents:
        raise ValueError(f"Path is outside the chosen folder: {rel}")
    return p


def unique(dest):
    if not dest.exists():
        return dest
    n = 1
    while True:
        cand = dest.with_name(f"{dest.stem} ({n}){dest.suffix}")
        if not cand.exists():
            return cand
        n += 1


def clip(text):
    if len(text) > MAX_CHARS:
        return text[:MAX_CHARS] + f"\n... (cut at {MAX_CHARS} characters)"
    return text


def focus_console():
    """Bring this terminal window to the front (best effort) so approval prompts are not hidden behind Office."""
    if os.name != "nt":
        return
    try:
        import ctypes

        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 9)  # restore if minimized
            ctypes.windll.user32.SetForegroundWindow(hwnd)
    except Exception:
        pass


def ask(action, details=None):
    focus_console()
    print(f"\a\n\nThe agent wants to {action}")
    for d in details or []:
        print(f"    - {d}")
    return input("Allow? [y/N] ").strip().lower() == "y"


def describe(op):
    return json.dumps(op, ensure_ascii=False)[:200]


def confirm_edit(kind, item, ops):
    if not CONFIRM_EDITS:
        return True
    return ask(f"edit the open {kind.title()} document {item.Name}:", [describe(o) for o in ops])


def _com():
    import win32com.client as wc

    return wc


def get_app(kind, start=False):
    wc = _com()
    name = APPS[kind]
    try:
        app = wc.GetActiveObject(name)
    except Exception:
        if not start:
            return None
        app = wc.Dispatch(name)
    try:
        app.Visible = True
    except Exception:
        pass
    return app


COLLECTION_ATTR = {"word": "Documents", "excel": "Workbooks", "powerpoint": "Presentations"}
ACTIVE_ATTR = {"word": "ActiveDocument", "excel": "ActiveWorkbook", "powerpoint": "ActivePresentation"}


def collection(kind, app):
    return getattr(app, COLLECTION_ATTR[kind])


def open_items(kind, app):
    coll = collection(kind, app)
    return [coll.Item(i) for i in range(1, coll.Count + 1)]


def active_item(kind, app):
    try:
        return getattr(app, ACTIVE_ATTR[kind])
    except Exception:
        return None


def find_item(kind, file=None, start=False):
    app = get_app(kind, start=start)
    if app is None:
        raise RuntimeError(f"{kind.title()} is not running. Use open_file or create_file first.")
    items = open_items(kind, app)
    if not items:
        raise RuntimeError(f"No {kind.title()} documents are open.")
    if not file:
        return app, active_item(kind, app) or items[0]
    key = str(file).strip().lower()
    try:
        full = str((ROOT / file).resolve()).lower()
    except Exception:
        full = key
    for it in items:
        name, fullname = str(it.Name).lower(), str(it.FullName).lower()
        if key in (name, fullname) or full == fullname:
            return app, it
    names = ", ".join(str(it.Name) for it in items)
    raise RuntimeError(f"{file!r} is not open in {kind.title()}. Open documents: {names}")


def existing_office_file(rel):
    p = safe(rel)
    if p.suffix.lower() not in OFFICE_EXT or not p.is_file() or p.name.startswith("~$"):
        raise ValueError(f"Not an Office file in the working folder: {rel}")
    return p


def target_in_root(rel, ext):
    p = safe(rel)
    if p.suffix.lower() != ext:
        p = p.with_name(p.name + ext)
    return unique(p)


def grid(v):
    if isinstance(v, tuple):
        return [list(r) if isinstance(r, tuple) else [r] for r in v]
    return [[v]]


def cell_text(v):
    return "" if v is None else str(v)


# ---------- files and opening ----------
def list_open(inp):
    out = []
    for kind in APPS:
        app = get_app(kind)
        if app is None:
            continue
        for it in open_items(kind, app):
            out.append(f"{kind} | {it.Name} | {it.FullName}")
    return "\n".join(out) or "No Word, Excel or PowerPoint documents are open."


def list_files(inp):
    folder = safe(inp.get("path") or ".")
    if not folder.is_dir():
        return "Not a folder."
    rows = []
    for p in sorted(folder.iterdir()):
        if p.name.startswith(("~$", ".")):
            continue
        if p.is_file() and p.suffix.lower() not in OFFICE_EXT:
            continue
        st = p.stat()
        kind = "dir" if p.is_dir() else "file"
        rows.append(
            f"{kind} | {p.relative_to(ROOT)} | {st.st_size} bytes | {datetime.fromtimestamp(st.st_mtime):%Y-%m-%d}"
        )
    return clip("\n".join(rows[:500])) or "(no Office files here)"


def open_file(inp):
    p = existing_office_file(inp["path"])
    kind = EXT_KIND[p.suffix.lower()]
    app = get_app(kind, start=True)
    for it in open_items(kind, app):
        if str(it.FullName).lower() == str(p).lower():
            try:
                it.Activate()
            except Exception:
                pass
            return f"{p.name} was already open in {kind.title()}."
    collection(kind, app).Open(str(p))
    return f"Opened {p.relative_to(ROOT)} in {kind.title()}."


def create_file(inp):
    kind = inp["app"]
    if kind not in APPS:
        return "app must be word, excel or powerpoint."
    target = target_in_root(inp["path"], NEW_EXT[kind])
    if not ask(f"create a new {kind.title()} file:", [str(target.relative_to(ROOT))]):
        return "The user declined. Nothing was created."
    app = get_app(kind, start=True)
    if kind == "word":
        doc = app.Documents.Add()
        doc.SaveAs2(str(target))
    elif kind == "excel":
        wb = app.Workbooks.Add()
        wb.SaveAs(str(target), 51)
    else:
        pres = app.Presentations.Add()
        pres.SaveAs(str(target))
    return f"Created and opened {target.relative_to(ROOT)}. It is empty."


# ---------- Word ----------
def w_style(doc, rng, const):
    try:
        rng.Style = doc.Styles.Item(const)
    except Exception:
        pass


def w_last_empty(doc):
    last = doc.Paragraphs.Item(doc.Paragraphs.Count).Range
    if last.Text not in ("\r", "\r\x07"):
        doc.Content.InsertParagraphAfter()
        last = doc.Paragraphs.Item(doc.Paragraphs.Count).Range
    return last


def w_append_paragraph(doc, text, style):
    last = w_last_empty(doc)
    last.InsertBefore(text)
    n = doc.Paragraphs.Count
    w_style(doc, doc.Paragraphs.Item(n).Range, style)


def w_add_table(doc, rows):
    if not rows:
        return
    last = w_last_empty(doc)
    w_style(doc, last, W_NORMAL)
    width = max(len(r) for r in rows)
    tbl = doc.Tables.Add(last, len(rows), width)
    try:
        tbl.Style = "Table Grid"
    except Exception:
        pass
    for r, row in enumerate(rows, 1):
        for c, val in enumerate(row, 1):
            tbl.Cell(r, c).Range.Text = str(val)


def w_items(blocks):
    items = []
    for b in blocks:
        kind = b.get("type")
        if kind == "heading":
            items.append((W_HEADING[max(1, min(int(b.get("level") or 1), 4))], b.get("text", "")))
        elif kind == "paragraph":
            items.append((W_NORMAL, b.get("text", "")))
        elif kind == "bullets":
            items.extend((W_BULLET, str(i)) for i in b.get("items") or [])
        elif kind == "table":
            items.append(("table", b.get("rows") or []))
    return items


def w_find_paragraph(doc, text):
    needle = text.lower()
    for i in range(1, doc.Paragraphs.Count + 1):
        if needle in str(doc.Paragraphs.Item(i).Range.Text).lower():
            return i
    return None


def word_read(inp):
    app, doc = find_item("word", inp.get("file"))
    total = doc.Paragraphs.Count
    lines = [f"{doc.Name} | {total} paragraphs | {doc.Tables.Count} tables | unsaved changes: {not bool(doc.Saved)}"]
    for i in range(1, min(total, 400) + 1):
        para = doc.Paragraphs.Item(i)
        text = str(para.Range.Text).rstrip("\r\x07").strip()
        if not text:
            continue
        try:
            style = str(para.Range.Style.NameLocal)
        except Exception:
            style = "?"
        lines.append(f"{i} [{style}] {text}")
    if total > 400:
        lines.append(f"... showing the first 400 of {total} paragraphs")
    return clip("\n".join(lines))


def word_edit(inp):
    app, doc = find_item("word", inp.get("file"))
    ops = inp.get("ops") or []
    if not ops:
        return "No ops given."
    if not confirm_edit("word", doc, ops):
        return "The user declined. The document was not changed."
    report = []
    for op in ops:
        kind = op.get("op")
        if kind == "replace":
            find, repl = op.get("find", ""), op.get("replace", "")
            if not find:
                report.append("replace: 'find' is empty")
                continue
            if len(find) > 255 or len(repl) > 255:
                report.append("replace: Word limits find and replace text to 255 characters")
                continue
            count = str(doc.Content.Text).lower().count(find.lower())
            if count == 0:
                report.append(f"replace {find!r}: text not found")
                continue
            rng = doc.Content
            f = rng.Find
            f.ClearFormatting()
            f.Replacement.ClearFormatting()
            # FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
            # MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace(2 = all)
            f.Execute(find, False, False, False, False, False, True, 1, False, repl, 2)
            report.append(f"replace {find!r}: about {count} occurrence(s) replaced")
        elif kind == "insert":
            items = w_items(op.get("blocks") or [])
            after = op.get("after_text")
            if after:
                if any(s == "table" for s, _ in items):
                    report.append("insert: tables can only be added at the end")
                    continue
                idx = w_find_paragraph(doc, after)
                if idx is None:
                    report.append(f"insert: no paragraph contains {after!r}")
                    continue
                for style, text in items:
                    doc.Paragraphs.Item(idx).Range.InsertParagraphAfter()
                    idx += 1
                    doc.Paragraphs.Item(idx).Range.InsertBefore(text)
                    w_style(doc, doc.Paragraphs.Item(idx).Range, style)
                report.append(f"inserted {len(items)} paragraph(s) after the paragraph containing {after!r}")
            else:
                for style, payload in items:
                    if style == "table":
                        w_add_table(doc, payload)
                    else:
                        w_append_paragraph(doc, payload, style)
                report.append(f"appended {len(items)} item(s) at the end")
        elif kind == "delete_paragraph":
            key = (op.get("containing") or "").lower()
            if not key:
                report.append("delete_paragraph: 'containing' is empty")
                continue
            hits = [i for i in range(1, doc.Paragraphs.Count + 1)
                    if key in str(doc.Paragraphs.Item(i).Range.Text).lower()]
            if not hits:
                report.append(f"delete_paragraph: no paragraph contains {key!r}")
            elif len(hits) > 3:
                report.append(f"delete_paragraph: {len(hits)} paragraphs match {key!r}; be more specific. Nothing deleted.")
            else:
                for i in reversed(hits):
                    doc.Paragraphs.Item(i).Range.Delete()
                report.append(f"deleted {len(hits)} paragraph(s) containing {key!r}")
        elif kind == "comment":
            idx = w_find_paragraph(doc, op.get("containing") or "\0")
            if idx is None:
                report.append(f"comment: no paragraph contains {op.get('containing')!r}")
                continue
            doc.Comments.Add(doc.Paragraphs.Item(idx).Range, op.get("text", ""))
            report.append(f"added a comment on paragraph {idx}")
        elif kind == "track_changes":
            doc.TrackRevisions = bool(op.get("enabled"))
            report.append(f"track changes {'on' if op.get('enabled') else 'off'}")
        else:
            report.append(f"unknown op {kind!r}")
    return f"Edited {doc.Name} (not saved yet).\n" + "\n".join(report)


# ---------- Excel ----------
def excel_read(inp):
    app, wb = find_item("excel", inp.get("file"))
    names = [str(wb.Worksheets.Item(i).Name) for i in range(1, wb.Worksheets.Count + 1)]
    sheet = inp.get("sheet") or str(wb.ActiveSheet.Name)
    if sheet not in names:
        return f"No sheet named {sheet!r}. Sheets: {names}"
    ws = wb.Worksheets.Item(sheet)
    rng = ws.Range(inp["range"]) if inp.get("range") else ws.UsedRange
    rows, cols = int(rng.Rows.Count), int(rng.Columns.Count)
    note = ""
    if rows * cols > 3000:
        keep = max(1, 3000 // cols)
        rng = rng.Resize(keep, cols)
        note = f"\n... showing the first {keep} of {rows} rows"
    values, formulas = grid(rng.Value), grid(rng.Formula)
    out = [f"Workbook {wb.Name} | sheets: {names}", f"Sheet {sheet!r}, range {rng.Address}:"]
    for r, row in enumerate(values):
        cells = []
        for c, v in enumerate(row):
            f = formulas[r][c] if r < len(formulas) and c < len(formulas[r]) else ""
            cells.append(cell_text(v) + (f" [{f}]" if str(f).startswith("=") else ""))
        out.append(" | ".join(cells))
    return clip("\n".join(out) + note)


def excel_edit(inp):
    app, wb = find_item("excel", inp.get("file"))
    ops = inp.get("ops") or []
    if not ops:
        return "No ops given."
    if not confirm_edit("excel", wb, ops):
        return "The user declined. The workbook was not changed."
    report = []
    for op in ops:
        kind = op.get("op")
        if kind == "add_sheet":
            name = str(op.get("name") or "Sheet")[:31]
            ws = wb.Worksheets.Add(After=wb.Worksheets.Item(wb.Worksheets.Count))
            ws.Name = name
            report.append(f"added sheet {name!r}")
            continue
        sheet = op.get("sheet") or str(wb.ActiveSheet.Name)
        try:
            ws = wb.Worksheets.Item(sheet)
        except Exception:
            report.append(f"{kind}: no sheet named {sheet!r}")
            continue
        if kind == "write":
            values = op.get("values") or []
            if not values:
                report.append("write: no values given")
                continue
            start = ws.Range(op.get("cell") or "A1")
            r0, c0 = int(start.Row), int(start.Column)
            width = max(len(r) for r in values)
            data = tuple(
                tuple(("" if v is None else v) for v in (list(row) + [""] * (width - len(row))))
                for row in values
            )
            target = ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + len(values) - 1, c0 + width - 1))
            target.Formula = data
            report.append(f"wrote {len(values)} row(s) x {width} column(s) to {sheet}!{target.Address}")
        elif kind == "format":
            rng = ws.Range(op["range"]) if op.get("range") else ws.UsedRange
            if op.get("bold") is not None:
                rng.Font.Bold = bool(op["bold"])
            if op.get("number_format"):
                rng.NumberFormat = op["number_format"]
            if op.get("autofit"):
                rng.Columns.AutoFit()
            report.append(f"formatted {sheet}!{rng.Address}")
        elif kind == "chart":
            src = op.get("source")
            if not src:
                report.append("chart: 'source' range is required")
                continue
            anchor = ws.Range(op.get("anchor") or "H2")
            shape = ws.Shapes.AddChart2(-1, CHART_TYPES.get(op.get("type") or "column", 51),
                                        anchor.Left, anchor.Top, 380, 230)
            chart = shape.Chart
            chart.SetSourceData(ws.Range(src))
            if op.get("title"):
                chart.HasTitle = True
                chart.ChartTitle.Text = op["title"]
            report.append(f"added a {op.get('type') or 'column'} chart from {src} on {sheet}")
        else:
            report.append(f"unknown op {kind!r}")
    return f"Edited {wb.Name} (not saved yet).\n" + "\n".join(report)


# ---------- PowerPoint ----------
def ppt_layout(pres):
    layouts = pres.SlideMaster.CustomLayouts
    for i in range(1, layouts.Count + 1):
        if "title and content" in str(layouts.Item(i).Name).lower():
            return layouts.Item(i)
    return layouts.Item(min(2, layouts.Count))


def ppt_text_ranges(slide):
    for j in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes.Item(j)
        try:
            if shape.HasTextFrame:
                yield shape.TextFrame.TextRange
            if shape.HasTable:
                tbl = shape.Table
                for r in range(1, tbl.Rows.Count + 1):
                    for c in range(1, tbl.Columns.Count + 1):
                        yield tbl.Cell(r, c).Shape.TextFrame.TextRange
        except Exception:
            continue


def ppt_read(inp):
    app, pres = find_item("powerpoint", inp.get("file"))
    out = [f"{pres.Name} | {pres.Slides.Count} slides"]
    for i in range(1, pres.Slides.Count + 1):
        out.append(f"--- Slide {i} ---")
        for tr in ppt_text_ranges(pres.Slides.Item(i)):
            text = str(tr.Text).strip()
            if text:
                out.append(text.replace("\r", "\n"))
    return clip("\n".join(out))


def ppt_edit(inp):
    app, pres = find_item("powerpoint", inp.get("file"))
    ops = inp.get("ops") or []
    if not ops:
        return "No ops given."
    if not confirm_edit("powerpoint", pres, ops):
        return "The user declined. The presentation was not changed."
    report = []
    for op in ops:
        kind = op.get("op")
        if kind == "replace_text":
            find, repl, n = op.get("find", ""), op.get("replace", ""), 0
            for i in range(1, pres.Slides.Count + 1):
                for tr in ppt_text_ranges(pres.Slides.Item(i)):
                    after = 0
                    for _ in range(500):
                        hit = tr.Replace(find, repl, after)
                        if hit is None:
                            break
                        n += 1
                        after = int(hit.Start) + int(hit.Length) - 1
            report.append(f"replace {find!r}: {n} replacement(s)" if n else f"replace {find!r}: text not found")
        elif kind == "add_slide":
            slide = pres.Slides.AddSlide(pres.Slides.Count + 1, ppt_layout(pres))
            try:
                slide.Shapes.Title.TextFrame.TextRange.Text = op.get("title", "")
            except Exception:
                pass
            bullets = op.get("bullets") or []
            if bullets:
                try:
                    slide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "\r".join(str(b) for b in bullets)
                except Exception:
                    report.append("add_slide: the layout has no body placeholder, bullets were skipped")
            report.append(f"added slide {pres.Slides.Count}: {op.get('title')!r}")
        elif kind == "delete_slide":
            idx = int(op.get("index") or 0)
            if idx < 1 or idx > pres.Slides.Count:
                report.append(f"delete_slide: no slide {idx}; the deck has {pres.Slides.Count}")
                continue
            pres.Slides.Item(idx).Delete()
            report.append(f"deleted slide {idx}")
        else:
            report.append(f"unknown op {kind!r}")
    return f"Edited {pres.Name} (not saved yet).\n" + "\n".join(report)


# ---------- saving and email ----------
def save_document(inp):
    kind = inp["app"]
    if kind not in APPS:
        return "app must be word, excel or powerpoint."
    app, item = find_item(kind, inp.get("file"))
    if inp.get("export_pdf"):
        target = target_in_root(inp["export_pdf"], ".pdf")
        if not ask(f"export {item.Name} as a PDF:", [str(target.relative_to(ROOT))]):
            return "The user declined. Nothing was exported."
        if kind == "word":
            item.ExportAsFixedFormat(str(target), 17)
        elif kind == "excel":
            item.ExportAsFixedFormat(0, str(target))
        else:
            item.SaveCopyAs(str(target), 32)
        return f"Exported {target.relative_to(ROOT)}"
    if inp.get("save_as"):
        target = target_in_root(inp["save_as"], NEW_EXT[kind])
        if not ask(f"save {item.Name} under a new name:", [str(target.relative_to(ROOT))]):
            return "The user declined. Nothing was saved."
        if kind == "word":
            item.SaveAs2(str(target))
        elif kind == "excel":
            item.SaveAs(str(target), 51)
        else:
            item.SaveAs(str(target))
        return f"Saved as {target.relative_to(ROOT)}"
    if not str(item.Path):
        return "This document has never been saved, so it has no file yet. Call save_document again with save_as."
    if not ask(f"save your changes to {item.FullName}"):
        return "The user declined. Changes are not saved."
    item.Save()
    return f"Saved {item.Name}"


def outlook_draft(inp):
    attachments = [existing_path(a) for a in inp.get("attachments") or []]
    details = [f"To: {inp.get('to', '')}", f"Subject: {inp.get('subject', '')}"] + \
              [f"Attachment: {a.name}" for a in attachments]
    if not ask("open an Outlook draft (it will not be sent):", details):
        return "The user declined. No draft was opened."
    outlook = _com().Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = inp.get("to", "")
    mail.Subject = inp.get("subject", "")
    mail.Body = inp.get("body", "")
    for a in attachments:
        mail.Attachments.Add(str(a))
    mail.Display()
    return "Draft opened in Outlook. It has not been sent; the user reviews and sends it."


def existing_path(rel):
    p = safe(rel)
    if not p.is_file():
        raise ValueError(f"File not found in the working folder: {rel}")
    return p


HANDLERS = {
    "list_open": list_open,
    "list_files": list_files,
    "open_file": open_file,
    "create_file": create_file,
    "word_read": word_read,
    "word_edit": word_edit,
    "excel_read": excel_read,
    "excel_edit": excel_edit,
    "ppt_read": ppt_read,
    "ppt_edit": ppt_edit,
    "save_document": save_document,
    "outlook_draft": outlook_draft,
}


def run_tool(name, inp):
    """Run a tool, retrying when an Office app is busy and rejects the call."""
    last = None
    for attempt in range(5):
        try:
            return HANDLERS[name](inp)
        except Exception as e:
            last = e
            text = str(e)
            if "-2147418111" in text or "rejected" in text.lower():
                time.sleep(1.5)
                continue
            raise
    raise RuntimeError(f"Office stayed busy (is a dialog box open?): {last}")


# ---------- folder selection ----------
KNOWN = {
    "1": ("Documents", "Personal"),
    "2": ("Desktop", "Desktop"),
    "3": ("Downloads", "{374DE290-123F-4565-9164-39C4925E467B}"),
}


def known_folder(name, reg_value):
    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
        ) as k:
            value, _ = winreg.QueryValueEx(k, reg_value)
        return os.path.expandvars(value)
    except Exception:
        return str(Path.home() / name)


def pick_folder():
    import tkinter as tk
    from tkinter import filedialog

    win = tk.Tk()
    win.withdraw()
    win.attributes("-topmost", True)
    chosen = filedialog.askdirectory(title="Choose the working folder for Office files")
    win.destroy()
    return chosen


def choose_folder():
    print("Which folder should the agent open, create and save files in?")
    for key, (name, reg) in KNOWN.items():
        print(f"  {key}. {name}  ({known_folder(name, reg)})")
    print("  4. Another folder (browse)")
    while True:
        choice = input("Enter a number: ").strip()
        if choice in KNOWN:
            name, reg = KNOWN[choice]
            return known_folder(name, reg)
        if choice == "4":
            return pick_folder()
        print("Type a number from 1 to 4.")


def set_root(raw):
    global ROOT
    root = Path(raw.strip().strip('"')).resolve()
    if not root.is_dir():
        raise SystemExit(f"Not a folder: {root}")
    if len(root.parts) < 3 or any(p in {"Windows", "Program Files", "Program Files (x86)"} for p in root.parts):
        raise SystemExit("Pick a more specific folder, not a drive root or system folder.")
    ROOT = root


# ---------- main ----------
def main():
    if "--tools" in sys.argv:
        print(json.dumps(TOOLS, indent=2))
        return

    from anthropic import Anthropic

    client = Anthropic()

    if not AGENT_ID or AGENT_ID.startswith("PASTE"):
        agent = client.beta.agents.create(
            name="Office live assistant",
            model="claude-sonnet-5-5",
            system=SYSTEM,
            tools=TOOLS,
        )
        print(f"Created agent: {agent.id}")
        print("Paste it into AGENT_ID at the top of this file, then run the script again.")
        return

    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    raw = args[0] if args else choose_folder()
    if not raw:
        raise SystemExit("No folder chosen.")
    set_root(raw)

    session = client.beta.sessions.create(
        agent=AGENT_ID,
        environment_id=ENVIRONMENT_ID,
        title="Office live assistant",
    )
    print(f"\nWorking folder: {ROOT}")
    print("The agent controls your open Office apps. Type 'exit' to quit.")
    print("To undo a change, use Ctrl+Z in the app, or close the document without saving.")

    while True:
        prompt = input("\nYou: ").strip()
        if not prompt:
            continue
        if prompt.lower() in {"exit", "quit"}:
            break

        with client.beta.sessions.events.stream(session.id) as stream:
            client.beta.sessions.events.send(
                session.id,
                events=[{"type": "user.message", "content": [{"type": "text", "text": prompt}]}],
            )
            for event in stream:
                match event.type:
                    case "agent.message":
                        for block in event.content:
                            if block.type == "text":
                                print(block.text, end="")
                    case "agent.custom_tool_use":
                        try:
                            out = run_tool(event.name, event.input)
                        except Exception as e:
                            out = f"Error: {e}"
                        print(f"\n[{event.name}] {str(out)[:300]}")
                        client.beta.sessions.events.send(
                            session.id,
                            events=[{
                                "type": "user.custom_tool_result",
                                "custom_tool_use_id": event.id,
                                "content": [{"type": "text", "text": out or "(no output)"}],
                            }],
                        )
                    case "session.status_idle":
                        stop = getattr(event, "stop_reason", None)
                        if getattr(stop, "type", None) == "requires_action":
                            continue
                        print("\n")
                        break


if __name__ == "__main__":
    main()
