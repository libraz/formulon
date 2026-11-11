"""Windows Excel book-location probe: CELL("filename"/"address"), INFO("directory"), x15ac:absPath,
BrtAbsPath15 (xlsb), and local-hyperlink rel Target / display across Save and Save As.

Run on the Windows host with Windows python (pywin32 installed; xlwings is not needed):

    py -3 tools\\oracle\\win_probes\\book_location.py C:\\fm_out

It starts its own hidden Excel instance (DispatchEx), so open user workbooks are never touched,
and always quits that instance. It writes only under C:\\fm (its scratch folder, created by this
script and marked with .formulon_probe; it refuses to run if C:\\fm exists without the marker) and
under the output folder. UNC cases go through the admin share \\\\localhost\\c$\\fm, i.e. the same
C:\\fm folder; a UNC failure is recorded and does not stop the run.

The cell context menu is read from Application.CommandBars ("Cell", "List Range Popup",
"Table List Popup") through COM: no menu is opened and no input is sent. That is the classic
control set and may differ from the rendered menu; a UIA read would need a real right-click.

Results go to <output folder>/results.json, next to Excel-saved copies of every book.
"""

import ctypes
import json
import os
import re
import shutil
import sys
import time
import traceback
import zipfile

import pythoncom
import win32com.client

SCRATCH = r"C:\fm"
MARKER = os.path.join(SCRATCH, ".formulon_probe")
XL_XLSX, XL_XLSB = 51, 50

# Book locations: (tag, folder for the first save, folder for the Save As).
BOOKS = [
    ("drive", r"C:\fm\a b", r"C:\fm\second"),
    ("unc", r"\\localhost\c$\fm", r"\\localhost\c$\fm\second u"),
]
FORMULAS = {
    "B1": '=CELL("filename",A1)',
    "B2": '=CELL("filename",Sheet2!A1)',
    "B3": '=CELL("filename")',
    "B4": '=CELL("address",Sheet2!A1)',
    "B5": "=CELL(\"address\",'S 2'!A1)",
    "B6": '=CELL("address",Sheet1!A1)',
    "B7": '=INFO("directory")',
}


def other_drive():
    """First existing drive letter other than C:, else D: (link targets need not exist)."""
    mask = ctypes.windll.kernel32.GetLogicalDrives()
    for i in range(26):
        letter = chr(ord("A") + i)
        if letter not in "ABC" and mask & (1 << i):
            return letter + ":"
    return "D:"


def link_targets():
    d = other_drive()
    return {
        "C10": r"C:\fm\a b\sub\f.xlsx",  # below the drive book's folder
        "C11": r"C:\fm\other\g.xlsx",  # sibling folder (shares C:\fm)
        "C12": r"C:\Temp\h.xlsx",  # shares only the drive root
        "C13": d + r"\x\i.xlsx",  # another drive
        "C14": r"\\localhost\c$\fm\sub\j.xlsx",  # UNC: shared with the UNC book, not with the drive book
        "C15": r"\\server\share\k.xlsx",  # UNC on another server
        "C16": r"C:\fm\a b\a%b.xlsx",  # stray percent
        "C17": r"C:\fm\a b\a%20b.xlsx",  # literal percent-two-zero
        "C18": r"C:\fm\a b\a#b.xlsx",  # hash
        "C19": r"C:\fm\a b\日本 x.xlsx",  # non-ASCII and space
    }


def err(e):
    return "ERR " + "".join(traceback.format_exception_only(type(e), e)).strip()[:300]


def safe(fn):
    try:
        return fn()
    except Exception as e:  # noqa: BLE001 - every COM failure is a result to record
        return err(e)


def remove_exact(path):
    if not path.lower().startswith(("c:\\fm\\", "\\\\localhost\\c$\\fm\\")):
        raise RuntimeError("refusing to delete outside the scratch folder: " + path)
    if os.path.exists(path):
        os.remove(path)


def read_cells(ws, refs):
    out = {}
    for c in refs:
        r = ws.Range(c)
        out[c] = {"formula": safe(lambda: r.Formula), "value": safe(lambda: r.Value), "text": safe(lambda: r.Text)}
    return out


def read_links(ws, refs):
    out = {}
    for c in refs:

        def one():
            r = ws.Range(c)
            if r.Hyperlinks.Count == 0:
                return "no hyperlink"
            h = r.Hyperlinks(1)
            return {
                "address": h.Address,
                "sub_address": h.SubAddress,
                "text_to_display": safe(lambda: h.TextToDisplay),
                "cell_text": r.Text,
            }

        out[c] = safe(one)
    return out


def xlsx_parts(path):
    z = zipfile.ZipFile(path)
    wbx = z.read("xl/workbook.xml").decode("utf-8")
    names = z.namelist()
    sheets = {}
    for n in sorted(names):
        m = re.match(r"xl/worksheets/(sheet\d+)\.xml$", n)
        if not m:
            continue
        sx = z.read(n).decode("utf-8")
        rel_name = "xl/worksheets/_rels/%s.xml.rels" % m.group(1)
        rels = z.read(rel_name).decode("utf-8") if rel_name in names else ""
        hl = re.findall(r"<hyperlink [^>]*/>", sx)
        if hl or rels:
            sheets[m.group(1)] = {"hyperlink_elements": hl, "rels_raw": rels}
    return {"absPath": re.findall(r"<x15ac:absPath [^>]*/>", wbx), "sheets": sheets}


def read_varint(data, i, max_bytes):
    """MS-XLSB record header field: 7 bits per byte, high bit = continuation."""
    v = 0
    for s in range(max_bytes):
        b = data[i]
        i += 1
        v |= (b & 0x7F) << (7 * s)
        if not b & 0x80:
            break
    return v, i


def xlsb_records(data):
    i, out = 0, []
    while i < len(data):
        t, i = read_varint(data, i, 2)
        n, i = read_varint(data, i, 4)
        out.append((t, data[i : i + n]))
        i += n
    return out


def xlsb_abspath(path):
    recs = xlsb_records(zipfile.ZipFile(path).read("xl/workbook.bin"))
    hits = []
    for idx, (t, pl) in enumerate(recs):
        if t == 0x817:
            cch = int.from_bytes(pl[:4], "little") if len(pl) >= 4 else None
            hits.append(
                {
                    "index": idx,
                    "prev": hex(recs[idx - 1][0]),
                    "next": hex(recs[idx + 1][0]) if idx + 1 < len(recs) else None,
                    "hex": pl.hex(),
                    "string": pl[4 : 4 + 2 * cch].decode("utf-16le", "replace") if cch else None,
                }
            )
    return {"record_seq_head": [hex(t) for t, _ in recs[:16]], "BrtAbsPath15": hits}


def new_book(xl):
    wb = xl.Workbooks.Add()
    while wb.Worksheets.Count < 3:
        wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    wb.Worksheets(1).Name = "Sheet1"
    wb.Worksheets(2).Name = "Sheet2"
    wb.Worksheets(3).Name = "S 2"
    ws = wb.Worksheets("Sheet1")
    ws.Activate()
    for c, f in FORMULAS.items():
        ws.Range(c).Formula = f
    return wb


def run_book(xl, tag, folder1, folder2, out_dir, R):
    res = R.setdefault(tag, {})
    p1 = os.path.join(folder1, "book.xlsx")
    p2 = os.path.join(folder2, "book.xlsx")
    for f in (folder1, folder2):
        os.makedirs(f, exist_ok=True)
    for p in (p1, p2, os.path.join(folder1, "book.xlsb")):
        remove_exact(p)
    links = link_targets()
    res["paths"] = {"save": p1, "save_as": p2}
    wb = new_book(xl)
    ws = wb.Worksheets("Sheet1")
    xl.CalculateFull()
    res["cells_unsaved"] = read_cells(ws, FORMULAS)
    wb.SaveAs(p1, XL_XLSX)
    res["fullname_after_saveas"] = wb.FullName
    res["cells_saved_before_recalc"] = read_cells(ws, FORMULAS)
    xl.CalculateFull()
    res["cells_saved"] = read_cells(ws, FORMULAS)
    # Same no-ref CELL with another sheet active, after a full recalc.
    wb.Worksheets("Sheet2").Activate()
    xl.CalculateFull()
    res["cells_saved_sheet2_active"] = read_cells(ws, FORMULAS)
    ws.Activate()
    shutil.copy(p1, os.path.join(out_dir, tag + "_1_nolinks.xlsx"))
    res["file_1_nolinks"] = xlsx_parts(p1)
    add = {}
    for c, t in links.items():
        add[c] = safe(lambda: ws.Hyperlinks.Add(Anchor=ws.Range(c), Address=t) and "ok")
    res["links_in"] = links
    res["links_add_result"] = add
    res["links_before_save"] = read_links(ws, links)
    wb.Save()
    res["links_after_save"] = read_links(ws, links)
    shutil.copy(p1, os.path.join(out_dir, tag + "_2_saved.xlsx"))
    res["file_2_saved"] = xlsx_parts(p1)
    wb.Close(False)
    wb = xl.Workbooks.Open(p1)
    ws = wb.Worksheets("Sheet1")
    res["links_after_reopen"] = read_links(ws, links)
    xl.CalculateFull()
    res["cells_after_reopen"] = read_cells(ws, FORMULAS)
    wb.SaveAs(p2, XL_XLSX)
    res["links_after_saveas"] = read_links(ws, links)
    xl.CalculateFull()
    res["cells_after_saveas"] = read_cells(ws, FORMULAS)
    shutil.copy(p2, os.path.join(out_dir, tag + "_3_saveas.xlsx"))
    res["file_3_saveas"] = xlsx_parts(p2)
    wb.Close(False)
    # xlsb: BrtAbsPath15 for the first folder.
    wb = new_book(xl)
    pb = os.path.join(folder1, "book.xlsb")
    wb.SaveAs(pb, XL_XLSB)
    wb.Close(False)
    shutil.copy(pb, os.path.join(out_dir, tag + "_4.xlsb"))
    res["xlsb"] = xlsb_abspath(pb)


def cell_menu(xl):
    out = {}
    for name in ("Cell", "List Range Popup", "Table List Popup"):

        def one():
            bar = xl.CommandBars(name)
            items = []
            for i in range(1, bar.Controls.Count + 1):
                c = bar.Controls(i)
                items.append(
                    {
                        "index": i,
                        "id": c.Id,
                        "caption": c.Caption,
                        "visible": c.Visible,
                        "enabled": c.Enabled,
                        "begin_group": c.BeginGroup,
                        "type": c.Type,
                    }
                )
            return items

        out[name] = safe(one)
    out["note"] = (
        "COM CommandBars lists the classic control set; the rendered right-click menu in "
        "Excel 365 may differ (UIA read skipped: it needs a real right-click)."
    )
    return out


def main():
    if len(sys.argv) != 2:
        sys.exit("usage: py -3 book_location.py <output folder, e.g. C:\\fm_out>")
    out_dir = os.path.abspath(sys.argv[1])
    if os.path.isdir(SCRATCH) and not os.path.exists(MARKER):
        sys.exit(SCRATCH + " exists and is not this probe's scratch folder; move it away or pick another host.")
    os.makedirs(SCRATCH, exist_ok=True)
    open(MARKER, "w").close()
    os.makedirs(out_dir, exist_ok=True)
    R = {"script": "tools/oracle/win_probes/book_location.py", "started": time.strftime("%Y-%m-%dT%H:%M:%S")}
    pythoncom.CoInitialize()
    xl = win32com.client.DispatchEx("Excel.Application")
    try:
        xl.Visible = False
        xl.DisplayAlerts = False
        R["env"] = {
            "Version": safe(lambda: xl.Version),
            "Build": safe(lambda: xl.Build),
            "OperatingSystem": safe(lambda: xl.OperatingSystem),
            "country": safe(lambda: xl.International(1)),
            "ui_lang": safe(lambda: xl.LanguageSettings.LanguageID(2)),
            "DefaultFilePath": safe(lambda: xl.DefaultFilePath),
            "cwd_python": os.getcwd(),
        }
        for tag, f1, f2 in BOOKS:
            try:
                run_book(xl, tag, f1, f2, out_dir, R)
            except Exception as e:  # noqa: BLE001
                R.setdefault(tag, {})["fatal"] = err(e)
                for i in range(xl.Workbooks.Count, 0, -1):
                    safe(lambda: xl.Workbooks(i).Close(False))
        R["cell_menu"] = cell_menu(xl)
    finally:
        for i in range(xl.Workbooks.Count, 0, -1):
            safe(lambda: xl.Workbooks(i).Close(False))
        xl.Quit()
        pythoncom.CoUninitialize()
        R["finished"] = time.strftime("%Y-%m-%dT%H:%M:%S")
        with open(os.path.join(out_dir, "results.json"), "w", encoding="utf-8") as fh:
            json.dump(R, fh, ensure_ascii=False, indent=1, default=str)
        print("wrote", os.path.join(out_dir, "results.json"))


if __name__ == "__main__":
    main()
