# -*- coding: utf-8 -*-
"""docx_edit.py — reusable helpers for editing .docx 施工方案 (OOXML direct edit).

Usage: import these functions in a throwaway script, or run modules directly:
  python docx_edit.py dump <file.docx>            # dump text to stdout/file
  python docx_edit.py verify <modified.docx> <orig.docx>  # run verification checklist
"""
import sys
import os
import re
import zipfile
from copy import deepcopy
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
T = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


# ---------------- reading ----------------

def load_doc(path):
    """Return (root, zipfile-namelist). Raise on non-zip."""
    with open(path, "rb") as f:
        head = f.read(4)
    if head[:2] != b"PK":
        raise ValueError("not a zip/docx file (check file header; .doc may be OLE)")
    with zipfile.ZipFile(path) as z:
        return etree.fromstring(z.read("word/document.xml")), z.namelist()


def para_text(p):
    return "".join(t.text or "" for t in p.findall(".//{%s}t" % W))


def direct_text(p):
    """Text of w:t directly inside this paragraph's runs (excludes nested drawings/textboxes)."""
    return "".join(t.text or "" for t in p.findall("./{%s}r/{%s}t" % (W, W)))


def para_style(p):
    pPr = p.find("{%s}pPr" % W)
    if pPr is None:
        return ""
    ps = pPr.find("{%s}pStyle" % W)
    return ps.get("{%s}val" % W) if ps is not None else ""


def dump_docx_text(path, out=None):
    """Dump '[index] <style> text' lines. out=None -> return list, else write file."""
    root, _ = load_doc(path)
    lines = []
    for i, p in enumerate(root.iter("{%s}p" % W)):
        t = para_text(p)
        if t.strip():
            lines.append(f"[{i}] <{para_style(p)}> {t}")
    if out:
        with open(out, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        return len(lines)
    return lines


# ---------------- editing ----------------

def set_text(p, new_text):
    ts = p.findall(".//{%s}t" % W)
    if not ts:
        r = etree.SubElement(p, "{%s}r" % W)
        t = etree.SubElement(r, "{%s}t" % W)
        t.text = new_text
        return
    ts[0].text = new_text
    for t in ts[1:]:
        t.text = None


def replace_in_para(p, old, new):
    """Paragraph-level replace across runs. DANGER: uses .//w:t (includes nested
    drawing/textbox text). For paragraphs containing drawings use replace_direct."""
    ts = p.findall(".//{%s}t" % W)
    merged = "".join(t.text or "" for t in ts)
    if old not in merged:
        return False
    merged = merged.replace(old, new)
    ts[0].text = merged
    for t in ts[1:]:
        t.text = None
    return True


def replace_direct(p, old, new):
    """Replace only in direct run text — safe for paragraphs that embed drawings."""
    ts = p.findall("./{%s}r/{%s}t" % (W, W))
    merged = "".join(t.text or "" for t in ts)
    if old not in merged:
        return False
    merged = merged.replace(old, new)
    ts[0].text = merged
    for t in ts[1:]:
        t.text = None
    return True


def find_para(root, key, contains=True, direct=False):
    for p in root.iter("{%s}p" % W):
        t = direct_text(p) if direct else para_text(p)
        if (key in t) if contains else (t.strip() == key):
            return p
    return None


def find_all_paras(root, key, direct=False):
    out = []
    for p in root.iter("{%s}p" % W):
        t = direct_text(p) if direct else para_text(p)
        if key in t:
            out.append(p)
    return out


def remove_para(p):
    parent = p.getparent()
    if parent is not None:
        parent.remove(p)


def set_table_cell(tc, text, para_idx=0):
    ps = tc.findall("{%s}p" % W)
    set_text(ps[para_idx], text)


def delete_table_row_by_text(tbl, key, col=0):
    for tr in tbl.findall("{%s}tr" % W):
        tcs = tr.findall("{%s}tc" % W)
        if len(tcs) > col and key in para_text(tcs[col]):
            tr.getparent().remove(tr)
            return True
    return False


def insert_row_copy(anchor_tr, cells_texts):
    """Deep-copy anchor row, fill cells with texts (empty string -> clear), insert after."""
    new_tr = deepcopy(anchor_tr)
    tcs = new_tr.findall("{%s}tc" % W)
    for i, txt in enumerate(cells_texts):
        if i < len(tcs):
            set_text(tcs[i].findall("{%s}p" % W)[0], txt)
    anchor_tr.addnext(new_tr)
    return new_tr


def make_para(text, style=None):
    p = etree.Element("{%s}p" % W)
    if style:
        pPr = etree.SubElement(p, "{%s}pPr" % W)
        ps = etree.SubElement(pPr, "{%s}pStyle" % W)
        ps.set("{%s}val" % W, style)
    r = etree.SubElement(p, "{%s}r" % W)
    t = etree.SubElement(r, "{%s}t" % W)
    t.text = text
    return p


# ---------------- numbering simulation ----------------

def simulate_heading_numbers(root, style_lvl):
    """Yield (rendered_num, level, style, text) for heading paragraphs in doc order.
    style_lvl: dict styleId -> level (e.g. {'1':0,'2':1,'3':2,'4':3})."""
    counters = [0] * 9
    for p in root.iter("{%s}p" % W):
        st = para_style(p)
        if st not in style_lvl:
            continue
        t = para_text(p)
        if not t.strip():
            continue
        lvl = style_lvl[st]
        counters[lvl] += 1
        for d in range(lvl + 1, 9):
            counters[d] = 0
        num = ".".join(str(counters[d]) for d in range(0, lvl + 1))
        yield num, lvl, st, t


# ---------------- verification ----------------

def verify_modified(dst_path, src_path=None, checks=None):
    """Run verification checklist. checks: list of (key, expected_count) on visible text.
    Returns list of problem strings (empty = all good)."""
    problems = []
    with zipfile.ZipFile(dst_path) as z:
        doc = z.read("word/document.xml").decode("utf-8")
        root = etree.fromstring(doc.encode("utf-8"))
        rels = z.read("word/_rels/document.xml.rels").decode("utf-8")
        dst_names = set(z.namelist())
    try:
        etree.fromstring(doc.encode("utf-8"))
    except Exception as e:
        problems.append(f"XML not well-formed: {e}")
    vis = "".join(para_text(p) for p in root.iter("{%s}p" % W))
    for key, expect in (checks or []):
        n = vis.count(key)
        if n != expect:
            problems.append(f"text {key!r}: found {n}, expected {expect}")
    refs = set(re.findall(r'r:(?:embed|id)="(rId\d+)"', doc)) | set(re.findall(r'r:id="(rId\d+)"', doc))
    rmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    missing = refs - set(rmap)
    if missing:
        problems.append(f"undefined rIds: {missing}")
    if any(t == "../NULL" for t in rmap.values()):
        problems.append("NULL rel targets present")
    try:
        import docx
        docx.Document(dst_path)
    except Exception as e:
        problems.append(f"python-docx open failed: {e}")
    if src_path:
        with zipfile.ZipFile(src_path) as z:
            src_media = {n for n in z.namelist() if n.startswith("word/media/")}
        dst_media = {n for n in dst_names if n.startswith("word/media/")}
        if src_media != dst_media:
            problems.append("media files changed (images were touched!)")
    return problems


def write_doc(root, src_path, dst_path):
    """Rebuild docx: copy all parts, replace document.xml."""
    tmp = dst_path + ".tmp"
    with zipfile.ZipFile(src_path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            data = zin.read(name)
            if name == "word/document.xml":
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(name, data)
    os.replace(tmp, dst_path)


# ---------------- CLI ----------------

if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "dump":
        n = dump_docx_text(sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
        print(f"dumped {n} paragraphs")
    elif len(sys.argv) >= 3 and sys.argv[1] == "verify":
        probs = verify_modified(sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
        if probs:
            print("PROBLEMS:")
            for p in probs:
                print(" -", p)
        else:
            print("ALL OK")
    else:
        print(__doc__)
