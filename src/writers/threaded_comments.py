"""Keep Excel's modern threaded comments through an openpyxl load/save.

openpyxl only knows the older "notes".  For every threaded comment Excel also
stores a note whose author is ``tc={thread id}``; openpyxl keeps that note but
drops the threaded-comment parts, so the comment vanishes from Excel's
Comments pane.  This module reads the threads from the source workbook and,
after the workbook is saved, writes them back next to the matching notes (at
the note's current cell, so rows the writer shifted stay attached).
"""
import os
import re
import zipfile
from html import unescape
from posixpath import normpath, join, dirname

_THREAD_REL = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
_PERSON_REL = "http://schemas.microsoft.com/office/2017/10/relationships/person"
_THREAD_CT = "application/vnd.ms-excel.threadedcomments+xml"
_PERSON_CT = "application/vnd.ms-excel.person+xml"

_COMMENT_RE = re.compile(r"<threadedComment\b[^>]*?(?:/>|>.*?</threadedComment>)", re.S)


def thread_key(author) -> str | None:
    """'tc={GUID}' (the note author Excel gives a threaded comment) -> normalised GUID."""
    if not author or not str(author).startswith("tc="):
        return None
    return str(author)[3:].strip("{} ").lower()


def _attr(tag: str, name: str) -> str | None:
    m = re.search(rf'\s{name}="([^"]*)"', tag)
    return unescape(m.group(1)) if m else None


def _resolve(base_dir: str, target: str) -> str:
    return normpath(target.lstrip("/")) if target.startswith("/") else normpath(join(base_dir, target))


def _sheet_parts(zf: zipfile.ZipFile) -> dict[str, str]:
    """{sheet name: part path} for a workbook zip."""
    wb = zf.read("xl/workbook.xml").decode("utf8")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf8")
    targets = {_attr(t, "Id"): _attr(t, "Target") for t in re.findall(r"<Relationship\b[^>]*>", rels)}
    out = {}
    for tag in re.findall(r"<sheet\b[^>]*>", wb):
        rid = re.search(r'\sr:id="([^"]*)"', tag)
        if rid and targets.get(rid.group(1)):
            out[_attr(tag, "name")] = _resolve("xl", targets[rid.group(1)])
    return out


def read_threads(path: str) -> dict:
    """Threads in a source workbook: {'threads': {guid: [xml, ...]}, 'root': str, 'persons': bytes}.

    Each xml string is one <threadedComment> element (top-level first, then its
    replies) with the cell reference replaced by the placeholder ``{REF}``.
    Returns {} when the workbook has no threaded comments.
    """
    try:
        zf = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile):
        return {}
    with zf:
        names = zf.namelist()
        parts = [n for n in names if n.startswith("xl/threadedComments/") and n.endswith(".xml")]
        if not parts:
            return {}
        threads: dict[str, list[str]] = {}
        root = None
        for part in parts:
            xml = zf.read(part).decode("utf8")
            if root is None:
                root = re.search(r"<ThreadedComments\b[^>]*>", xml).group(0)
            tops, replies = {}, {}
            for raw in _COMMENT_RE.findall(xml):
                tag = re.match(r"<threadedComment\b[^>]*", raw).group(0)
                cid = (_attr(tag, "id") or "").strip("{} ").lower()
                parent = _attr(tag, "parentId")
                raw = re.sub(r'(\sref=")[^"]*(")', r"\1{REF}\2", raw, count=1)
                if parent:
                    replies.setdefault(parent.strip("{} ").lower(), []).append(raw)
                else:
                    tops[cid] = raw
            for cid, raw in tops.items():
                threads[cid] = [raw] + replies.get(cid, [])
        persons = next((zf.read(n) for n in names if n.startswith("xl/persons/") and n.endswith(".xml")), None)
        return {"threads": threads, "root": root, "persons": persons}


def restore_threads(output_path: str, source: dict, wanted: dict[str, list[tuple[str, str]]]) -> int:
    """Write threaded comments into the saved workbook.

    wanted: {sheet name: [(cell coordinate, thread key), ...]} for every note in
    the saved workbook that came from a threaded comment.  Returns how many
    threads were restored.
    """
    if not source or not wanted:
        return 0
    with zipfile.ZipFile(output_path) as zf:
        files = {n: zf.read(n) for n in zf.namelist()}
        order = zf.namelist()
        sheet_parts = _sheet_parts(zf)

    content_types = files["[Content_Types].xml"].decode("utf8")
    wb_rels = files["xl/_rels/workbook.xml.rels"].decode("utf8")
    added = 0
    new_parts = 0

    for sheet, cells in wanted.items():
        part = sheet_parts.get(sheet)
        body = []
        for coord, key in cells:
            for raw in source["threads"].get(key, []):
                body.append(raw.replace("{REF}", coord))
            added += 1 if source["threads"].get(key) else 0
        if not part or not body:
            continue
        new_parts += 1
        tc_part = f"xl/threadedComments/threadedComment{new_parts}.xml"
        files[tc_part] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            + source["root"] + "".join(body) + "</ThreadedComments>"
        ).encode("utf8")
        order.append(tc_part)
        content_types = content_types.replace(
            "</Types>", f'<Override PartName="/{tc_part}" ContentType="{_THREAD_CT}"/></Types>')

        rels_part = f"{dirname(part)}/_rels/{part.rsplit('/', 1)[-1]}.rels"
        rel = f'<Relationship Id="rIdTC{new_parts}" Type="{_THREAD_REL}" Target="/{tc_part}"/>'
        if rels_part in files:
            files[rels_part] = files[rels_part].decode("utf8").replace("</Relationships>", rel + "</Relationships>").encode("utf8")
        else:
            files[rels_part] = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                + rel + "</Relationships>").encode("utf8")
            order.append(rels_part)

    if not new_parts:
        return 0

    if source.get("persons") and "xl/persons/person.xml" not in files:
        files["xl/persons/person.xml"] = source["persons"]
        order.append("xl/persons/person.xml")
        content_types = content_types.replace(
            "</Types>", f'<Override PartName="/xl/persons/person.xml" ContentType="{_PERSON_CT}"/></Types>')
        wb_rels = wb_rels.replace(
            "</Relationships>",
            f'<Relationship Id="rIdPersons" Type="{_PERSON_REL}" Target="/xl/persons/person.xml"/></Relationships>')

    files["[Content_Types].xml"] = content_types.encode("utf8")
    files["xl/_rels/workbook.xml.rels"] = wb_rels.encode("utf8")

    tmp = output_path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name in order:
            out.writestr(name, files[name])
    os.replace(tmp, output_path)
    return added
