"""
Documentation search helpers — scan AlphaCAM install dir + local chm/ for HTML docs.
"""

from __future__ import annotations

import os
import re
import logging

from .config import detect_alphacam_dir, CHM_DOCS

log = logging.getLogger("alphacam-bridge.docs")

# Project root directory (where the chm/ folder might be)
_skill_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# ---------------------------------------------------------------------------
# Doc root discovery
# ---------------------------------------------------------------------------

def _get_doc_roots() -> list[str]:
    """Return directories to search for documentation HTML files."""
    roots: list[str] = []

    # 1. AlphaCAM install dir — search tempacamapi and any _html folders
    acam_dir = detect_alphacam_dir()
    if acam_dir:
        tempacamapi = os.path.join(acam_dir, "tempacamapi")
        if os.path.isdir(tempacamapi):
            roots.append(tempacamapi)
        for entry in os.listdir(acam_dir):
            full = os.path.join(acam_dir, entry)
            if entry.endswith("_html") and os.path.isdir(full):
                if full not in roots:
                    roots.append(full)

    # 2. Project-level chm/ folder (fallback for offline / extracted docs)
    project_chm = os.path.join(_skill_dir, "chm")
    if os.path.isdir(project_chm):
        for entry in os.listdir(project_chm):
            full = os.path.join(project_chm, entry)
            if entry.endswith("_html") and os.path.isdir(full):
                if full not in roots:
                    roots.append(full)

    return roots


def _is_doc_file(name: str) -> bool:
    return name.lower().endswith(".htm") or name.lower().endswith(".html")


# ---------------------------------------------------------------------------
# Categories / listing
# ---------------------------------------------------------------------------

def _get_doc_categories() -> dict[str, int]:
    """Return doc source directories with file counts."""
    result: dict[str, int] = {}
    for root in _get_doc_roots():
        count = 0
        for _r, _dirs, files in os.walk(root):
            count += sum(1 for f in files if _is_doc_file(f))
        if count > 0:
            label = os.path.basename(root)
            result[label] = count
    return result


# ---------------------------------------------------------------------------
# Find / search helpers
# ---------------------------------------------------------------------------

def _find_doc_file(name: str) -> str | None:
    """Find a doc HTML file by name (case-insensitive, partial match)."""
    name_lower = name.lower()
    if not name_lower.endswith(".htm"):
        name_lower += ".htm"

    roots = _get_doc_roots()

    # First pass: exact match
    for root in roots:
        for r, _dirs, files in os.walk(root):
            for f in files:
                if f.lower() == name_lower:
                    return os.path.join(r, f)

    # Second pass: partial match
    for root in roots:
        for r, _dirs, files in os.walk(root):
            for f in files:
                if _is_doc_file(f) and name_lower in f.lower():
                    return os.path.join(r, f)

    return None


def _strip_html(html: str) -> str:
    """Crude HTML-to-text conversion: strip tags, decode entities."""
    text = re.sub(r'<script[^>]*>.*?</script>', '', html,
                  flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'<style[^>]*>.*?</style>', '', text,
                  flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'</?(?:p|div|br|tr|li|h\d|table|section)[^>]*>', '\n',
                  text, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', '', text)
    text = text.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')
    text = text.replace('&nbsp;', ' ').replace('&#160;', ' ')
    text = re.sub(r'\n{3,}', '\n\n', text)
    text = re.sub(r'&#(\d+);', lambda m: chr(int(m.group(1))), text)
    return text.strip()


def _get_doc_title(filepath: str) -> str:
    """Extract the <title> from an HTML file."""
    try:
        with open(filepath, "r", encoding="utf-8", errors="replace") as f:
            content = f.read(4096)
        m = re.search(r'<title[^>]*>(.*?)</title>', content,
                      re.IGNORECASE | re.DOTALL)
        if m:
            return _strip_html(m.group(1))
    except Exception:
        pass
    return os.path.basename(filepath).replace(".htm", "").replace("_", " ")


def _search_docs(query: str, max_results: int = 20) -> list[dict]:
    """Search doc page filenames and titles for a query string across all doc roots."""
    q = query.lower()
    results = []
    for root in _get_doc_roots():
        for r, _dirs, files in os.walk(root):
            for f in files:
                if not _is_doc_file(f):
                    continue
                filepath = os.path.join(r, f)
                rel_path = os.path.relpath(filepath, root)
                score = 0
                if q in f.lower():
                    score += 2
                title = _get_doc_title(filepath)
                if q in title.lower():
                    score += 1
                if score > 0:
                    results.append({
                        "file": f,
                        "title": title,
                        "path": rel_path,
                        "source": os.path.basename(root),
                        "score": score,
                    })
    results.sort(key=lambda x: -x["score"])
    return results[:max_results]


# ---------------------------------------------------------------------------
# Handlers (synchronous, called from dispatcher)
# ---------------------------------------------------------------------------

def handle_list_docs() -> dict:
    """Handle the list_docs tool."""
    roots = _get_doc_roots()
    categories = _get_doc_categories()
    acam_dir = detect_alphacam_dir()
    return {
        "alphacam_install_dir": acam_dir or "(not detected)",
        "doc_search_roots": roots,
        "doc_categories": categories,
        "total_html_files": sum(categories.values()),
        "chm_docs": {k: v["desc"] for k, v in CHM_DOCS.items()},
        "tip": "Use read_doc(name='Path_TrimWithCuttingGeos') to read a specific API page. "
                "Use search_docs(query='offset') to find matching pages.",
    }


def handle_read_doc(name: str) -> dict:
    """Handle the read_doc tool."""
    filepath = _find_doc_file(name)
    if not filepath:
        raise FileNotFoundError(
            f"Document '{name}' not found. Use search_docs() to find matching pages."
        )
    with open(filepath, "r", encoding="utf-8", errors="replace") as f:
        html = f.read()
    text = _strip_html(html)
    # Show path relative to the closest doc root
    roots = _get_doc_roots()
    rel_path = filepath
    for root in roots:
        try:
            candidate = os.path.relpath(filepath, root)
            if not candidate.startswith(".."):
                rel_path = candidate
                break
        except Exception:
            pass
    title = _get_doc_title(filepath)
    MAX_LEN = 8000
    if len(text) > MAX_LEN:
        text = text[:MAX_LEN] + f"\n\n... [truncated, full length: {len(text)} chars]"
    return {
        "title": title,
        "file": os.path.basename(filepath),
        "path": rel_path,
        "content": text,
    }


def handle_search_docs(query: str) -> dict:
    """Handle the search_docs tool."""
    results = _search_docs(query)
    if not results:
        return {
            "query": query,
            "count": 0,
            "results": [],
            "tip": "Try a different keyword, or use list_docs() to browse categories.",
        }
    return {
        "query": query,
        "count": len(results),
        "results": [
            {"file": r["file"], "title": r["title"],
             "path": r["path"], "source": r.get("source", "")}
            for r in results
        ],
    }


async def handle_convert_chm_to_html(chm_path: str,
                                     output_dir: str | None = None) -> dict:
    """Convert a .chm file to HTML using hh.exe decompile."""
    import shutil
    import subprocess
    import tempfile
    import asyncio

    if not os.path.isfile(chm_path):
        raise FileNotFoundError(f"CHM file not found: {chm_path}")
    if not chm_path.lower().endswith(".chm"):
        raise ValueError(f"File is not a .chm file: {chm_path}")

    if output_dir is None:
        base = os.path.splitext(chm_path)[0]
        output_dir = base + "_html"
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    chm_abspath = os.path.abspath(chm_path)
    with tempfile.TemporaryDirectory(prefix="chm_extract_") as tmp_dir:
        tmp_chm = os.path.join(tmp_dir, "source.chm")
        shutil.copy2(chm_abspath, tmp_chm)

        cmd = ["hh.exe", "-decompile", output_dir, tmp_chm]
        log.info(f"Running: {' '.join(cmd)}")

        try:
            proc = await asyncio.create_subprocess_exec(
                *cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            )
            stdout, stderr = await asyncio.wait_for(
                proc.communicate(), timeout=120)
        except asyncio.TimeoutError:
            raise RuntimeError(
                "hh.exe timed out after 120 seconds. The .chm file may be "
                "corrupted or very large.")

        if proc.returncode != 0:
            error_msg = stderr.decode("utf-8", errors="replace").strip()
            raise RuntimeError(
                f"hh.exe failed (exit code {proc.returncode}): {error_msg}")

    # Collect extracted files
    extracted = []
    total_size = 0
    for root, _dirs, files in os.walk(output_dir):
        for f in sorted(files):
            filepath = os.path.join(root, f)
            size = os.path.getsize(filepath)
            total_size += size
            rel_path = os.path.relpath(filepath, output_dir)
            extracted.append({"path": rel_path, "size": size})

    return {
        "output_dir": output_dir,
        "chm_file": os.path.abspath(chm_path),
        "file_count": len(extracted),
        "total_size_bytes": total_size,
        "files": extracted,
    }
