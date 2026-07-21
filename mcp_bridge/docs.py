"""
Documentation search helpers — scan AlphaCAM install dir + local chm/ for HTML docs.
Includes content-snippet preview, CHM_DOCS awareness, and search result cache.
"""

from __future__ import annotations

import os
import re
import time
import logging
from typing import Any

from .config import detect_alphacam_dir, CHM_DOCS

log = logging.getLogger("alphacam-bridge.docs")

# Project root directory (where the chm/ folder might be)
_skill_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ---------------------------------------------------------------------------
# Simple in-memory search cache (TTL = 60 seconds)
# ---------------------------------------------------------------------------
_search_cache: dict[str, tuple[float, list[dict]]] = {}
_SEARCH_CACHE_TTL = 60.0


def _cache_key(query: str, search_content: bool) -> str:
    return f"{query.lower()}|||{search_content}"


def _get_cached(key: str) -> list[dict] | None:
    entry = _search_cache.get(key)
    if entry is None:
        return None
    ts, results = entry
    if time.monotonic() - ts > _SEARCH_CACHE_TTL:
        del _search_cache[key]
        return None
    return results


def _set_cache(key: str, results: list[dict]):
    _search_cache[key] = (time.monotonic(), results)
    # Evict oldest if cache grows too large
    if len(_search_cache) > 50:
        oldest = min(_search_cache.items(), key=lambda x: x[1][0])
        del _search_cache[oldest[0]]


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
# CHM status helpers
# ---------------------------------------------------------------------------

def _get_chm_status() -> list[dict]:
    """Return status of every known CHM: whether it has been extracted to HTML."""
    project_chm = os.path.join(_skill_dir, "chm")
    results: list[dict] = []
    for key, info in CHM_DOCS.items():
        chm_file = info["file"]
        desc = info["desc"]
        # Determine the expected _html directory name
        base = os.path.splitext(os.path.basename(chm_file))[0]
        html_dir_name = base + "_html"
        # Check if an _html folder exists (in chm/ or in install dir)
        converted = False
        html_dir_path = ""
        # Check project chm/ first
        candidate = os.path.join(project_chm, html_dir_name)
        if os.path.isdir(candidate):
            converted = True
            html_dir_path = candidate
        else:
            # Check install dir
            acam_dir = detect_alphacam_dir()
            if acam_dir:
                # Also check parent dir of chm_file
                chm_parent = os.path.dirname(os.path.join(acam_dir, chm_file))
                if os.path.isdir(chm_parent):
                    for entry in os.listdir(chm_parent):
                        if entry.lower() == html_dir_name.lower():
                            full = os.path.join(chm_parent, entry)
                            if os.path.isdir(full):
                                converted = True
                                html_dir_path = full
                                break
                # Also check acam_dir root
                root_candidate = os.path.join(acam_dir, html_dir_name)
                if os.path.isdir(root_candidate):
                    converted = True
                    html_dir_path = root_candidate
        results.append({
            "key": key,
            "file": chm_file,
            "desc": desc,
            "converted": converted,
            "html_dir": html_dir_path if converted else "",
        })
    return results


# ---------------------------------------------------------------------------
# Categories / listing
# ---------------------------------------------------------------------------

def _get_doc_categories(expand: bool = False) -> dict[str, Any]:
    """Return doc source directories with file counts (and optionally file lists)."""
    result: dict[str, Any] = {}
    for root in _get_doc_roots():
        files_list = []
        count = 0
        for r, _dirs, files in os.walk(root):
            for f in files:
                if _is_doc_file(f):
                    count += 1
                    if expand:
                        rel_path = os.path.relpath(os.path.join(r, f), root)
                        files_list.append({"file": f, "path": rel_path})
        if count > 0:
            label = os.path.basename(root)
            entry: dict = {"file_count": count}
            if expand:
                entry["files"] = files_list
            result[label] = entry
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


def _extract_snippet(text: str, query: str, context_chars: int = 80) -> str:
    """Extract a text snippet around the first occurrence of *query* (case-insensitive)."""
    q = query.lower()
    t = text.lower()
    idx = t.find(q)
    if idx == -1:
        return text[:context_chars * 2] + ("…" if len(text) > context_chars * 2 else "")
    start = max(0, idx - context_chars)
    end = min(len(text), idx + len(q) + context_chars)
    snippet = text[start:end]
    if start > 0:
        snippet = "…" + snippet
    if end < len(text):
        snippet = snippet + "…"
    return snippet


def _search_docs(query: str, max_results: int = 20,
                 search_content: bool = False) -> list[dict]:
    """Search doc pages by filename, title, and optionally content.

    Also searches CHM_DOCS descriptions for matching keywords.
    """
    q = query.lower()

    # Check cache first
    ck = _cache_key(q, search_content)
    cached = _get_cached(ck)
    if cached is not None:
        return cached

    results: list[dict] = []

    # --- Search CHM_DOCS descriptions ---
    for key, info in CHM_DOCS.items():
        desc = info["desc"]
        score = 0
        if q in key.lower():
            score += 2
        if q in desc.lower():
            score += 2
        if score > 0:
            results.append({
                "file": info["file"],
                "title": f"[CHM] {key} — {desc}",
                "path": info["file"],
                "source": "CHM_DOCS",
                "score": score + 10,  # Boost CHM hits so they appear early
                "snippet": f"📦 .chm 文档（未转换）: {desc}",
                "chm_key": key,
            })

    # --- Search extracted HTML docs ---
    for root in _get_doc_roots():
        for r, _dirs, files in os.walk(root):
            for f in files:
                if not _is_doc_file(f):
                    continue
                filepath = os.path.join(r, f)
                rel_path = os.path.relpath(filepath, root)
                score = 0
                snippet = ""
                if q in f.lower():
                    score += 2
                title = _get_doc_title(filepath)
                if q in title.lower():
                    score += 1
                if score > 0 or search_content:
                    content = ""
                    if score == 0 or search_content:
                        try:
                            with open(filepath, "r", encoding="utf-8",
                                      errors="replace") as fh:
                                raw = fh.read(16384)
                            content = _strip_html(raw)
                            if q in content.lower():
                                score += 1
                        except Exception:
                            pass
                    # Extract snippet if content was read
                    if content:
                        snippet = _extract_snippet(content, q)
                    elif title:
                        snippet = title
                if score > 0:
                    results.append({
                        "file": f,
                        "title": title,
                        "path": rel_path,
                        "source": os.path.basename(root),
                        "score": score,
                        "snippet": snippet,
                    })

    results.sort(key=lambda x: -x["score"])
    final = results[:max_results]

    # Cache the result
    _set_cache(ck, final)
    return final


# ---------------------------------------------------------------------------
# Handlers (synchronous, called from dispatcher)
# ---------------------------------------------------------------------------

def handle_list_docs(expand: bool = False) -> dict:
    """Handle the list_docs tool — now includes CHM_DOCS status."""
    roots = _get_doc_roots()
    categories = _get_doc_categories(expand=expand)
    total = 0
    for c in categories.values():
        total += c["file_count"] if isinstance(c, dict) else c
    acam_dir = detect_alphacam_dir()

    # Build CHM status block
    chm_status = _get_chm_status()
    converted_count = sum(1 for c in chm_status if c["converted"])
    pending_count = len(chm_status) - converted_count

    return {
        "alphacam_install_dir": acam_dir or "(not detected)",
        "doc_search_roots": roots,
        "doc_categories": categories,
        "total_html_files": total,
        "chm_docs_index": {
            "total": len(chm_status),
            "converted": converted_count,
            "pending": pending_count,
            "entries": chm_status,
        },
        "tip": "read_doc(name) / search_docs(query) / list_docs(expand=True) / chm_to_html_all()",
    }


def handle_read_doc(name: str, max_len: int = 8000) -> dict:
    """Handle the read_doc tool. max_len=0 means no truncation."""
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
    full_length = len(text)
    if max_len > 0 and len(text) > max_len:
        text = text[:max_len] + f"\n...[truncated {full_length} chars, set max_len=0 for full]"
    return {
        "title": title,
        "file": os.path.basename(filepath),
        "path": rel_path,
        "content": text,
        "full_length": full_length,
        "truncated": max_len > 0 and full_length > max_len,
    }


def handle_search_docs(query: str, search_content: bool = False) -> dict:
    """Handle the search_docs tool — now returns snippets and CHM hits."""
    results = _search_docs(query, search_content=search_content)
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
            {
                "file": r["file"],
                "title": r["title"],
                "path": r["path"],
                "source": r.get("source", ""),
                "snippet": r.get("snippet", ""),
                "chm_key": r.get("chm_key"),
            }
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

    # Invalidate search cache so newly extracted pages appear
    _search_cache.clear()

    return {
        "output_dir": output_dir,
        "chm_file": os.path.abspath(chm_path),
        "file_count": len(extracted),
        "total_size_bytes": total_size,
        "files": extracted,
    }


async def handle_convert_all_chm(output_base_dir: str | None = None) -> dict:
    """Convert all pending (not-yet-extracted) .chm files to HTML.

    Extracted output goes to ``chm/{Key}_html/`` inside the project folder,
    which is automatically picked up by ``_get_doc_roots()``.
    """
    project_chm = os.path.join(_skill_dir, "chm")
    if not os.path.isdir(project_chm):
        raise FileNotFoundError(
            f"Project chm/ directory not found at {project_chm}. "
            "Run the doc index setup first."
        )

    statuses = _get_chm_status()
    pending = [s for s in statuses if not s["converted"]]

    if not pending:
        return {
            "total_pending": 0,
            "converted": 0,
            "failed": 0,
            "results": [],
            "message": "All CHM files are already converted. Nothing to do.",
        }

    results = []
    converted_count = 0
    failed_count = 0

    for entry in pending:
        key = entry["key"]
        chm_rel = entry["file"]

        # Resolve the absolute path of the .chm file
        acam_dir = detect_alphacam_dir()
        if acam_dir:
            chm_abs = os.path.join(acam_dir, chm_rel)
        else:
            # Try project chm/
            chm_abs = os.path.join(project_chm, os.path.basename(chm_rel))

        if not os.path.isfile(chm_abs):
            log.warning(f"CHM file not found: {chm_abs}, skipping")
            results.append({"key": key, "file": chm_rel, "status": "skipped",
                            "error": "file not found"})
            failed_count += 1
            continue

        # Output goes to chm/{Key}_html/
        out_dir = os.path.join(project_chm, f"{key}_html")
        log.info(f"Converting {key} ({chm_abs}) → {out_dir}")

        try:
            conv_result = await handle_convert_chm_to_html(
                chm_path=chm_abs, output_dir=out_dir)
            results.append({"key": key, "file": chm_rel, "status": "ok",
                            "file_count": conv_result["file_count"]})
            converted_count += 1
        except Exception as e:
            log.exception(f"Failed to convert {key}")
            results.append({"key": key, "file": chm_rel, "status": "failed",
                            "error": str(e)})
            failed_count += 1

    return {
        "total_pending": len(pending),
        "converted": converted_count,
        "failed": failed_count,
        "results": results,
    }
