# -*- coding: utf-8 -*-
"""Snapshot the RUNNING CDM components to backup/ before making further changes.

Usage: python tools/running_snapshot.py [tag]
Writes backup/alphacam_running_<Component>_<tag>_<stamp>.bas (GBK).
"""
import datetime
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # 仓库根（tools/ 的上一级）
TAG = sys.argv[1] if len(sys.argv) > 1 else "snapshot"
WANT = ("Make", "modAutoImportNest", "frmAutoNest", "Events")


def main():
    import win32com.client as w

    app = w.GetActiveObject("aroutaps.Application")
    proj = None
    for i in range(1, app.VBE.VBProjects.Count + 1):
        p = app.VBE.VBProjects(i)
        if p.Name == "CDM":
            proj = p
            break
    if proj is None:
        raise SystemExit("FAIL: CDM project not found")

    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = os.path.join(ROOT, "backup")
    os.makedirs(out_dir, exist_ok=True)

    for j in range(1, proj.VBComponents.Count + 1):
        c = proj.VBComponents(j)
        if c.Name not in WANT:
            continue
        cm = c.CodeModule
        txt = cm.Lines(1, cm.CountOfLines)
        path = os.path.join(out_dir, "alphacam_running_%s_%s_%s.bas" % (c.Name, TAG, stamp))
        with open(path, "wb") as f:
            f.write(txt.encode("gbk", errors="replace"))
        print("%-20s %5d lines -> %s" % (c.Name, cm.CountOfLines, os.path.basename(path)))
    print("DONE")


if __name__ == "__main__":
    main()
