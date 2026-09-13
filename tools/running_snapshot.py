# -*- coding: utf-8 -*-
"""Snapshot the RUNNING project components to backup/ before making changes.

Usage: python tools/running_snapshot.py [tag] [Project] [comp1,comp2,...]
Writes backup/alphacam_running_<Component>_<tag>_<stamp>.bas (GBK).

Project defaults to CDM. Omit the component list to use the CDM default set
(Make/modAutoImportNest/frmAutoNest/Events); for any other project, omitting it
snapshots ALL components.
"""
import datetime
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # 仓库根（tools/ 的上一级）
TAG = sys.argv[1] if len(sys.argv) > 1 else "snapshot"
PROJ = sys.argv[2] if len(sys.argv) > 2 else "CDM"
if len(sys.argv) > 3:
    WANT = tuple(x.strip() for x in sys.argv[3].split(",") if x.strip())
elif PROJ == "CDM":
    WANT = ("Make", "modAutoImportNest", "frmAutoNest", "Events")
else:
    WANT = None      # None = 该工程的全部组件
CDM_DEFAULT = ("Make", "modAutoImportNest", "frmAutoNest", "Events")


def main():
    import win32com.client as w

    app = w.GetActiveObject("aroutaps.Application")
    proj = None
    for i in range(1, app.VBE.VBProjects.Count + 1):
        p = app.VBE.VBProjects(i)
        if p.Name == PROJ:
            proj = p
            break
    if proj is None:
        raise SystemExit("FAIL: project %s not found" % PROJ)
    print("project: %s  want: %s" % (PROJ, "ALL" if WANT is None else list(WANT)))

    # 受保护的工程读不到组件集合, 这里给出明确指引而不是裸 traceback
    # （AlphaCAM 自带插件工程默认是锁的, 例: CDM 在重开 AlphaCAM 后会恢复保护）
    try:
        total = proj.VBComponents.Count
    except Exception as e:
        raise SystemExit(
            "FAIL: 读不到工程 %s 的组件 (%s)\n"
            "      工程受保护时无法读写。请在 AlphaCAM VBA 编辑器中解除保护\n"
            "      （工具 → <工程>属性 → 保护）后重试。" % (PROJ, e))

    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = os.path.join(ROOT, "backup")
    os.makedirs(out_dir, exist_ok=True)

    for j in range(1, total + 1):
        c = proj.VBComponents(j)
        if WANT is not None and c.Name not in WANT:
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
