# -*- coding: utf-8 -*-
"""Generic deploy of one repo file into the RUNNING CDM project component.

    python tools/component_deploy.py audit  <Component> <repoPath> <baselineGlob>
    python tools/component_deploy.py deploy <Component> <repoPath> <baselineGlob>

baselineGlob picks the pre-change snapshot (newest match) used as the "clean
baseline" the running code must equal before we are willing to overwrite it.
Comparisons are case-insensitive: VBA normalises member-name case (.Add -> .add).
"""
import glob
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # 仓库根（tools/ 的上一级）
TMP = os.path.join(ROOT, "tmp")


def norm(s):
    return s.replace("\r\n", "\n").replace("\r", "\n").strip("\n") + "\n"


def main():
    a = sys.argv
    if len(a) < 5:
        raise SystemExit(__doc__)
    mode, comp_name, repo_rel, base_glob = a[1], a[2], a[3], a[4]
    repo_path = os.path.join(ROOT, repo_rel)
    cands = sorted(glob.glob(os.path.join(ROOT, base_glob)), key=os.path.getmtime)
    if not cands:
        raise SystemExit("FAIL: no baseline matched %s" % base_glob)
    base_path = cands[-1]
    print("baseline: %s" % os.path.relpath(base_path, ROOT))

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

    comp = None
    comps = proj.VBComponents
    for j in range(1, comps.Count + 1):
        if comps(j).Name == comp_name:
            comp = comps(j)
            break
    if comp is None:
        raise SystemExit("FAIL: component %s not found" % comp_name)

    cm = comp.CodeModule
    run = cm.Lines(1, cm.CountOfLines)
    repo = open(repo_path, "rb").read().decode("gbk")
    base = open(base_path, "rb").read().decode("gbk")

    in_sync = norm(run).lower() == norm(repo).lower()
    clean = norm(run).lower() == norm(base).lower()
    print("%-20s running_lines=%-6d ==repo=%-6s ==baseline=%-6s"
          % (comp_name, cm.CountOfLines, in_sync, clean))

    if mode != "deploy":
        return
    if in_sync:
        print("SKIP (already up to date)")
        return
    if not clean:
        raise SystemExit("ABORT: running %s matches neither repo nor baseline "
                         "(edited in the VBA editor?)" % comp_name)

    before = cm.CountOfLines
    cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(repo)
    ok = norm(cm.Lines(1, cm.CountOfLines)).lower() == norm(repo).lower()
    print("DEPLOY %-20s %d -> %d lines verify=%s" % (comp_name, before, cm.CountOfLines, ok))
    if not ok:
        cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString(run)
        raise SystemExit("ROLLBACK applied (verify failed)")
    os.makedirs(TMP, exist_ok=True)
    print("DONE")


if __name__ == "__main__":
    main()
