# -*- coding: utf-8 -*-
"""Post-deploy verification for the running CDM project (piece-uid release, v1.9;
frmAutoNest material combo, v1.10).

1. re-reads the running code and asserts the v1.9 markers are present / the
   superseded matching logic is gone
2. forces a project-wide VBA compile by running a pure, side-effect-free function
   (Functions.gs_FixSQL -- string replace only)
3. reports the outcome; any modal dialog is left for the caller to inspect

VBA normalises member-name case (.Add -> .add) and re-writes line endings to
CRLF, so all text comparisons are case-insensitive.

ASCII-only stdout on purpose (console codepage is cp936).
"""
import os
import sys

UID_ATTR = "LicomUSrlg_alphadoor_piece_uid"

MARKERS = {
    "Make": [
        ("v2.2 header",         "版本: v2.2 (2026-09-11)"),
        ("v2.2 uid note",       "m_CreateAlphaCAMDrawingsOfSheets 给每个 part instance 的全部路径写"),
        ("v1.8 ScreenUpd restore", "    ActiveDrawing.ScreenUpdating = True"),
        ("v1.8 ProjBar restore",   "    Frame.ProjectBarUpdating = True"),
        ("v1.8 comment",           "v1.8 \u4fee\u590d"),
        ("label fix: dedup snapshot",  "colGeo.Add SheetPath"),
        ("label fix: collect hatches", "colHatch.Add ps"),
        ("label fix: delete all",      "colHatch.Item(i).Delete"),
        ("v1.9 uid const",      'Private Const DEF_ATT_PIECE_UID As String = "%s"' % UID_ATTR),
        ("v1.9 uid stamp",      "NestPath.Attribute(DEF_ATT_PIECE_UID) = strUID"),
        ("v1.9 uid reuse",      'If strUID = "" Then strUID = "" & NestPath.Attribute(DEF_ATT_PIECE_UID)'),
        ("v1.9 uid alloc",      '"U" & Format$(lngNextUID, "0000")'),
        ("v1.9 uid cleared",    'NestPath.Attribute(DEF_ATT_PIECE_UID) = strUID'),
        ("B fix lngPK reset",   "lngPK = 0"),
    ],
    "modAutoImportNest": [
        ("menu entry = CDM",   "用法: AlphaCAM 菜单 → CDM → 自动化生产排版"),
        ("menu chain noted",   "CDM 菜单 → Events.m_AutoImportNest(Events.bas:2866) → AutoImportNest()"),
        ("entry comment",      "主入口（CDM 菜单 → 自动化生产排版"),
        ("v1.10 header",       "v1.10 (2026-09-13)"),
        ("v1.10 override param",  "Optional ByVal sMaterialOverride As String = \"\""),
        ("v1.10 override check",  "校验窗体材料"),
        ("v1.10 override assign", 'If sMaterialOverride <> "" Then sMat = sMaterialOverride'),
        ("v1.10 override log",    "窗体材料="),
        ("v1.10 passthrough",     "bOverwrite, sMaterialOverride)"),
        ("v1.9 header",        "v1.9 (2026-09-11)"),
        ("v1.9 uid const",     "Private Const DEF_ATT_PIECE_UID"),
        ("v1.9 prefetch",      '"SELECT PK, DetailID, SheetName, PressPieceUID FROM AD_REPORT_DATA"'),
        ("v1.9 uid index",     'dicUID.Exists(sKey & "|" & strUID)'),
        ("v1.9 pair queue",    "dicQ(sKey)"),
        ("v1.9 claimed",       "dicClm.Add CStr(lngPKRow), 1"),
        ("v1.9 update by pk",  '" WHERE PK=" & lngPKRow'),
        ("v1.9 backfill uid",  '"R" & Format$(lngPKRow, "0000")'),
        ("v1.9 dedup delete",  '" AND SheetName IN (" & sSheetList & ") AND PK NOT IN ("'),
        ("v1.9 col gate",      "blnUIDCol = False"),
        ("v1.9 uid sep write", 'gdb_CDM.Execute "UPDATE AD_REPORT_DATA SET PressPieceUID='),
        ("v1.9 graceful",      "If blnUIDCol And strUID <> \"\" Then"),
        ("v1.9 col self-heal", '"SELECT TOP 1 PressPieceUID FROM AD_REPORT_DATA"'),
        ("v1.8 header",      "\u7248\u672c: v1.8 (2026-09-10)"),
        ("B2 Kill+RmDir",    "Kill sBakDir & \"*.*\""),
        ("B3 cleanup log",   "\u4e34\u65f6\u5d4c\u5957\u6863\u6848\u5df2\u6e05\u7406"),
        ("B1 tail restore",  "Frame.ProjectBarUpdating = True"),
        ("v1.7 header",      "\u7248\u672c: v1.7 (2026-09-10)"),
        ("ImportCSV txn on", "gdb_CDM.BeginTrans"),
        ("ImportCSV commit", "gdb_CDM.CommitTrans"),
        ("regen EH restore", "blnScratchOpened"),
        ("wait +sheets",     "lngExpectedDoors + lngSheetCount"),
        ("no material arg",  "ByVal bRunNest As Boolean, _"),
    ],
    "Events": [
        ("v1.9 field block", 'mbln_DBFieldExists(r, "PressPieceUID")'),
        ("v1.9 alter",       "ALTER TABLE AD_REPORT_DATA ADD PressPieceUID VARCHAR(64)"),
    ],
    "frmAutoNest": [
        ("4-arg call",       "Not chkOnlyImport.Value, chkOverwrite.Value"),
        ("no empty material", "PtrSafe"),
        ("v1.10 combo load",  "m_LoadMaterials"),
        ("v1.10 combo query", "SELECT Name, MaterialDefault FROM AD_MATERIALS ORDER BY Name"),
        ("v1.10 preselect",   "m_SelectMaterial sDef"),
        ("v1.10 list guard",  "If cboMaterial.ListIndex < 0 Then"),
        ("v1.10 5-arg call",  "chkOverwrite.Value, "),
        ("v1.10 pass material", "cboMaterial.Value"),
        ("v1.10 last material", 'SaveSetting "CCC", "AutoImportNest", "LastMaterial"'),
    ],
}
ABSENT = {
    "Make": [
        "        ps.Delete",
        "Set SheetPath2 = Nothing",
        # Make.bas must NOT reference the optional DB column at all: the DDL can
        # only succeed at CDM start-up, so a missing column must never break the
        # legacy report path.
        "PressPieceUID",
    ],
    "modAutoImportNest": [
        "sDefaultMaterial As String",
        "Sub glng_EnsureMaterial",
        "sMaterialName As String",
        "    'ActiveDrawing.ScreenUpdating = True\n    'Frame.ProjectBarUpdating = True",
        " WHERE DetailID=\" & lngDetail",
        # the superseded 3-key match must be gone; the new UPDATE writes the
        # counter back, so match the full old WHERE fragment, not just the name
        " AND PressDoorCounter=\" & lngCnt & \" AND SheetName=",
        # the auto-nest entry lives in the CDM menu, never in the CCC功能 menu
        "CCC功能 → 自动化生产排版",
        "CCC 功能菜单触发",
    ],
    "Events": [],
    "frmAutoNest": [
        "AutoImportNestWithParams _\n        Trim$(txtCSV), \"\u81ea\u52a8\u5316\u751f\u4ea7\", \"\", ",
    ],
}
WATCH = ("Make", "modAutoImportNest", "frmAutoNest", "Events")


def main():
    import win32com.client as w

    app = w.GetActiveObject("aroutaps.Application")
    vbe = app.VBE
    proj = None
    for i in range(1, vbe.VBProjects.Count + 1):
        p = vbe.VBProjects(i)
        if p.Name == "CDM":
            proj = p
            break
    if proj is None:
        raise SystemExit("FAIL: CDM project not found")

    comps = proj.VBComponents
    code, names = {}, []
    for j in range(1, comps.Count + 1):
        c = comps(j)
        names.append(c.Name)
        if c.Name in WATCH:
            code[c.Name] = c.CodeModule.Lines(1, c.CodeModule.CountOfLines)
    print("CDM components (%d): %s" % (len(names), sorted(names)))

    ok = True
    for cname in sorted(MARKERS):
        src = code.get(cname, "")
        low = src.lower()
        print("-- %s (%d lines)" % (cname, src.count("\n") + (1 if src else 0)))
        crlf = src.count("\r\n")
        bare = src.count("\n") - crlf
        eol_ok = (crlf > 0 and bare == 0)
        ok &= eol_ok
        print("   EOL    CRLF=%-5d bare_LF=%-5d %s"
              % (crlf, bare, "OK" if eol_ok else "*** MIXED/LF ***"))
        # VBA normalises member-name case (e.g. .Add -> .add), so compare case-insensitively
        for label, needle in MARKERS[cname]:
            hit = needle.lower() in low
            ok &= hit
            print("   MARKER %-24s %s" % (label, "OK" if hit else "*** MISSING ***"))
        for needle in ABSENT.get(cname, []):
            gone = needle.lower() not in low
            ok &= gone
            print("   ABSENT %-24s %s" % (needle[:24].replace("\n", " "),
                                          "OK" if gone else "*** STILL PRESENT ***"))

    print("-- compile probe: CDM.Functions.gs_FixSQL(\"a'b\") --")
    try:
        r = app.Run("CDM.Functions.gs_FixSQL", "a'b")
        good = (r == "a''b")
        ok &= good
        print("   compile+run %s, returned %r (expect %r)"
              % ("OK" if good else "UNEXPECTED", r, "a''b"))
    except Exception as e:
        ok = False
        print("   *** PROBE FAILED: %s" % e)

    print("RESULT=%s" % ("PASS" if ok else "FAIL"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
