# -*- coding: utf-8 -*-
"""Post-deploy verification for the RUNNING CCC功能 project (modRamp v2.0).

Mirrors tools/deploy_probe.py (which is CDM-specific) for the CCC add-in project.

1. re-reads the running code and asserts the v2.0 markers are present and the
   superseded v1.x logic is gone
2. forces a project-wide VBA compile by running a pure, side-effect-free function
   (modRamp.RampVersion -- returns a constant, touches nothing)

VBA normalises member-name case (.Count -> .count) and reformats long numeric
literals (0.0174532925199433 -> 1.74532925199433E-02), so all text comparisons
are case-insensitive and we avoid long decimal literals in the source entirely.

ASCII-only stdout on purpose (console codepage is cp936).
"""
import sys

PROJECT = "CCC功能"

MARKERS = {
    "modRamp": [
        ("v2.1 header",        "v2.1.0 \u53d8\u66f4"),
        ("core: tail anchor",  "\u951a\u5b9a\u5728\u8f6e\u5ed3\u672b\u7aef"),
        ("core: prohibition",  "[\u7981\u6b62] \u628a\u659c\u5761\u6539\u6210"),
        ("closed-path gate",   "blnClosed = tp.Closed"),
        ("HBT score",          "score(k) = nl + ne"),
        ("exposed sides",      "Function ExposedSides"),
        ("all-parts mode",     "If minSize <= 0 Then"),
        ("start-point fix",    "\u671d\u5411\u6392\u7248\u4e2d\u5fc3\u90a3\u4e00\u4fa7\u7684\u3010\u8f83\u957f\u8fb9\u3011\u7684\u4e2d\u70b9"),
        ("element length",     "Function ElemLen"),
        ("pi as function",     "Pi = 4 * Atn(1)"),
        ("slow small parts",   "slowApplied = slowApplied + 1"),
        ("undo point",         "App.SetUndoPoint"),
        ("version function",   "Function RampVersion"),
        # --- v2.1.0 新增 ---
        ("v2.1 finish paths",  "Set newPaths = mtp.Finish"),
        ("v2.1 OpNo backfill", "newPaths(q).OpNo = CInt(origOpNo)"),
        ("v2.1 marker on new", "newPaths(q).Attribute(ATT_RAMP_DONE) = 1"),
        ("v2.1 OrderAll",      "drw.Operations.OrderAll"),
        ("v2.1 collect OpNo",  "colOpNo.Add CLng(tp.OpNo)"),
        # --- v2.1.1 新增（第 2 轮实机反馈）---
        ("v2.1.1 header",      "v2.1.1 \u53d8\u66f4"),
        ("v2.1.1 outside-in",  "\u8fdc\u7684\u5148\u5207 = \u4ece\u5916\u5f80\u5185"),
        ("v2.1.1 sheet center", "Function FindSheetCenter"),
        ("v2.1.1 retract Z20", "mtp.Add3DRapid startX, startY, SAFE_Z_UP"),
        ("v2.1.1 safe level",  "mdNew.SafeRapidLevel = SAFE_Z_UP"),
        ("v2.1.1 pbar off",    "Frame.ProjectBarUpdating = False"),
        ("v2.1.1 pbar on",     "Frame.ProjectBarUpdating = True"),
        # --- v2.1.2 新增（第 3 轮实机反馈: 参考中心算错）---
        ("v2.1.2 header",      "v2.1.2 \u53d8\u66f4"),
        ("v2.1.2 ref fallback", "ux1 = CDbl(colBX1(1))"),
        # --- v2.2.0 新增（第 4 轮实机反馈: 未匹配小板件的刀路要排到小板件之后）---
        ("v2.2 header",        "v2.2.0 \u53d8\u66f4"),
        ("v2.2 op name attr",  "LicomUKDMBOperationName"),
        ("v2.2 method name fn", "Function MethodNameOf"),
        ("v2.2 reorder call",  "If ReorderSmallFirst(drw) Then"),
        ("v2.2 reorder fn",    "Function ReorderSmallFirst"),
        ("v2.2 ordermanual",   "drw.OrderManual dst"),
        # v2.2.0: 不看 GetExtent 了, 参考中心改两级兜底
        ("v2.2 batch center",  "GetRefCenter drw, Nothing, colBX1, colBY1, colBX2, colBY2, refX, refY"),
        ("v2.2 start-side center", "SetGeoStartToSheetSide drw, ni, tp, toolGeo, refX, refY"),
    ],
    "frmRamp": [
        ("v2.1 control list",  "\u63a7\u4ef6\u6e05\u5355\uff0811 \u4e2a"),
        ("v2.1 7-arg call",    "slowSmall"),
        ("tool not preselected", "cmbMethodTool.ListIndex = -1"),
        ("tool not saved",     "\u4e0d\u518d\u4fdd\u5b58\u5200\u5177"),
        ("minSize 0 allowed",  "\u5c0f\u6761\u8303\u56f4\u4e0d\u80fd\u4e3a\u8d1f\u6570"),
    ],
}
ABSENT = {
    "modRamp": [
        "bestDist",                                  # v1.x start-point search
        "\u671d\u6392\u7248\u4e2d\u5fc3\u65b9\u5411\u504f\u79fb\u6574\u6761\u8fb9\u957f",  # v1.x hop-by-edge
        "DEG2RAD          As Double = 0",            # v1.x long decimal literal
        # v2.1.0 移除的微连接/留皮
        "Function BuildTabWindows",
        "Sub AddContourWithTabs",
        "Function InTabWindow",
        "doTabsUse",
        "tabApplied",
        "tabZ = -",
        "drw.GetExtent gx1",                         # v2.1.2/v2.2.0: 该 API 不可信, 已不用
        'SetAttribute "LicomUKDMBOperationName"',     # v2.2.1: VBA 里必须用带参属性写法
    ],
    "frmRamp": [
        "\u5c0f\u6761\u8303\u56f4\u5fc5\u987b\u5927\u4e8e 0",   # 旧校验(minSize<=0 就报错)
        "chkTabs.Value",                             # v2.1.0 已删除的控件(代码里不应再引用)
        "txtTabStock.",
        "g_lastDoTabs",
        "g_lastTabStock",
        "g_lastMethodTool",                          # 刀具不再记忆
    ],
}
WATCH = ("modRamp", "frmRamp")


def main():
    import win32com.client as w

    app = w.GetActiveObject("aroutaps.Application")
    vbe = app.VBE
    proj = None
    for i in range(1, vbe.VBProjects.Count + 1):
        if vbe.VBProjects(i).Name == PROJECT:
            proj = vbe.VBProjects(i)
            break
    if proj is None:
        print("FAIL: project %s not found (protected?)" % PROJECT)
        return 1
    try:
        print("project %s: %d components" % (PROJECT, proj.VBComponents.Count))
    except Exception as e:
        print("FAIL: cannot read components (protected?): %s" % e)
        return 1

    code = {}
    for j in range(1, proj.VBComponents.Count + 1):
        c = proj.VBComponents(j)
        if c.Name in WATCH:
            code[c.Name] = c.CodeModule.Lines(1, c.CodeModule.CountOfLines)

    ok = True
    for cname in sorted(MARKERS):
        src = code.get(cname, "")
        low = src.lower()
        print("-- %s (%d lines)" % (cname, src.count("\n") + (1 if src else 0)))
        crlf = src.count("\r\n")
        bare = src.count("\n") - crlf
        print("   EOL    CRLF=%-5d bare_LF=%-5d %s"
              % (crlf, bare, "OK" if crlf > 0 and bare == 0 else "*** MIXED/LF ***"))
        for label, needle in MARKERS[cname]:
            hit = needle.lower() in low
            ok &= hit
            print("   MARKER %-22s %s" % (label, "OK" if hit else "*** MISSING ***"))
        for needle in ABSENT.get(cname, []):
            gone = needle.lower() not in low
            ok &= gone
            print("   ABSENT %-22s %s" % (needle[:22],
                                          "OK" if gone else "*** STILL PRESENT ***"))

    if "modRamp" in code:
        print("-- compile probe: %s.modRamp.RampVersion() --" % PROJECT)
        try:
            r = app.Run("%s.modRamp.RampVersion" % PROJECT)
            print("   compile+run OK, returned %r" % (r,))
            ok &= isinstance(r, str) and r.startswith("modRamp")
        except Exception as e:
            ok = False
            print("   *** PROBE FAILED (project does not compile): %s" % e)

    print("RESULT=%s" % ("PASS" if ok else "FAIL"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
