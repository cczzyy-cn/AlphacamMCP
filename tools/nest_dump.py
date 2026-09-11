# -*- coding: utf-8 -*-
"""READ-ONLY: dump the active drawing's nest structure.

Iterates with For Each (the production code does the same; indexed access can
raise error 5 on these collections).  The error handler APPENDS so partial
output survives.

    python tools/nest_dump.py
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # 仓库根（tools/ 的上一级）
OUT = os.path.join(ROOT, "tmp", "nest_dump.txt")
CNT = "LicomUKljo_alphadoor_nest_door_count"
DET = "LicomUSrlg_alphadoor_order_detail_id"


def vba():
    out = OUT.replace("\\", "\\\\")
    return "\n".join([
        "   Dim iFile As Integer, Ni As Object, Nsh As Object, Npi As Object, P As Object",
        "   Dim nS As Long, nP As Long, nPath As Long",
        "   Dim sCnt As String, sDet As String, sAll As String",
        "   iFile = FreeFile",
        '   Open "%s" For Output As #iFile' % out,
        "   On Error GoTo EH",
        '   Print #iFile, "ActiveDrawing=[" & ActiveDrawing.Name & "]"',
        '   Print #iFile, "FullName=[" & ActiveDrawing.FullName & "]"',
        '   Print #iFile, "Geometries=" & ActiveDrawing.Geometries.Count',
        "   Set Ni = ActiveDrawing.GetNestInformation",
        "   If Ni Is Nothing Then",
        '      Print #iFile, "NO NEST INFORMATION (not a nest drawing)"',
        "      Close #iFile",
        "      Exit Sub",
        "   End If",
        '   Print #iFile, "Sheets=" & Ni.Sheets.Count',
        "   For Each Nsh In Ni.Sheets",
        "      nS = nS + 1",
        '      Print #iFile, "== Sheet #" & nS & " name=[" & Nsh.Name & "] parts=" & Nsh.Parts.Count',
        "      For Each Npi In Nsh.Parts",
        "         nP = nP + 1",
        "         sAll = \"\"",
        "         nPath = 0",
        "         On Error Resume Next",
        "         nPath = Npi.Paths.Count",
        "         On Error GoTo EH",
        "         For Each P In Npi.Paths",
        '            sCnt = "" & P.Attribute("%s")' % CNT,
        '            sDet = "" & P.Attribute("%s")' % DET,
        '            sAll = sAll & " [cnt=" & sCnt & ",det=" & sDet & "]"',
        "         Next P",
        '         Print #iFile, "   part#" & nP & " name=[" & Npi.Name & "] paths=" & nPath & sAll',
        "      Next Npi",
        "   Next Nsh",
        "   Close #iFile",
        "   Exit Sub",
        "EH:",
        "   Dim en As Long, ed As String",
        "   en = Err.Number: ed = Err.Description",
        "   On Error Resume Next",
        "   Close #iFile",
        '   Open "%s" For Append As #iFile' % out,
        '   Print #iFile, "ERR " & en & " | " & ed & " (at sheet#" & nS & " part#" & nP & ")"',
        "   Close #iFile",
    ])


def main():
    if os.path.exists(OUT):
        os.remove(OUT)
    sys.path.insert(0, ROOT)
    from alphacam_com import AlphaCAM

    res = AlphaCAM().run_vba_line(vba())
    print("run_vba_line: %s" % res.get("status"))
    if not os.path.exists(OUT):
        print("no output file")
        return 1
    print("\n==== nest dump ====")
    for line in open(OUT, "rb").read().decode("gbk", errors="replace").splitlines():
        print("  " + line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
