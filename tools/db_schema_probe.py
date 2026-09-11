# -*- coding: utf-8 -*-
"""Ensure AD_REPORT_DATA.PressPieceUID exists -- using the PROJECT'S OWN
connection (gdb_CDM / gbln_ConnectToDB), i.e. exactly what the deployed
self-heal runs.  run_vba_line injects into a temp module inside CDM, so the
project globals are in scope.

Also dumps the job's rows with PK so the new matching can be replayed.

    python tools/db_schema_probe.py [JobName]
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # 仓库根（tools/ 的上一级）
OUT = os.path.join(ROOT, "tmp", "addcol.txt")
DB = r"D:\2016\LICOMDAT\CDM Data\CDM.mdb"


def vba(job):
    esc = job.replace("'", "''").replace('"', '""')
    out = OUT.replace("\\", "\\\\")
    return "\n".join([
        "   Dim iFile As Integer, rs As Object, n As Long, e1 As Long",
        "   Dim blnOk As Boolean, sConn As String",
        "   iFile = FreeFile",
        '   Open "%s" For Output As #iFile' % out,
        "   On Error GoTo EH",
        "   blnOk = gbln_ConnectToDB()",
        '   Print #iFile, "connected=" & blnOk',
        "   If Not blnOk Then",
        "      Close #iFile",
        "      Exit Sub",
        "   End If",
        "   ' --- 1) probe (same statement the deployed self-heal uses) ---",
        "   On Error Resume Next",
        '   Set rs = gdb_CDM.Execute("SELECT TOP 1 PressPieceUID FROM AD_REPORT_DATA")',
        "   e1 = Err.Number",
        "   Err.Clear",
        "   On Error GoTo EH",
        '   Print #iFile, "probe_err=" & e1',
        "   If e1 <> 0 Then",
        "      On Error Resume Next",
        '      gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD PressPieceUID VARCHAR(64)"',
        "      e1 = Err.Number",
        "      Err.Clear",
        "      On Error GoTo EH",
        '      Print #iFile, "alter_err=" & e1 & " (" & Err.Description & ")"',
        "      On Error Resume Next",
        '      Set rs = gdb_CDM.Execute("SELECT TOP 1 PressPieceUID FROM AD_REPORT_DATA")',
        "      e1 = Err.Number",
        "      Err.Clear",
        "      On Error GoTo EH",
        '      Print #iFile, "reprobe_err=" & e1',
        "   Else",
        "      rs.Close",
        '      Print #iFile, "column already present"',
        "   End If",
        "   ' --- 2) dump rows ---",
        '   Set rs = gdb_CDM.Execute("SELECT PK, DetailID, SheetName, PressDoorCounter, PressPieceUID, PressDoorImage FROM AD_REPORT_DATA WHERE OrderID=(SELECT OrderID FROM AD_ORDERS WHERE JobName=\'%s\') ORDER BY PK")' % esc,
        "   Do While Not rs.EOF",
        "      n = n + 1",
        '      Print #iFile, "PK=" & rs.Fields("PK") & " DetailID=" & rs.Fields("DetailID") & _',
        '            " cnt=" & ("" & rs.Fields("PressDoorCounter")) & " uid=[" & ("" & rs.Fields("PressPieceUID")) & "]"',
        '      Print #iFile, "    img=" & ("" & rs.Fields("PressDoorImage"))',
        "      rs.MoveNext",
        "   Loop",
        '   Print #iFile, "ROWCOUNT=" & n',
        "   rs.Close",
        "   Close #iFile",
        "   Exit Sub",
        "EH:",
        "   Dim en As Long, ed As String",
        "   en = Err.Number: ed = Err.Description",
        "   On Error Resume Next",
        "   Close #iFile",
        '   Open "%s" For Append As #iFile' % out,
        '   Print #iFile, "ERR " & en & " | " & ed',
        "   Close #iFile",
    ])


def main():
    job = sys.argv[1] if len(sys.argv) > 1 else "9-11纳百川"
    if os.path.exists(OUT):
        os.remove(OUT)
    sys.path.insert(0, ROOT)
    from alphacam_com import AlphaCAM

    res = AlphaCAM().run_vba_line(vba(job))
    print("run_vba_line: %s" % res.get("status"))
    if not os.path.exists(OUT):
        print("no output file")
        return 1
    text = open(OUT, "rb").read().decode("gbk", errors="replace")
    for line in text.splitlines():
        print("  " + line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
