# -*- coding: utf-8 -*-
"""Query the CDM DB the project's own way: run DAO inside AlphaCAM's VBA
(via alphacam_com.run_vba_line) and read the text the macro writes out.

My python is 64-bit and Jet 4.0 is 32-bit only, so this is the only route.

    python tools/db_query.py "<JobName>"
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # 仓库根（tools/ 的上一级）
OUT = os.path.join(ROOT, "tmp", "db_report.txt")
DB = r"D:\2016\LICOMDAT\CDM Data\CDM.mdb"


def vba(job_name):
    esc = job_name.replace("'", "''")
    out = OUT.replace("\\", "\\\\")
    db = DB.replace("\\", "\\\\")
    return "\n".join([
        "   On Error GoTo EH",
        "   Dim dbe As Object, db As Object, rs As Object, sql As String",
        "   Dim iFile As Integer, n As Long, lngOrderID As Long",
        "   Dim sImg As String, sPrev As String",
        "   iFile = FreeFile",
        '   Open "%s" For Output As #iFile' % out,
        '   Set dbe = CreateObject("DAO.DBEngine.36")',
        "   Dim tries As Integer, opened As Boolean",
        "   opened = False",
        "   For tries = 1 To 12",
        "      On Error Resume Next",
        '      Set db = dbe.OpenDatabase("%s", False, True)' % db,
        "      If Err.Number = 0 Then opened = True",
        "      Err.Clear",
        "      On Error GoTo EH",
        "      If opened Then Exit For",
        "      Dim t0 As Single",
        "      t0 = Timer",
        "      Do While Timer - t0 < 1.5",
        "         DoEvents",
        "      Loop",
        "   Next tries",
        "   If Not opened Then",
        '      Print #iFile, "DB LOCKED after retries"',
        "      Close #iFile",
        "      Exit Sub",
        "   End If",
        '   Set rs = db.OpenRecordset("SELECT OrderID FROM AD_ORDERS WHERE JobName=\'%s\'")' % esc,
        "   If rs.EOF Then",
        '      Print #iFile, "ORDER NOT FOUND"',
        "   Else",
        '      lngOrderID = rs.Fields("OrderID")',
        '      Print #iFile, "OrderID=" & lngOrderID',
        "   End If",
        "   rs.Close",
        "   If lngOrderID > 0 Then",
        '      sql = "SELECT * FROM AD_REPORT_DATA WHERE OrderID=" & lngOrderID & _',
        '            " ORDER BY SheetName, PressDoorCounter, DetailID"',
        "      Set rs = db.OpenRecordset(sql)",
        "      n = 0",
        "      Dim f As Long",
        "      Do While Not rs.EOF",
        "         n = n + 1",
        '         Print #iFile, "--- row " & n & " ---"',
        "         For f = 0 To rs.Fields.Count - 1",
        '            Print #iFile, "   " & rs.Fields(f).Name & "=" & ("" & rs.Fields(f).Value)',
        "         Next f",
        "         rs.MoveNext",
        "      Loop",
        '      Print #iFile, "ROWCOUNT=" & n',
        "      rs.Close",
        "   End If",
        "   db.Close",
        "   Close #iFile",
        "   Exit Sub",
        "EH:",
        "   Dim en As Long, ed As String",
        "   en = Err.Number: ed = Err.Description",
        "   On Error Resume Next",
        "   Close #iFile",
        '   Open "%s" For Output As #iFile' % out,
        '   Print #iFile, "ERR " & en & " | " & ed',
        "   Close #iFile",
    ])


def main():
    job = sys.argv[1] if len(sys.argv) > 1 else "9-11纳百川"
    if os.path.exists(OUT):
        os.remove(OUT)

    sys.path.insert(0, ROOT)
    from alphacam_com import AlphaCAM

    acam = AlphaCAM()
    print("connected: %s" % acam.get_info().get("path", ""))
    res = acam.run_vba_line(vba(job))
    print("run_vba_line: %s" % res.get("status"))

    if not os.path.exists(OUT):
        print("no output file -> macro did not write")
        return 1
    text = open(OUT, "rb").read().decode("gbk", errors="replace")
    print("\n==== AD_REPORT_DATA for %s ====" % job)
    for line in text.splitlines():
        print("  " + line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
