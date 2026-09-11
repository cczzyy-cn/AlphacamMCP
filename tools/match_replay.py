# -*- coding: utf-8 -*-
"""Replay the NEW matching algorithm against the REAL data.

Inputs are the two live dumps:
  tmp/nest_dump.txt  -- every part instance of the active nest drawing, in
                        SH.Parts order, with its DetailID and NEST_DOOR_COUNT
  tmp/db_report.txt  -- AD_REPORT_DATA rows for the job (PK + DetailID +
                        PressDoorCounter + PressDoorImage)

The VBA in modAutoImportNest 5.3 is a direct transcription of this, so the
replay shows what the next "生成标签" will do to the current bad rows.

    python tools/match_replay.py
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # 仓库根（tools/ 的上一级）
NEST = os.path.join(ROOT, "tmp", "nest_dump.txt")
DB = os.path.join(ROOT, "tmp", "db_report.txt")


def gbk(path):
    return open(path, "rb").read().decode("gbk", errors="replace")


def parse_instances():
    """[(sheet, order_in_sheet, detail, cnt)] in drawing order."""
    out = []
    sheet = ""
    order = 0
    for line in gbk(NEST).splitlines():
        m = re.match(r"\s*== Sheet #\d+ name=\[(.*?)\]", line)
        if m:
            sheet = m.group(1)
            order = 0
            continue
        m = re.match(r"\s*part#\d+ name=\[(.*?)\] paths=(\d+)\s*(.*)$", line)
        if not m:
            continue
        attrs = m.group(3)
        d = re.search(r"\[cnt=(\d*),det=(\d*)\]", attrs)
        order += 1
        if d and d.group(2):
            out.append((sheet, order, int(d.group(2)), int(d.group(1))))
    return out


def parse_rows():
    """[(pk, detail, cnt, image)] ordered by PK (the SELECT ... ORDER BY PK)."""
    rows = []
    cur = {}
    for line in gbk(DB).splitlines():
        m = re.match(r"\s*PK=(\d+)", line)
        if m:
            if cur.get("pk"):
                rows.append(cur)
            cur = {"pk": int(m.group(1)), "detail": None, "cnt": None, "img": ""}
            continue
        m = re.match(r"\s*DetailID=(\d+)", line)
        if m and cur:
            cur["detail"] = int(m.group(1))
        m = re.match(r"\s*PressDoorCounter=(\d*)", line)
        if m and cur:
            cur["cnt"] = int(m.group(1) or 0)
        m = re.match(r"\s*PressDoorImage=(.*)", line)
        if m and cur:
            cur["img"] = m.group(1).strip()
    if cur.get("pk"):
        rows.append(cur)
    return sorted(rows, key=lambda r: r["pk"])


def main():
    inst = parse_instances()
    rows = parse_rows()
    if not inst or not rows:
        print("missing input dumps (run tools/nest_dump.py / tools/db_query.py)")
        return 1

    print("== drawing instances (SH.Parts order) ==")
    for s, o, d, c in inst:
        print("   sheet=%-10s order=%d DetailID=%d cnt=%d" % (s, o, d, c))
    print("== AD_REPORT_DATA rows (PK order) ==")
    for r in rows:
        print("   PK=%-8d DetailID=%-8s cnt=%-4s %s"
              % (r["pk"], r["detail"], r["cnt"], os.path.basename(r["img"])))

    # --- rebuild the two indices exactly like the VBA does ---
    queue, idx, claimed = {}, {}, {}
    for r in rows:
        k = "%s|%s" % (r["detail"], "Sheet A1")
        queue.setdefault(k, []).append(r["pk"])
        idx.setdefault(k, 1)

    print("\n== replaying the new 5.3 ==")
    new = []
    for sheet, order, detail, cnt in inst:
        k = "%s|%s" % (detail, sheet)
        pk = 0
        q = queue.get(k, [])
        i = idx.get(k, 1)
        while i <= len(q):
            if q[i - 1] not in claimed:
                pk = q[i - 1]
                idx[k] = i + 1
                break
            i += 1
        if pk == 0:
            print("   UNMATCHED detail=%d cnt=%d" % (detail, cnt))
            continue
        claimed[pk] = True
        img = "..._Sheet A1_%d.emf" % cnt
        old = next(r for r in rows if r["pk"] == pk)
        print("   detail=%-8d cnt=%-3d -> PK=%-8d  cnt %s->%d  img %s -> %s"
              % (detail, cnt, pk, old["cnt"], cnt,
                 os.path.basename(old["img"]), os.path.basename(img)))
        new.append((pk, detail, cnt, img))

    unclaimed = [r for r in rows if r["pk"] not in claimed]
    print("\n   rows claimed=%d  unclaimed(deleted)=%d" % (len(claimed), len(unclaimed)))
    for r in unclaimed:
        print("      DELETE PK=%d" % r["pk"])

    imgs = [n[3] for n in new]
    print("\nRESULT images=%s" % [os.path.basename(x) for x in imgs])
    print("RESULT unique=%s  duplicate_count=%d"
          % (len(set(imgs)) == len(imgs), len(imgs) - len(set(imgs))))
    return 0


if __name__ == "__main__":
    sys.exit(main())
