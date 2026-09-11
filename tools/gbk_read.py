# -*- coding: utf-8 -*-
"""Print a line range of a GBK-encoded source file.

    python tools/gbk_read.py <file> <start> [end]
"""
import sys

path = sys.argv[1]
start = int(sys.argv[2])
end = int(sys.argv[3]) if len(sys.argv) > 3 else start + 60

raw = open(path, "rb").read().decode("gbk", errors="replace")
lines = raw.splitlines()
print("# %s : lines %d-%d of %d" % (path, start, end, len(lines)))
for i in range(start - 1, min(end, len(lines))):
    print("%5d| %s" % (i + 1, lines[i].rstrip("\r")))
