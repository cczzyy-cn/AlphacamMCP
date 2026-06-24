"""Test: mark selected + iterate erase."""
import sys, win32com.client

app = win32com.client.Dispatch("aroutaps.Application")
app.Visible = True
drw = app.ActiveDrawing

# Delete all
p = drw.GetFirstGeo()
i = 0
while p and i < 1000:
    try: p.Erase(); p = drw.GetFirstGeo(); i += 1
    except: break

# Create
rect = drw.CreateRectangle(0, 0, 300, 300)
l1 = drw.Create2DLine(-150, 150, 450, 150)
blen = rect.Length
print(f"Start: {drw.GetGeoCount()}")

drw.SetGeosSelected(False)
rect.Selected = True

# Find line
line = None
p = drw.GetFirstGeo()
while p:
    if abs(p.Length - blen) > 0.1:
        line = p
    try: p = p.GetNext()
    except: break

# Break
segs = line.BreakWithCuttingGeos()
print(f"Break: {segs.Count} segments, total: {drw.GetGeoCount()}")

# Mark external ones as selected
drw.SetGeosSelected(False)
for j in range(1, segs.Count + 1):
    seg = segs.Item(j)
    h = seg.Length / 2
    r = seg.PointAtDistanceAlongPathG(h, 0, 0, 0, None)
    if isinstance(r, tuple) and r[0]:
        inside = rect.IsPointInside(r[1], r[2])
        print(f"  seg{j}: len={seg.Length:.0f} at ({r[1]:.0f},{r[2]:.0f}) inside={inside}")
        if not inside:
            seg.Selected = True
            print(f"    -> MARKED for deletion")

# Now delete by iterating drawing
deleted = 0
p = drw.GetFirstGeo()
while p and deleted < 10:
    if p.Selected:
        print(f"  Erasing: len={p.Length:.0f}")
        p.Erase()
        deleted += 1
        p = drw.GetFirstGeo()
    else:
        try: p = p.GetNext()
        except: break

print(f"Deleted: {deleted}")
print(f"Final: {drw.GetGeoCount()}")
drw.ZoomAll()
