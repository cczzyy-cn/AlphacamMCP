"""Quick boundary trim test - no sleep."""
import sys, win32com.client

app = win32com.client.Dispatch("aroutaps.Application")
app.Visible = True
drw = app.ActiveDrawing

# Delete all existing (with loop guard)
max_iter = 1000
p = drw.GetFirstGeo()
i = 0
while p and i < max_iter:
    try:
        p.Erase()
        p = drw.GetFirstGeo()
        i += 1
    except:
        break

# Create scene
rect = drw.CreateRectangle(0, 0, 300, 300)
l1 = drw.Create2DLine(-150, 150, 450, 150)
l2 = drw.Create2DLine(150, -150, 150, 450)
l3 = drw.Create2DLine(-50, -50, 350, 350)
l4 = drw.Create2DLine(-50, 350, 350, -50)
print(f"Scene: {drw.GetGeoCount()} geos")

blen = rect.Length
drw.SetGeosSelected(False)

# Collect original line refs by position
refs = []
p = drw.GetFirstGeo()
while p:
    if abs(p.Length - blen) > 0.1:
        refs.append(p)
    try: p = p.GetNext()
    except: break

# Select boundary
rect.Selected = True

# Break each line
for line in refs:
    if line.TestIntersectPath(rect, 0, 0):
        segs = line.BreakWithCuttingGeos()
        if segs:
            print(f"Line len={line.Length:.0f}: broke into {segs.Count}")

print(f"After break: {drw.GetGeoCount()} geos")

# Collect all current geos
all_geos = []
p = drw.GetFirstGeo()
while p:
    all_geos.append(p)
    try: p = p.GetNext()
    except: break

# Delete external segments
deleted = 0
drw.SetGeosSelected(False)
for geo in all_geos:
    if abs(geo.Length - blen) < 0.1:
        continue
    h = geo.Length / 2
    r = geo.PointAtDistanceAlongPathG(h, 0, 0, 0, None)
    if isinstance(r, tuple) and r[0]:
        if not rect.IsPointInside(r[1], r[2]):
            try:
                geo.Erase()
                deleted += 1
            except Exception as e:
                print(f"Erase failed: {e}")

print(f"Deleted: {deleted} external segments")
print(f"Final: {drw.GetGeoCount()} geos")
drw.SetGeosSelected(False)
drw.ZoomAll()
print("DONE")
