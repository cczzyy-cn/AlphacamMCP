"""Test trim_with_boundary logic against live AlphaCAM."""
import sys, win32com.client
sys.path.insert(0, r'C:\Users\C\AppData\Roaming\reasonix\global-workspace\.reasonix\skills\alphacam-bridge')

app = win32com.client.Dispatch('aroutaps.Application')
app.Visible = False
app.New()
drw = app.ActiveDrawing

# Scene: rectangle + 4 crossing lines
rect = drw.CreateRectangle(0, 0, 300, 300)
l1 = drw.Create2DLine(-150, 150, 450, 150)
l2 = drw.Create2DLine(150, -150, 150, 450)
l3 = drw.Create2DLine(-50, -50, 350, 350)
l4 = drw.Create2DLine(-50, 350, 350, -50)

blen = rect.Length
drw.SetGeosSelected(False)
rect.Selected = True

# Collect lines that intersect boundary
lines = []
p = drw.GetFirstGeo()
while p is not None:
    if abs(p.Length - blen) > 0.001:
        if p.TestIntersectPath(rect, 0, 0):
            lines.append(p)
    try:
        p = p.GetNext()
    except:
        break

# Break at intersections and delete external segments
deleted = 0
for line in lines:
    segs = line.BreakWithCuttingGeos()
    if segs:
        for j in range(1, segs.Count + 1):
            seg = segs.Item(j)
            r = seg.PointAtDistanceAlongPathG(seg.Length / 2, 0, 0, 0, None)
            if isinstance(r, tuple) and r[0] and not rect.IsPointInside(r[1], r[2]):
                seg.Erase()
                deleted += 1

print('TEST_RESULT: deleted=%d' % deleted)
print('TEST_PASS' if deleted == 4 else 'TEST_FAIL')
app.New()
app.Visible = False
