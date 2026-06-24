"""Test grid line trimming with pthOffset boundary."""
import sys, win32com.client, time

app = win32com.client.Dispatch("aroutaps.Application")
app.Visible = True
time.sleep(0.5)
app.New()
drw = app.ActiveDrawing
print(f"Fresh drawing: {drw.GetGeoCount()} geos")

# Step 1: Create the arch shape (pthIn)
L, R = 50, 350  # leftEdge, rightEdge
B, T = 50, 300   # bottom, top
CX = (L + R) / 2

print("Creating arch shape...")
geo = drw.Create2DGeometry(L, B)
geo.AddLine(L, T)
geo.AddLine(CX - 40, T)
geo.AddArc2Point(CX, T + 50, CX + 40, T)  # arch
geo.AddLine(R, T)
geo.AddLine(R, B)
pthIn = geo.CloseAndFinishLine()
pthIn.Fillet(10)
print(f"Arch: closed={pthIn.Closed}, len={pthIn.Length:.0f}")

# Step 2: Create offset (pthOffset)
print("Creating offset...")
offsetPaths = pthIn.Offset(15, 1)
pthOffset = offsetPaths.Item(1)
print(f"Offset: closed={pthOffset.Closed}, len={pthOffset.Length:.0f}")

# Step 3: Create X grid lines (same logic as template)
import math
spacing = 30

def draw_x(cx, cy):
    l1 = drw.Create2DLine(cx, cy, R, cy + (R - cx))
    l2 = drw.Create2DLine(cx, cy, L, cy + (cx - L))
    l3 = drw.Create2DLine(cx, cy, R, cy - (R - cx))
    l4 = drw.Create2DLine(cx, cy, L, cy - (cx - L))
    return [l1, l2, l3, l4]

print(f"\nDrawing X grid from Y={B} to Y=350, spacing={spacing}...")
all_lines = []
y = B
while y <= 350:
    lines = draw_x(CX, y)
    all_lines.extend(lines)
    y += spacing

print(f"Total grid lines: {len(all_lines)}")

# Step 4: Test intersection and midpoint
print(f"\n=== Testing trimming ===")
print(f"pthOffset.IsPointInside({CX}, {B+10}): {pthOffset.IsPointInside(CX, B+10)}")

kept = 0
deleted = 0
for gl in all_lines:
    h = gl.Length / 2
    r = gl.PointAtDistanceAlongPathG(h, 0, 0, 0, None)
    if isinstance(r, tuple) and r[0]:
        mx, my = r[1], r[2]
        inside = pthOffset.IsPointInside(mx, my)
        intersects = gl.TestIntersectPath(pthOffset, 0, 0)
        
        if inside or intersects:
            kept += 1
            # Trim outside parts
            elem = gl.GetFirstElem()
            sx, sy = elem.StartXG, elem.StartYG
            elem2 = gl.GetLastElem()
            ex, ey = elem2.EndXG, elem2.EndYG
            
            drw.SetGeosSelected(False)
            pthOffset.Selected = True
            
            if not pthOffset.IsPointInside(sx, sy):
                gl.TrimWithCuttingGeos(sx, sy)
                # re-get end point
                drw.SetGeosSelected(False)
                pthOffset.Selected = True
                elem2 = gl.GetLastElem()
                ex, ey = elem2.EndXG, elem2.EndYG
            
            if not pthOffset.IsPointInside(ex, ey):
                gl.TrimWithCuttingGeos(ex, ey)
            
            drw.SetGeosSelected(False)
        else:
            gl.Erase()
            deleted += 1
            
print(f"Kept: {kept}, Deleted: {deleted}")
print(f"Final geos: {drw.GetGeoCount()}")
drw.ZoomAll()
print("DONE")
