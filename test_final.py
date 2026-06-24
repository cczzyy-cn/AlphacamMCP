"""Test TrimWithCuttingGeos in a single clean run."""
import sys, win32com.client, time

app = win32com.client.Dispatch("aroutaps.Application")
app.Visible = True
time.sleep(1)

drw = app.ActiveDrawing

# Delete existing geos safely
p = drw.GetFirstGeo()
while p:
    try:
        p.Erase()
        p = drw.GetFirstGeo()
    except:
        break
time.sleep(0.5)

# Create rectangle boundary
rect = drw.CreateRectangle(0, 0, 300, 300)
print("Rectangle created:", rect.Length)

# Create one crossing line
line = drw.Create2DLine(-150, 150, 450, 150)
print("Line created:", line.Length)
print("Total geometries:", drw.GetGeoCount())

blen = rect.Length

# Select boundary
drw.SetGeosSelected(False)
rect.Selected = True

# Get line start point
elem = line.GetFirstElem()
sx = elem.StartXG
sy = elem.StartYG
print(f"Line start point: ({sx}, {sy}) - outside rectangle")

# TrimWithCuttingGeos - deletes the part containing (sx, sy)
line.TrimWithCuttingGeos(sx, sy)
print("TrimWithCuttingGeos completed")

print("Final geometries:", drw.GetGeoCount())

# Check what we have now
p = drw.GetFirstGeo()
idx = 1
while p:
    inside = rect.IsPointInside(p.Length/2, 0) if idx == 1 else False
    print(f"  Geo {idx}: length={p.Length:.0f}, closed={p.Closed}")
    idx += 1
    try: p = p.GetNext()
    except: break

drw.SetGeosSelected(False)
drw.ZoomAll()
print("DONE - check the AlphaCAM window")
