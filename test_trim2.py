"""Test TrimWithCuttingGeos properly."""
import sys, win32com.client, time

sys.path.insert(0, r"C:\Users\C\AppData\Roaming\reasonix\global-workspace\.reasonix\skills\alphacam-bridge")

# Start fresh
time.sleep(1)
app = win32com.client.Dispatch("aroutaps.Application")
app.Visible = True
time.sleep(0.5)

# Use current active drawing (don't call New())
drw = app.ActiveDrawing
print("Drawing:", drw.Name)

# Delete existing geometries
p = drw.GetFirstGeo()
while p is not None:
    try:
        p.Erase()
        p = drw.GetFirstGeo()
    except:
        break

time.sleep(0.3)

# Create test scene
rect = drw.CreateRectangle(0, 0, 300, 300)
line = drw.Create2DLine(-150, 150, 450, 150)
print(f"Created: {drw.GetGeoCount()} geos")
print(f"Rect length: {rect.Length}, Line length: {line.Length}")

blen = rect.Length

# Select boundary
drw.SetGeosSelected(False)
rect.Selected = True

# Get line start point
elem = line.GetFirstElem()
sx = elem.StartXG
sy = elem.StartYG
print(f"Line start: ({sx}, {sy})")

# Try TrimWithCuttingGeos
print("Calling TrimWithCuttingGeos...")
result = line.TrimWithCuttingGeos(sx, sy)
print(f"Result: {result}")

print(f"Final geos: {drw.GetGeoCount()}")
drw.SetGeosSelected(False)
drw.ZoomAll()
print("DONE")
