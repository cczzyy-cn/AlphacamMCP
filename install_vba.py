"""Install TrimByBoundary VBA macro into AlphaCAM via COM."""
import sys, win32com.client

sys.path.insert(0, r"C:\Users\C\AppData\Roaming\reasonix\global-workspace\.reasonix\skills\alphacam-bridge")

code_path = r"C:\Users\C\AppData\Roaming\reasonix\global-workspace\.reasonix\skills\alphacam-bridge\TrimByBoundary.bas"

with open(code_path, "r", encoding="utf-8") as f:
    code = f.read()

app = win32com.client.Dispatch("aroutaps.Application")

try:
    vbe = app.VBE
    proj = vbe.ActiveVBProject
    for comp in proj.VBComponents:
        if comp.Name == "TrimByBoundary":
            proj.VBComponents.Remove(comp)
            break
    module = proj.VBComponents.Add(1)
    module.Name = "TrimByBoundary"
    module.CodeModule.AddFromString(code)
    print("SUCCESS: VBA macro TrimByBoundary installed into AlphaCAM!")
    print("Press Alt+F8, select TrimByBoundary, click Run.")
except Exception as e:
    print("VBE method failed:", e)
    print()
    print("Manual install:")
    print("1. Alt+F11 in AlphaCAM")
    print("2. Insert -> Module")
    print("3. Copy code from:")
    print("   " + code_path)
    print("4. Close editor")
    print("5. Alt+F8 -> TrimByBoundary -> Run")
