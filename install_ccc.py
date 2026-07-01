"""Install CCC功能 VBA modules into AlphaCAM via COM."""
import sys, win32com.client, os

base = os.path.dirname(os.path.abspath(__file__))
ccc_dir = os.path.join(base, "CCC功能")
modules = ["Events", "modTrim", "modOffset", "modSort", "modMirror"]

app = win32com.client.Dispatch("aroutaps.Application")
vbe = app.VBE
proj = vbe.ActiveVBProject

for name in modules:
    # Remove existing module if present
    for comp in list(proj.VBComponents):
        if comp.Name == name:
            proj.VBComponents.Remove(comp)
            print(f"Removed existing: {name}")
            break
    
    # Add module and read code
    module = proj.VBComponents.Add(1)
    module.Name = name
    path = os.path.join(ccc_dir, name + ".bas")
    with open(path, "r", encoding="gb2312") as f:
        code = f.read()
    module.CodeModule.AddFromString(code)
    print(f"Installed: {name}.bas ({len(code)} chars)")

print("\nDone! Restart AlphaCAM or switch to it to test.")
