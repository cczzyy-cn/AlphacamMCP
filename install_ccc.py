"""Install CCC功能 VBA modules into AlphaCAM via COM."""
import sys, win32com.client, os

base = os.path.dirname(os.path.abspath(__file__))
ccc_dir = os.path.join(base, "CCC功能")

# Standard modules (.bas)
modules = ["Events", "modTrim", "modOffset", "modSort", "modMirror", "modRamp", "modMirrorPath"]
# Form modules (.txt — code only, controls must be set up separately)
forms = ["frmRamp"]

app = win32com.client.Dispatch("aroutaps.Application")
vbe = app.VBE
proj = vbe.ActiveVBProject

# Install standard modules
for name in modules:
    # Remove existing module if present
    for comp in list(proj.VBComponents):
        if comp.Name == name:
            proj.VBComponents.Remove(comp)
            print(f"Removed existing: {name}")
            break
    
    # Add module and read code
    module = proj.VBComponents.Add(1)  # 1 = standard module
    module.Name = name
    path = os.path.join(ccc_dir, name + ".bas")
    with open(path, "r", encoding="gb2312") as f:
        code = f.read()
    module.CodeModule.AddFromString(code)
    print(f"Installed: {name}.bas ({len(code)} chars)")

# Install form code (form must already exist or be created first)
for name in forms:
    for comp in list(proj.VBComponents):
        if comp.Name == name:
            proj.VBComponents.Remove(comp)
            print(f"Removed existing form: {name}")
            break

    # Try to create form (3 = vbex_static_frm)
    try:
        form = proj.VBComponents.Add(3)
        form.Name = name
        path = os.path.join(ccc_dir, name + ".txt")
        with open(path, "r", encoding="gb2312") as f:
            code = f.read()
        form.CodeModule.AddFromString(code)
        print(f"Installed: {name}.txt as form ({len(code)} chars)")
        if name == "frmRamp":
            print(f"  ⚠ 请在 AlphaCAM VBA 编辑器中手动添加以下控件到 {name} 窗体：")
            print(f"      - ComboBox: cmbMethodTool")
            print(f"      - TextBox:  txtMinSize (默认250)、txtCutDepth (默认18)、txtRampAngle (默认30)")
            print(f"      - CommandButton: cmdOK (确定)、cmdCancel (取消)")
        elif name == "frmMirrorPath":
            print(f"  Warning: frmMirrorPath is no longer used. Install modMirrorPath.bas only.")
    except Exception as e:
        print(f"  ⚠ 无法创建窗体 {name}：{e}")
        print(f"    请手动导入：VBA编辑器 → 插入 → 用户窗体 → 导入文件 {name}.txt")

print("\nDone! Restart AlphaCAM or switch to it to test.")
