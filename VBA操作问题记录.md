# AlphaCAM VBA 操作问题记录

记录通过 COM 自动化（`win32com` + `aroutaps.Application`）操作 AlphaCAM 2016 R1 内置 VBA 时遇到的各种问题、根因与解决方案。
适用对象：`AdoorMain` 门样式宏（`AdoorEvents` 模块）、CDM 门板项目、一般 AlphaCAM VBA 脚本。

---

## 1. 连接与实例

### 1.1 ProgID 是 `aroutaps.Application`，不是 `AlphaCAM.Application`

**现象：** `win32com.client.GetActiveObject('AlphaCAM.Application')` 报
`com_error: (-2147221005, '无效的类字符串', ...)`。

**根因：** AlphaCAM 2016 的 COM ProgID 是 `aroutaps.Application`。

**解决：**

```python
app = win32com.client.GetActiveObject('aroutaps.Application')   # 连接已运行实例
# 或 app = win32com.client.Dispatch('aroutaps.Application')     # 启动/连接
```

### 1.2 多个 AlphaCAM 进程 → 连错实例

**现象：** 同一台机器开了多个 `Acam.exe`（任务管理器可见多个进程），`GetActiveObject`
返回的实例可能不是期望的那个；A 实例能访问 VBE、B 实例能画图，两边状态不一致。

**判断方法：**

```powershell
Get-Process -Name Acam | Select-Object Id, MainWindowTitle
```

VBA 编辑器打开的实例其窗口标题类似
`Microsoft Visual Basic for Applications - C:\...项目名 - [模块名 (代码)]`。

**解决：** 先确认目标实例再操作。MCP bridge 与独立 Python 脚本各自 `GetActiveObject`
可能连到不同实例——出现"bridge 读不到模块 / 脚本画不了图"等诡异现象时，优先怀疑实例错位。

**⚠️ 2026-09-13 补充（危险教训）：`GetActiveObject` 连到的实例，往往不是屏幕上那个可见窗口。**
当天实测：

| 实例 | 窗口 |
|---|---|
| pid 16636 | **COM 连到的就是它**（弹窗、VBE 都在它上面），但它的主窗口**不在** `AlphaCAM_3DMILL` 类窗口列表里 |
| pid 28860 | `AlphaCAM_3DMILL`，**可见**，标题 `3D 5-轴鉋花机专业版: 9-9测试自动化` |
| pid 14964 | 同名窗口但不可见 |

⇒ **绝对不要用"键盘/鼠标"去验证 COM 操作的效果** —— 我先用 `FindWindow("AlphaCAM_3DMILL")`
找窗口，找到的是 **28860**；如果那次 `SetForegroundWindow` 成功，Ctrl+Z 就会**撤销 28860 里那张
真实图纸**（而不是我在 16636 里的测试图）。
而且 Windows 前台锁定会让 `SetForegroundWindow` 静默失败——当天连发 5 次 Ctrl+Z，
**全部落到了前台的应用（WPS Office）身上**。所以：

- **只用 COM/文件两条路验证**（读回对象属性、读回文件），必要时用 `see` 截图**看**；
- 需要按 Ctrl+Z / 点按钮这类"人机"手段时，**交给用户手动做**；
- 若确实要激活窗口，先 `EnumWindows` 按 **PID 匹配**（不要只按窗口类/标题），
  且必须**先读回 `GetForegroundWindow` 确认真的切过去了**再发按键。

### 1.3 `VBE` / `VBComponents` 访问失败（`'NoneType' object has no attribute 'VBComponents'`）

**现象：** MCP 工具 `list_vba_modules` 报
`Failed to list VBA modules: 'NoneType' object has no attribute 'VBComponents'`；
`run_vba_line` 同样失败。

**根因：** bridge 连接实例的 `app.VBE.ActiveVBProject` 返回 `None`（该实例的 VBA 环境
未初始化，或连错了实例）。**不等于**模块不存在。

**解决：**
- 直接脚本化访问：`vbe = app.VBE; proj = vbe.ActiveVBProject`，遍历 `proj.VBComponents`。
- 模块通常就在 `ActiveVBProject`（VBA 编辑器当前打开的项目）里，无需遍历全部项目。
- 若遍历 `vbe.VBProjects` 遇"工程已被保护"（CDM 等），跳过该项目的组件访问即可。

### 1.4 CDM 项目"受保护"与"可读"

**现象：** 遍历 `vbe.VBProjects` 时访问 `CDM` 项目的 `VBComponents` 抛错：
`该工程已被保护，不能执行操作`（VbLR6.chm, 50289）——但**有时又能读**。

**处理：**
- 逐项目访问一律用 try/except 跳过受保护项目（`'CDM' in p.Name` 时先试探）。
- CDM 项目解锁后可直接读源码：`Make`（门板生成主逻辑，7926 行）、`UserStyleTestMain`
  （用户样式测试）、`CDoor`、`COuterToolpath`、`CPathData`、`CUserStyle`、`Database` 等。
- 注意 `ActiveVBProject` 会随 VBA 编辑器当前选中项目变化——定位模块要遍历所有项目，
  不要假设活动项目。

**规律（2026-09-13 用户现场确认）——不是「主动打开过」的插件，其 VBA 工程一律是保护状态：**
- 本次 27 个工程的实测完全吻合：**只有 `CCC功能` 可读写** —— 因为它在 VBA 编辑器里被**主动打开过**
  （编辑器标题当时显示 `CCC功能 - frmRamp (UserForm)`）；其余 **26 个（含 `CDM`）都是 AlphaCAM
  启动时被动加载**的，全部受保护。
- 这也解释了「同一天里 `CDM` 先能写、重启后不能写」：上午它是被主动打开的状态，重启后变回被动加载。
- ⇒ **想改某个插件，先在 VBA 编辑器里把它打开（激活任一组件）再试读写**；
  否则一律按**只读**处理。不要把"这次能写"当成持久属性。
- 被动加载的工程连 `VBComponents` 都读不了：`running_snapshot.py` / `component_deploy.py` 会抛
  `该工程已被保护`。已改为**明确报错并提示**「先在 VBA 编辑器里 工具 → <工程>属性 → 保护 解除」，
  不再是裸 traceback。
- 改完记得把保护加回去（`.arb` 会存下保护状态，下次启动即生效）。

---

## 2. 宏调用（`Application.Run`）

### 2.1 宏名必须是 `Project.Module.Macro` 完整格式

**现象：** `app.Run('AdoorEvents.Sindeg')`、`app.Run('Sindeg')` 都失败（E_FAIL），
带完整项目名后成功。

**解决：**

```python
proj = app.VBE.ActiveVBProject
res = app.Run(proj.Name + '.AdoorEvents.Sindeg', 30.0)   # → 0.5
```

### 2.2 新插入的宏无法通过 `Run` 调用（"未找到所需的有效名称"）

**现象：** 往模块里 `InsertLines` / `AddFromString` 一个新 `Sub`，随后
`app.Run('项目.模块.新宏')` 报
`(-2147352567, '发生意外。', (0, 'APC.ApcHost.7', '未能找到所需的有效名称。', ...))`；
而模块里**已存在**的宏运行正常。

**根因：** AlphaCAM 2016 的 `Application.Run` 基于项目加载/编译时的宏名表，
**运行时动态添加的宏不会被解析到**。这是平台限制，不是代码错误。

**影响与对策：**
- 不影响实际使用：`AdoorMain` 等项目加载时就存在的宏，修改其**代码体**后运行的是新代码（见 2.3）。
- 想验证新代码，用"已有宏换体"技巧（见 5.1），或直接修改已存在宏再改回。

### 2.3 修改已存在宏的代码体 → 运行时立即生效

**验证方法：** 临时把 `Sindeg` 函数体改为 `Sindeg = 42`，`Run` 后返回 `42.0`；
改回原样返回 `0.5`。

**结论：** `Run` 调用时按**当前模块代码**编译执行，修改 `AdoorMain` 等已有宏的
函数体无需重新加载项目即生效。

### 2.4 `Run` 报 E_FAIL（`0x80004005`）通用错误

**现象：** `(-2147352567, '发生意外。', (0, None, None, None, 0, -2147467259), None)`。

**排查顺序：**
1. 宏名格式是否 `Project.Module.Macro`（最常见）。
2. 模块是否被破坏（残留游离代码 / 语法错误导致项目编译失败，见 3.4、5.3）。
3. 是否新插入的宏（见 2.2）。

---

### 2.5 新装模块 Run 报"VBA 在编译时遇到错误"（APC.ApcHost.7）→ 过程名下划线开头

**现象：** `install_vba_module` 添加成功，但 `Run("CDM.模块名.过程名")` 报：
```
(-2147352567, '发生意外。', (0, 'APC.ApcHost.7', 'VBA 在编译时遇到错误。\r\n', None, 0, -2147467259), None)
```
与 2.4 的 E_FAIL 不同，**错误来源是 `APC.ApcHost.7`**，且宏名格式、工程编译均正常。

**根因：** 临时模块里过程名 `Public Sub _MCP_Run()` **以下划线开头**——VBA 标识符（模块名/过程名/变量名）**必须以字母开头**，下划线开头是编译错误。工程内已有宏不受影响（不重新编译），只有新模块编译时暴露。

**解决：** 过程名改字母开头（如 `MCPRun`）。**模块名同理**：`module.Name = "_MCP_TEMP_xxx"` 会赋值失败，模块保持默认名"模块N"，且按原名清理找不到 → 每次失败残留一个"模块N"。

**配套教训（alphacam_com.py `run_vba_line` 连环 bug，已修复）：**
1. 模块名 `_MCP_TEMP_...` 下划线开头 → Name 赋值失败 → 残留"模块N"（改 `MCP_TEMP_...`）
2. 过程名 `_MCP_Run` 下划线开头 → "VBA 在编译时遇到错误"（改 `MCPRun`）
3. 宏名缺工程前缀 `Project.Module.Macro` → E_FAIL（见 2.1，加 `proj.Name & "."`）

**验证：** 修复后连续 5 次 `run_vba_line` 成功、零残留（组件数恢复原始值）。

### 2.6 `Run` 最多 10 个位置参数 → 多参过程要注入临时模块调用（2026-09-13）

**现象：** `app.Run("CCC功能.modRamp.ApplyRampEntry", *args)` 传 9 个实参时报：
```
IAlphaCamApp.Run() takes from 1 to 10 positional arguments but 11 were given
```
即 `Run` 的签名上限是 **self + 宏名 + 9 个可选实参**。要调用**参数更多的过程**就撞墙。

**解决：** 把调用写成 **VBA 代码**，注入临时模块后 `Run` 那个无参过程 —— VBA 侧没有实参个数限制：

```python
proj = app.VBE.ActiveVBProject
mod = proj.VBComponents.Add(1)                       # 1 = 标准模块
mod.Name = "MCP_TEST_" + uuid.uuid4().hex[:8]        # 必须字母开头(见 2.5)
mod.CodeModule.AddFromString(
    "Public Sub MCPRun()\n" + "modRamp.ApplyRampEntry 0,18,10,\"\",\"\",0,True,0.8,True" + "\nEnd Sub")
try:
    app.Run("%s.%s.MCPRun" % (proj.Name, mod.Name))
finally:
    proj.VBComponents.Remove(mod)                    # 传组件对象, 不是名字
```

**注意：**
- 临时模块加在 **`ActiveVBProject`** 上 —— 跑之前确认活动工程就是你要动的那个（VBA 编辑器里
  切一下文件就变，见 1.4）。本次目标工程是 **`CCC功能`**（不是 CDM）。
- `Remove` 要传**组件对象**；传名字会抛错，异常路径下记得清理以免残留"模块N"。

---
## 3. 模块代码读写（CodeModule）

### 3.1 读取：`Lines(1, CountOfLines)` 行尾是 CRLF

```python
n = cm.CountOfLines
code = cm.Lines(1, n)          # 行分隔符为 \r\n，最后一行后无分隔符
line = cm.Lines(ln, 1)         # 取单行（不含行尾符）
```

### 3.2 写入：`DeleteLines` 全量替换 + `AddFromString`

推荐全量替换（避免残留）：

```python
cm.DeleteLines(1, cm.CountOfLines)
cm.AddFromString(code)         # 行尾 CRLF 或 LF 均可，VBA 自动规范化
```

实测 CRLF 与 LF 行尾的 `AddFromString` 行为一致，均按内容行数计行。

### 3.3 Python 文本模式写入 → `\r\r\n` 污染（重要）

**现象：** COM 读出的代码行尾是 `\r\n`；用 Python `open(path, 'w')`（默认文本模式）
写入时 `\n → \r\n`，文件里变成 `\r\r\n`（双 CR）。之后读回做字符串匹配/替换全部落空。

**解决：**

```python
# 写备份时禁用二次转换
with open(path, 'w', encoding='utf-8', newline='') as fh:
    fh.write(code)

# 从文件读回时统一规范化行尾
import re
code = re.sub(r'\r\r\n|\r\n|\r', '\n', open(path, encoding='utf-8').read())
```

**教训：** 字符串模式匹配（`str.count`/`replace`）前，先确认行尾；报"pattern not
found"时第一反应检查 `\r`。

### 3.4 删除宏必须删整个块，不能只删两行

**现象：** 清理测试宏时用 `DeleteLines(ln, 2)`（Sub 头 + 下一行），宏体全部残留，
模块里堆满无头的游离代码（`Set f = ...`、`With ... End With` 在模块级非法）→
项目编译失败，后续一切 `Run` 都报"未找到所需的有效名称"。

**解决：** 定位 `Sub` 声明行到对应 `End Sub`，整块删除：

```python
start = <找到 'Public Sub Xxx' 的行>
end = start
while cm.Lines(end, 1).strip() != 'End Sub':
    end += 1
cm.DeleteLines(start, end - start + 1)
```

或用"整模块重写"（3.2）兜底，彻底清除残留。

### 3.5 VBA 会**重写长小数字面量** → 部署读回校验永远失败（2026-09-13，重要）

**现象：** 仓库文件里写 `Private Const DEG2RAD As Double = 0.0174532925199433`，
`AddFromString` 写入并保存后，**从运行工程读回来的文本变成了**
`Private Const DEG2RAD As Double = 1.74532925199433E-02`。
于是 `component_deploy.py` 的"读回校验"永远不相等 → 每次都判定部署失败并回滚；
即使绕过校验，此后每次 `audit` 也永远显示 `==repo=False`。

**根因：** VBA 在保存模块时会**规范化数字字面量**（本次是转成 15 位有效数字的科学计数法）。
这不是行尾/大小写那种可以靠"忽略大小写"绕过的差异 —— 它是**字符内容**变了。

**规避（推荐第一种）：**
1. **不要在源码里写长小数字面量**，改用运行时计算：
   ```vba
   Private Function Pi() As Double
       Pi = 4 * Atn(1)
   End Function
   Private Function Deg2Rad() As Double
       Deg2Rad = Pi() / 180
   End Function
   ```
   （`modRamp.bas` v2.0 就是这么改的，改完读回校验立刻 `verify=True`。）
2. 若必须写字面量，就按 VBA 规范化的形态写（科学计数法、15 位有效数字），
   但这需要先实测一次它到底重写成什么。

**排查手法：** 把"仓库文本"和"运行版文本"逐行做**大小写归一化后的**对比，
真实差异会自己浮出来（本次 45 处差异里只有 1 处是真实差异，其余全是大小写/标签名归一化）：
```python
sm = difflib.SequenceMatcher(None, [l.lower() for l in repo], [l.lower() for l in running])
```

**同批发现的其它归一化（这些不影响校验，因为比较是大小写不敏感的）：**
`.Count` → `.count`、`.FinalDepth` → `.finalDepth`、标签 `NextTp:` → `nextTP:`、
局部变量名 `scX` → `scx`（VBA 按**声明处**的大小写统一）。

---

## 4. VBA 语言陷阱

### 4.1 `Dim` 不能写在循环体/条件块内

**现象：** `Dim ta As Path` 写在 `For` 循环内、`Dim nc As Collection` 写在 `If` 块内，
AlphaCAM 报"声明重复"编译错误。

**根因：** VBA 把所有 `Dim` 提升到过程级，同名变量出现在不同块中会被误判重复声明。

**解决：** 所有 `Dim` 集中在过程顶部；块内只保留赋值（去掉 `Dim` 关键字）。

### 4.2 `MsgBox` 会阻塞自动化 —— 但只阻塞**调用方**（2026-09-13 三条实测）

**现象：** 宏里 `MsgBox` 弹窗后，发起调用的那个进程会一直等（`app.Run` 不返回），表现为"卡住"。

**实测的三条规律（都很反直觉，务必记住）：**

| # | 结论 | 证据 |
|---|---|---|
| 1 | **只阻塞调用方，不阻塞 COM 本身** | 回执框还挂在屏幕上时，**另一个进程**的 COM 调用（`GetActiveObject` + 读写工程）正常返回。与 §7.8 的 `ReadTextFile` 现象一致 |
| 2 | **`WM_COMMAND`/`IDOK` 关不掉它** | 对 `#32770` 窗口连发 8 次 `PostMessage(WM_COMMAND, IDOK)` 毫无作用，窗口一直在 |
| 3 | **对「确定」按钮发 `BM_CLICK` 一次即关** | `SendMessage(button_hwnd, 0x00F5, 0, 0)`；先 `EnumChildWindows` 找 `class == "Button"` 的子窗口 |

**⇒ `tools/win_ctl.py --dismiss <pid>` 目前用 `WM_COMMAND/IDOK`，对 AlphaCAM 的 VBA `MsgBox` 是失效的**，
应改为 `BM_CLICK`。另外它只按窗口类 `#32770` 过滤，**会误伤其它程序**（2026-09-13 实测误点了
360压缩 的三个解压进度窗）—— 一律**按标题精确匹配**。

**解决（自动化时的正确姿势）：**
- 验证用代码内不要 `MsgBox`；需要输出信息时写文件（见 5.2）。**这是在 VBA 里最省事的做法。**
- 若被调用的是**已发布的插件**（不能为测试改它的代码，例：`modRamp` 完成时必弹回执框）：
  1. **调用放后台作业**跑（`python -u`），别让前台调用被卡死；
  2. **另起一个独立监视进程**轮询目标标题的窗口，**先把窗口文本记下来**再点掉它。
     回执框内容往往就是唯一的结果摘要（本次靠它读到 `候选/受理/已应用/跳过/微连接/小件降速` 计数）；
  3. 见 §5.5 的参考实现。

### 4.3 浮点"等于"判断

`L0orR1 = 0` 这类判断：`L0orR1` 是用户变量直接赋值时无计算误差，`= 0` 可用；
若值经过运算，用 `Abs(L0orR1) < 0.0000001` 更稳。

### 4.4 错误处理器里 `On Error GoTo 0` 会清空 `Err`

**现象：** 宏出错进 `EH:` 后先 `On Error Resume Next` 再 `On Error GoTo 0`，
最后 `Print Err.Number` 输出 `0`（错误信息丢失）。

**根因：** `On Error GoTo 0` 会清除当前 `Err` 对象。

**解决：** 进入 EH 立即把错误保存到变量，再处理文件/清理：

```vba
EH:
   Dim en As Long, ed As String
   en = Err.Number: ed = Err.Description
   On Error Resume Next
   Close #1
   Open "C:\path\out.txt" For Output As #1
   Print #1, "ERR " & en & " | " & ed
   Close #1
```

### 4.5 `Drawing` 没有 `Count` 属性

**现象：** 用 `App.ActiveDrawing.Count` 取几何数报 438（对象不支持属性/方法）；
`Toolpaths.Count` 正常。

**解决：** 几何数用 `GetFirstGeo()` 遍历计数（参考 `get_all_geometries` 实现），
刀路数用 `Drawing.Toolpaths.Count`。

---

### 4.6 VBA 模块名/过程名必须以字母开头（下划线开头非法）

**规则：** VBA 标识符（模块名、过程名、变量名、常量名）**必须以字母开头**（A-Z/a-z），后跟字母/数字/下划线；**下划线开头（如 `_MCP_Run`、`_MCP_TEMP_`）编译错误**。

**实际影响（自动化场景）：**
- `VBComponents.Add(1)` 后 `module.Name = "_xxx"` 会静默失败（不抛错），模块保持默认名"模块N"
- 代码里写 `Public Sub _xxx()` → 编译错误"VBA 在编译时遇到错误"
- 生成动态模块/过程名时：**统一字母开头**，如 `MCP_TEMP_<hex>` / `MCPRun`

**排查线索：** 新增模块才报编译错误、已有模块正常 → 先查动态生成的模块名/过程名是否符合标识符规则。

---
### 4.7 `Or`/`And` 不短路：`Ni Is Nothing Or Ni.Sheets.Count` 会报错误 91

**现象：** 宏运行弹出 VBA 运行时错误框"运行时错误 '91'：对象变量或 With 块变量未设置"，VBE 进入中断状态（标题显示"[正在运行]"）。

**根因（两个叠加）：**
1. **VBA 的 `Or`/`And` 不短路求值**：`If Ni Is Nothing Or Ni.Sheets.Count = 0 Then` 中，
   即使 `Ni Is Nothing` 为 True，仍会继续求值 `Ni.Sheets.Count`——`Ni` 为 Nothing 时 `Ni.Sheets` 触发错误 91。
   **必须分开判断**：
   ```vba
   If Ni Is Nothing Then Exit Sub   ' 先判 Nothing
   If Ni.Sheets.Count = 0 Then Exit Sub
   ```
2. **`On Error GoTo 0` 禁用过程全部错误处理**：局部 `On Error Resume Next` 后若用
   `On Error GoTo 0` 恢复，会**清掉过程原有的 `On Error GoTo EH`**，后续任何错误直接弹出
   VBA 错误框（不被 EH 捕获）。恢复主错误处理要用 **`On Error GoTo EH`**，不是 `GoTo 0`。

**修复模板：**
```vba
Set Ni = Nothing
On Error Resume Next
Set Ni = ActiveDrawing.GetNestInformation   ' 可能失败，吞掉
On Error GoTo EH                            ' ← 恢复主错误处理器（勿用 GoTo 0）
If Ni Is Nothing Then MsgBox "...": Exit Sub
If Ni.Sheets.Count = 0 Then MsgBox "...": Exit Sub
```

---

### 4.8 `win32com` 动态绑定：**方法必须加括号**，否则报错极具误导性（2026-09-13）

**现象：** 三种看起来完全不同的报错，其实是同一个原因：
```
AttributeError: 'function' object has no attribute 'Name'
AttributeError: 'method' object has no attribute 'ToolInOut'
AttributeError: 'method' object has no attribute 'Sheets'
```
**根因：** 用 `GetActiveObject`（后期绑定）时，**不加括号访问一个方法**拿到的是
"绑定方法对象"本身，而不是它返回的值。VBA 里可以省略无参调用的括号（`Set ni = drw.GetNestInformation`），
**Python 不行**。

**判据（记住了能省很多时间）：** 报错里出现 `'function' object has no attribute` 或
`'method' object has no attribute` → **几乎一定是漏了括号**。本次连踩三次：

| 写错 | 写对 |
|---|---|
| `app.GetCurrentTool` | `app.GetCurrentTool()` |
| `drw.GetNestInformation` | `drw.GetNestInformation()` |
| `geo.Finish` | `geo.Finish()`（`Create2DGeometry` 的收尾调用）|
| `drw.GetToolPathCount` | `drw.GetToolPathCount()` |

> **判据之外的另一半**：如果是**属性**（`tp.MinXL`、`ops.Count`）就**不能**加括号，
> 加了会报另一类错。拿不准时先 `EnsureDispatch` 后用 `dir()` 看它是方法还是属性。

**顺带：后期绑定下 `dir()` 不给成员。** 想知道对象到底有什么，用 typelib 包装：

```python
import win32com.client as w
app = w.gencache.EnsureDispatch("Ar5axaps.Application")   # 注意是 Ar5axaps, 见 §1.1
print([n for n in dir(app) if not n.startswith("_")])      # 本次实测: App 100 个成员
d = app.ActiveDrawing
print([n for n in dir(d) if "Geo" in n])                   # Drawing 296 个成员
el = d.GetFirstToolPath.Elements.Item(1)
print([n for n in dir(el) if "Z" in n])                    # Element 79 个成员
```

**注意：反编译出来的接口**（如 `小条先切/AlphaCAMRouter/IElement.vb`）**不能用来断言"没有某成员"** ——
那些 `.vb` 里有大量 `_VtblGap` 占位，说明接口方法被省略了。实测 `Element` 就有
`Length` / `StartZL` / `EndZL` / `StartZG` / `EndZG` / `LeadIn` / `LeadOut` 等反编译文件里看不到的成员，
而这些正是"读回 Z 剖面做验收"的关键（见 §5.5）。

### 4.9 `Drawing.GetExtent` 的参数**不是** `(x1,y1,x2,y2,...)`：第 3 个是 **Z1**（2026-09-14 实测，重要）

**官方签名**（`tempacamapi\Objects\Drawing\Drawing_GetExtent.htm`）：

```
drw.GetExtent (X1, Y1, Z1, X2, Y2, Z2)     ' 6 个都是 ByRef Double, 顺序是 X, Y, Z, X, Y, Z
```

**犯错写法**（我按"最小 XY / 最大 XY"的直觉排的）：

```vba
drw.GetExtent gx1, gy1, gx2, gy2, 0, 0     ' ← 错
```

它把第 3 个参数当成了 X2，于是实际取回的东西是：

| 变量 | 我以为 | 实际拿到 |
|---|---|---|
| `gx1` | MinX = -2.5 | MinX = -2.5 ✓ |
| `gy1` | MinY = -2.5 | MinY = -2.5 ✓ |
| `gx2` | MaxX = 1602.5 | **Z1 = -18.0**（负的切深！）|
| `gy2` | MaxY = 302.5 | **X2 = 1602.5** |

于是 `(gx1+gx2)/2 = -10.25`、`(gy1+gy2)/2 = 800` —— 一个完全错位的"中心点"。
后面两个字面量 `0, 0` 把 Y2/Z2 丢掉了：**VBA 允许给 ByRef 出参传表达式**（静默生成临时变量），
**不报错**，所以编译器不会救你。

**后果**：`modRamp` v2.1.1 用这个中心做"从外往内"排序的基准，距离几乎被 X 方向支配，
排序退化成**「把原顺序倒过来」**——实测序列 `700, 250, 100, 600`，与"错误中心"的预测**逐位吻合**。
真机表现就是用户说的"顺序没有从外往内"。

**修法**：按真实顺序取回，并且**取回后校验**，不合法的盒子退化为自洽的兜底值：

```vba
Dim gx1 As Double, gy1 As Double, gz1 As Double
Dim gx2 As Double, gy2 As Double, gz2 As Double
drw.GetExtent gx1, gy1, gz1, gx2, gy2, gz2
If gx2 > gx1 And gy2 > gy1 Then          ' 合法性校验: X2>X1 且 Y2>Y1
    cx = (gx1 + gx2) / 2
    cy = (gy1 + gy2) / 2
    Exit Sub
End If
' 不合法 -> 用候选件联合包围盒中心(整批料的中心), 不依赖 GetExtent
```

**一般化教训（比这一个 bug 值钱）：**

1. **"取值型"COM 调用（全是 ByRef 出参）必须先去 `tempacamapi` 查签名**，别按参数名的直觉排。
   `GetExtent` / `GetFeedExtent` / `Paths.GetExtentL` 名字很像但顺序各不相同
   （只有 `Paths.GetExtentL` 才是 `(MinX, MinY, MaxX, MaxY)`）。
2. **传字面量给 ByRef 出参不报错**，所以这类错误只能靠"**把取回的值原样打出来看**"发现 ——
   本次就是先写了一个只读诊断把 6 个值全打出来才定位到的（不要靠推理猜）。
3. **对任何"基准点 / 基准值"都要加合法性校验**（盒子 `X2>X1 And Y2>Y1`、中心点在范围内……），
   不合法就走一个自洽的兜底，**不要让错误值静默传播到算法里**——否则算法"看起来在跑"，
   结果却是无意义的（这里排序确实执行了，只是按错误的基准排）。

#### 4.9.1 同一个 API 的第二个坑：`GetExtent` 的 `Y2` 会返回**确定性错值**（2026-09-14）

v2.1.2 把参数顺序改对后，兜底中心是"先试 `GetExtent`，不好用再退候选件包围盒"。
第 4 轮实机验证时发现参考中心整体偏了 1188.5mm，追下去是：

| 来源 | 值 |
|---|---|
| `drw.GetExtent(...)` | `(-2.5, -2.5, -18, 1602.5, **2679.53**, 20)` |
| 真实联合包围盒（几何+刀路，逐元素算） | `(0, 0, -18) - (1600, 300, 20)` → 中心 `(800, 150)` |

- `Y2` 应该是 **302.5**，它给的是 **2679.5339467738304**；连调 5 次、`Redraw()` 之后**都一样**
  —— **不是随机垃圾，是确定性的错值**（像是复用了某个内部字段）。
- 逐元素扫描全部几何与刀路：**没有任何实体越界**（最大 y = 300）→ 错在 API，不在图纸数据。
- 后果：`"从外往内"` 用 `|refX-x| + |refY-y|` 排序时，`refY` 偏大 1188mm 会让 Y 项支配整个距离，
  顺序退化成"按 Y 距离排"（本次数据恰好与正确顺序相同，属于运气，不能当验证通过）。

**结论（已落到 v2.2.0）**：`GetExtent` **不再用于任何参考点**，改两级兜底：

1. 有排版 → **板件几何中心**（`NestSheet.Geometry`，最准）；
2. 无排版 → **候选件联合包围盒中心**（只依赖我们自己的候选数据，永远自洽）。

`modRamp` 里现在**没有一处 `drw.GetExtent` 调用**；`tools/ccc_probe.py` 用 ABSENT 断言
`drw.GetExtent gx1` 把这个决定锁死（防止以后有人"顺手"再用回来）。

**教训升级**：一个 API 出错一次，可能是"我用错了"；**同一个 API 用两种不同方式连着坑你，
就该把它下架**，换成"只依赖自己数据的算法"，而不是继续研究它的怪毛病、打第二个补丁。

---

## 5. 验证技巧（无法直接调 `AdoorMain` 时）

### 5.1 已有宏换体测试

新插入的宏不能 `Run`（见 2.2），但**已有宏**（如 `Sindeg`）运行的是当前代码体。
临时替换其函数体为测试逻辑 → `Run` → 检查结果 → 恢复原函数体。

已验证可用于：
- 触发目标代码的编译（测试体调用了 `MirrorX`、FastGeometry 等，编译错会直接暴露）。
- 在真实 `App.ActiveDrawing` + `CreateFastGeometry` 环境跑通几何构建逻辑。
- 验证 `Path.Group` 赋值/读回、`GetNextGroupNumberForGeometries` 等。

### 5.2 用文件输出代替 MsgBox

```vba
Open "C:\path\out.txt" For Output As #1
Print #1, "value=" & x
Close #1
```

脚本侧运行宏后读取文件内容比对，不阻塞。

**坑：`Open` 后未 `Close` → 文件被 AlphaCAM 进程锁定（WinError 32），删不掉。**
宏里每次 `Open` 必须配 `Close`；被锁后用一个空宏执行无参 `Close`（关闭所有文件号）
释放句柄，再删除文件。

### 5.3 编译验证

没有公开的 VBE 编译 API；可靠做法是 `Run` 一个**已存在**的宏（如 `Sindeg`）——
运行前 VBA 会编译，若模块有语法错误会抛错。模块被破坏时连这个也会失败，
先用 3.2 整模块重写恢复。

### 5.4 定位"哪一行出错"：逐步记录法

用 `On Error Resume Next` + 每步后把 `Err.Number` 追加到文件（`For Append`），
一次运行就能定位到具体失败调用，避免"错误被 EH 覆盖"：

```vba
On Error Resume Next
App.SelectTool ...
Open "out.txt" For Append As #1
Print #1, "step2 err=" & Err.Number
Close #1
Err.Clear
...下一调用...
```

---

### 5.5 验收「已发布插件」的套路：后台调用 + 弹窗监视 + 逐元素读回（2026-09-13）

场景：要验证的插件（`CCC功能/modRamp`）**结尾必弹回执框**，且它**会删掉原刀路再重建** ——
不能为了测试改它的代码，也不能只凭回执框的计数就下结论。

**三段式：**

1. **调用放后台**（`python -u`）：`app.Run` 会一直等回执框被关掉，前台跑必然"卡住"。
2. **独立进程盯弹窗**：轮询**目标标题**的窗口（**绝不按窗口类匹配**，见 §4.2），
   **先把窗口里的 `Static` 文本记进日志**，再对 `Button` 子窗口发 `BM_CLICK`。
   这份日志就是每个用例的结果摘要（本次读到 `候选/受理/已应用/跳过/微连接/小件降速` 计数）。
3. **逐元素读回做真正的验收**：
   ```python
   els = tp.Elements                      # tp = 处理后的刀路
   for i in range(1, els.Count + 1):
       e = els.Item(i)
       (e.StartXL, e.StartYL, e.StartZL, e.EndXL, e.EndYL, e.EndZL,
        e.IsLine, e.IsRapid, e.Length)     # StartZL/EndZL = 局部 Z（见 §4.8）
   ```

**为什么第 3 步不可省（本次的教训）：**
- 只看回执框计数 → 漏掉了两个真缺陷（无排版图纸整体失败、rapid 段被算进面积）；
- 只看 `Path.Length` → 会被 **rapid 定位段**带偏（见 §7.9），我因此误判了两次。

**测试台本身也要防坑：**
- 造件后先**读回确认形状**（`Length` / bbox / 元素数 / `Closed`），再开始测；
- 两件**别放共线紧邻**：第 2 条刀路会带一个**从前一件收尾点出发的 rapid 段**，
  它的 `MinXL~MaxXL` / `Length` 会把两件都框进去（本次 `bbox=[0,0,1005,300]`
  而 `GetFeedExtent=[405,0]-[1005,300]`）；
- 跑完把测试几何/刀路清掉（见 §7.2）。

**⚠️ 测试会往生产工程里**漏**临时模块（2026-09-13 实际发生）：**
本次往 `CCC功能` 工程里漏了 **2 个 `MCP_TEST_xxxx` 模块**（各 3 行）。原因：注入临时模块后
靠 `finally` 清理，但**工具调用被中断 → Python 进程被杀 → `finally` 没执行**。
它们会被一起写进 `.arb`，污染插件。**防法：**
- 注入器的**开头先清一遍**残留（`for c in VBComponents: if c.Name.upper().startswith("MCP_TEST")`）；
- **删之前先把名字读出来**：`Remove(comp)` 之后 `comp.Name` 会抛
  `系统错误 &H80070006 … 句柄无效`（本次第一版清理脚本就因此中断，只清掉一半）；
- 每次跑完用 `tools/running_snapshot.py <tag> <工程名>` **全量导出一次**，
  看组件数是否回到原值（`CCC功能` 正常是 **12 个**）。

### 5.6 静态"配平"检查先把**计数器自己的 bug**排除掉（2026-09-14，我的检查工具骗了我一次）

自己写的 `tmp/check_vba2.py` 用正则数结构配平，v2.1.2 改完报：

```
If / End If            65   / 65    OK
For / Next             21   / 20    *** 不配平 ***
```

**代码是好的，是计数器错了。** 三个错：

1. 数 `Next` 时把 **`On Error Resume Next`** 也算进去了；
2. 数 `For` 时把 **`Exit For`** 也算进去了；
3. **行中**的 `Next` 漏了（`For k = 1 To n: ord(k) = k: Next k` 这种单行写法，
   正则只匹配行首 `^Next`）。

修法：先去掉行尾注释 → 删掉 `On Error Resume Next` 与 `Exit For` →
再用 `\bFor\b` / `\bNext\b` **全文计数**（而不是只看行首）。改完 `For / Next 21 / 21 OK`。

**教训**：静态检查报"不配平 / 不一致"时，**先怀疑计数器**（尤其是自己刚写的），再去改代码 ——
否则会把好代码改坏。真值来源是 **VBA 编译器**（§5.3，以及部署探针里那次
`RampVersion()` 的"编译+运行"往返，见 `tools/ccc_probe.py`）。

---
## 6. 门板宏常见操作要点

### 6.1 路径分组：与门类型刀路关联时必须用**固定组号**

`GetNextGroupNumberForGeometries` 返回动态组号（取决于生成时绘图已有组），
**不稳定**——而门类型刀路（`AD_DOOR_PATHS.GroupID`）是固定值，动态组号会导致
刀路关联不上（刀路丢失/落到错误几何）。

```vba
Geo1.Group = 1   ' 固定组号，与 AD_DOOR_PATHS.GroupID 一一对应
Geo2.Group = 2
```

方案 A 经验：宏返回的每个几何设固定组 → 刀路按组落在正确几何上。

### 6.2 关于宽度中心轴镜像（`L0orR1 = 0` 时）

- 镜像 = 所有 X 坐标映射为 `width - X`。
- `FastGeometry.KnownArc` 的第二参数是 `CW`（True=顺时针），**镜像后必须取反**，
  否则圆弧凸侧反向。

```vba
.KnownArc R2, Not mirror, MirrorX(W + R2, width, mirror), length - H - R2
.KnownArc R1, mirror,     MirrorX(width - R1, width, mirror), length - H + R1
```

- `Path.Group` 属性可读写；镜像不改变路径 bbox 与周长，可用作验证特征。

### 6.3 `Path.Offset` 的 Left/Right 是**相对行进方向**，不是绝对的内/外（重要）

**现象：** 同一段代码 `Geo.Offset(B, -1)`（`acamRIGHT`）在 `L0R1=0` 时是内偏移，
`L0R1=1`（镜像梯形，路径方向反转）时变成**外偏移**，图形跑到外面。

**根因：** `Offset(Distance, Side)` 的 `acamLEFT(1)/acamRIGHT(-1)` 相对**路径行进方向**。
路径方向（CW/CCW）由顶点顺序决定，镜像图形方向会反转（用 Shoelace 公式可算有向面积验证）：
- CW 路径：Right = 内侧；
- CCW（镜像）路径：Right = 外侧。

**解决：** 偏移侧随路径方向切换：

```vba
If L0R1 = 0 Then OffsSide = -1 Else OffsSide = 1   ' 方向随 L0R1 镜像而反转
Set Offs = Geo1.Offset(B, OffsSide)
```

**验证特征：** 内偏移路径的 bbox 完全在原始路径 bbox 内，且周长更短。

---

## 7. 其他环境问题

### 7.1 Bash heredoc 里嵌大段 VBA/Python 会卡住

超长 heredoc（内含 Python 三重引号字符串）会导致 MCP 桥输入缓冲截断、工具卡死。
**解决：** 脚本先 `write_file` 落盘，再 `python script.py` 执行（本项目一直采用此方式）。

### 7.2 测试几何清理

验证宏创建的几何可用 `App.ActiveDrawing` 相关方法或 MCP `delete_all_geometries`
清理；COM 直接 `Run` 的宏创建的几何有时不持久化（无文档事务上下文），
运行后检查 `get_drawing_info` 的 `geo_count` 确认。

**🚨 2026-09-13 血泪警告：跑任何"清空图纸/造测试件"的脚本前，必须先打印并确认 `ActiveDrawing` 是谁。**

当天我用「清空 → 画 2 个矩形 → 造工序 → 跑 modRamp」的测试台反复验证，一直以为自己在一张
**空的未命名图纸**上（早期探针确实打印 `Name='' FullName=''`、0 几何 0 刀路）。但操作到后面
再打印时变成了：

```
Name     = '9-9测试自动化'
FullName = 'D:\2016\NC\9-9测试自动化\9-9测试自动化.ard'     ← 真实路径、被 AlphaCAM 独占打开
Modified = True
```

也就是说：**`ActiveDrawing` 会在会话期间变化**（同机多实例 + `GetActiveObject` 返回的实例
不稳定，见 §1.2），而我的测试台第一步就是 `SetGeosSelected True` + `DeleteSelected`
（**删光所有几何与刀路**）。一旦它落到真实作业图上，就是**直接删数据**。

**硬性规矩（写测试脚本时照抄）：**
1. **每次**清空前都读 `ActiveDrawing` 的 `Name` / `FullName` 并**打印/断言**；
   不满足预期（例如"必须是空图或我指定的测试文件"）**立刻抛出退出，不要清理**。
2. 不要用"上一次探针显示它是空的"作为依据 —— 实例/图纸随时可能变。
3. 需要一张干净的测试图时，**显式新建**（`App.New`）并记下它的 `Name`，
   之后每一步都用这个 `Name` 校验；或让用户准备一张可弃的测试图。
4. 清理范围宁可小（只删自己刚造的几何），也不要"全选删除"。

> **本次结局（2026-09-13 用户确认）**：那只图 `9-9测试自动化.ard` **本来就是测试用图**，
> 虚惊一场、无数据损失。但上面的规矩照样成立 —— 当时我拿"更早的探针显示它是空图"当依据，
> 而依据必须是**当场读到的** `Name`/`FullName`。

---

### 7.3 窗体存在损坏控件（读属性报"无效参数"）

**现象：** 遍历 `Designer.Controls` 时，某个控件读 `Name` 即报：
```
(-2147352567, '发生意外。', (0, 'Forms.Form.1', '无效参数。', 'fm20.hlp', 0, -2147024809), None)
```
但该控件**不影响工程编译与运行**（VBA 宏列表正常、宏可执行）。

**原因：** 手动创建窗体时残留的损坏/无类型控件（可能是 OLE 控件库丢失或设计器异常对象）。

**处理：** 可忽略；如需清理，用 `Controls.Remove` 按序号删除，删除前确认不是被代码引用的控件（代码只按 `Name` 引用，未知控件未被引用可安全删除）。

---
### 7.4 CDM.arb 损坏：AlphaCAM 启动报"取得选项ID失败 / 无法打开CDM / Error loading CDM Processing"

**现象：** AlphaCAM 启动时弹窗：
```
无法打开CDM。R1\StartUp\CDM\CDM.arb
取得选项ID失败:C:\Program Files (x86)\Vero Software\Alphacam 2016 ...
Error loading CDM Processing.
```
CDM 菜单/功能全部缺失（`list_vba_modules` 只剩 57 个组件，`modAutoImportNest`/`frmAutoNest` 消失）。

**根因：** `CDM.arb`（OLE 复合文档，含 VBA 工程源码+窗体+`Licom/OptionID` 配置）中 **`Licom/OptionID` 流丢失**。
触发链：**AlphaCAM 退出/保存时把内存 VBA 工程持久化写回 `CDM.arb`**（文件大小从 4.3MB → 5.2MB，流数 264 → 333）——
若此时进程崩溃（如 RPC 断开/宏执行中删模块），OLE 结构写入不完整 → OptionID 流丢失 → 下次启动加载失败。

**诊断方法（不依赖 AlphaCAM）：**
```python
import olefile
ole = olefile.OleFileIO(r'...\StartUp\CDM\CDM.arb')
ole.exists('Licom/OptionID')          # False = 损坏（关键流）
len(ole.listdir())                     # 与备份对比（正常 264，损坏 333）
```
- 关键流：`Licom/OptionID`（8 字节）、`Licom/AlphaCAM`、`vao/The VBA Project/...`（VBA 工程数据）
- `CDM.err` 里的 `CDM.ctx` 错误（"NOT ENOUGH LINES FOR $600"）是**旧的非致命问题**，勿混淆

**恢复流程（已验证）：**
1. 备份损坏文件：`copy CDM.arb backup/CDM.arb_<日期>_broken.bak`
2. 从完整备份恢复：`copy backup/CDM.arb.bak CDM.arb`（必须 AlphaCAM **完全关闭**，注意写权限）
3. 重启 AlphaCAM → CDM 正常加载
4. **重装丢失的代码**（恢复版本不含近期改动）：
   - `install_vba_module('modAutoImportNest', 本地bas)` 
   - `install_vba_module('Events', 本地bas)`（含菜单注册）
   - **重建 UserForm**：`VBComponents.Add(3)` → Name → `Designer.Controls.Add` 10 个控件 → `CodeModule.AddFromString(本地txt)`
     （控件清单/坐标见 `CDM功能/frmAutoNest_手动创建.md`；v1.10 起多了 `lblMaterial` + `cboMaterial`。
     顺序铁律：**先加控件、再灌引用它们的代码**，反了会编译报"找不到方法或数据成员"）

**预防（重要，2026-08-15 实测补充）：**
- ⚠️ **此环境 AlphaCAM 保存 CDM.arb 会反复丢失 `Licom/OptionID` 流**（8/14、8/15 已发生两次）：
  反复 `install_vba_module` 触发自动保存后，某次保存 OptionID 就丢（触发条件不明确，正常保存也可能丢）
- **恢复后立即备份**：`python backup_cdm_arb.py`（校验 OptionID，缺失会明确提示不可用）
- **每次装完模块后检查 OptionID**：olefile 读 `Licom/OptionID`，丢了马上从最近可用备份恢复
- 建议工作流：改动代码 → 重装 → **立即备份** → 继续；崩溃/RPC 断开后必查 CDM.arb
- **定期备份 `StartUp\CDM\CDM.arb`**（崩溃后必查）
- 崩溃/RPC 断开后先检查 CDM.arb 大小与 `Licom/OptionID` 流，再决定是否恢复
- 宏执行中/对话框残留时避免 `install_vba_module`（删模块操作易触发保存崩溃）

**2026-09-10 实测补充：三组反直觉证据 + 退出保存本身可修复**

① **mtime 冻结 ≠ 没被写过。** 损坏期间 `CDM.arb` 被 AlphaCAM 以**写方式独占持有**：
```powershell
Get-FileHash CDM.arb   # 报 "being used by another process"，根本读不了
Get-Item CDM.arb       # LastWriteTime 停在最后一次 flush
```
该文件随后**内容确实变了**（`OptionID` 消失），但 mtime 一直没动。
→ **不能用 mtime 判断"是否被改过 / 是否已保存"**，这是本次最容易误判的一点。

> **2026-09-13 又踩了一次同样的坑（补充判据）：**
> 给 `CCC功能.arb` 做备份时，它 mtime 停在 **2026-07-29**，我据此判断"今天改的 `modRamp` 还没落盘" ——
> **错了**。用 olefile 搜内容才发现该 .arb 里**已有 v2.0.2 的代码**（`RampVersion` / `BuildTabWindows`
> / `v2.0.2` 均命中），说明 AlphaCAM **运行期间就在写它**。
>
> **正确判据（正反两面都要查，不能只搜一个字符串）：**
> ```python
> # 正面: 新代码独有语句应命中
> "BuildTabWindows".encode("gbk") in stream
> # 反面: 旧代码独有语句应【0 命中】
> "朝排版中心方向偏移整条边长".encode("gbk") in stream
> ```
> 本次只查正面会误判 —— 因为 **v1.x 的 `bestDist` 字符串在 .arb 里仍有 1 个流命中**（VBA 的
> 源码保留流里留了**陈旧碎片**）。只有"旧语句 0 命中 + 新语句命中"才能定性。
> 另：`VBProject.Saved` 在宿主型 VBA 里**恒为 True**，同样不能用来判断是否已落盘。

② **运行中的 CDM 完全可用，不代表文件完好。** `OptionID` 缺失期间，用户仍完整跑通了
「自动化生产排版 + 重新生成标签」（6 个标签 EMF 正常产出）。**内存里的工程是好的，
危险只在下一次启动**。→ "功能正常所以 .arb 没事"这个推断是错的。

③ **备份脚本也会产出不可用副本。** 源文件被锁时 `shutil.copy2` 可能拷出**缺流的副本**：
```
已备份 -> backup/CDM.arb_20260910_134926.bak
大小: 5013504 字节 | 流数量: 269 | Licom/OptionID: [缺失 - 该备份不可用!]
```
→ `backup_cdm_arb.py` 的 `Licom/OptionID` 判定是权威判据；**报"缺失"的备份必须立刻重命名标记**
（如 `..._corrupt_noOptionID.bak`），否则以后会被误当成回退点。
另注意：源文件被锁时，直接 `olefile` 读**活动** `.arb` 反而能读（OLE 结构可共享读），
所以"能读出结构"和"备份可用"是两件事，以备份报告的判定为准。

**计数口径（两条容易混，今天都量过）：**

| 方法 | 正常 | 损坏 |
|---|---|---|
| `backup_cdm_arb.py` 报告的「流数量」 | **270** | 269 |
| `ole.listdir(streams=True, storages=True)` | **333** | 332 |

关键是**两者都只差 1**，那一个就是 `Licom/OptionID`。
（本文档早前记的 "正常 264 / 损坏 333" 是更早、组件更少时的口径，别直接拿来比。）

**新恢复手段：优雅关闭本身就能修复（今天实测有效），不一定需要从备份还原。**
AlphaCAM 的**退出保存会把内存工程完整重写**，`OptionID` 随之回来：
1. 先备份现状（哪怕是坏的，留证）：`python backup_cdm_arb.py`
2. **优雅关闭** AlphaCAM —— `PostMessage WM_CLOSE` 到 `AlphaCAM_3DMILL` 主窗口。
   本次**没有任何对话框**，0.5 秒干净退出（关闭前 `ActiveDrawing.Modified = False`，
   确认过没有未保存图纸）
3. 校验：`ole.exists('Licom/OptionID')` → True，流数量恢复 270 / 333
4. 立即备份 → 得到新的已知可用版
5. **重启验证**：无"取得选项ID失败"弹窗，且运行中模块版本正确

**如何判断代码是否真的落盘（不需要解压 MS-OVBA）：**
VBA 模块源码以 MS-OVBA 压缩存储在 `vao/The VBA Project/_VBA_Project/VBA/<模块名>`。
压缩流中**字面量的首次出现是原样存储**的，所以可直接在原始字节里搜**版本独有的 ASCII 标识符**：
```python
data = ole.openstream('vao/The VBA Project/_VBA_Project/VBA/modAutoImportNest').read()
b'ScreenUpdating' in data      # 该版本独有的标识符命中 → 证明这一版已落盘
```
⚠️ 两个坑：**命中可信、未命中不可信**（长字符串会被 copy token 打断）；
且**绝不能用流大小判断内容** —— 同一工程不同保存方式会让流大小剧烈波动
（本次未改动的 `frmAutoNest` 也涨了 202 字节，`Make` 源码只加约 400 字节却涨了 36,677 字节）。

---

### 7.5 加工道次窗口需调整视图后才能操作（AlphaCAM 固有现象）

**现象：** 视图/屏幕状态变化后（缩放、隐藏/显示路径、宏操作视图），
加工道次（Operations）窗口**暂时无法操作**；**手动调整一下视图（缩放/重绘）后才恢复正常**。

**结论：** 这是 **AlphaCAM 固有行为**，不是代码 bug——视图状态变化后窗口需重新绑定。
- 宏操作视图后，末尾补 `ActiveDrawing.ZoomAll`（或 `Redraw`）自动触发刷新
- 排查时勿误判为宏引入的问题

**相关教训（"重新生成标签"功能）：**
- **不要在用户正在编辑的图上调用 `m_CreateAlphaCAMDrawingsOfSheets`**——它内部
  `SaveAs`（覆盖文件）+ `MoveToDrawing`（搬走路径）+ `OpenDrawing`（重开图纸）
  会破坏加工道次/刀路关联（用户实测"道次乱、刀路对不上门板"）
- **只读生成方案（已验证）**：遍历嵌套件，`Visible=False` 临时隐藏其他件 →
  `ZoomToBox` 到当前件 → `SaveEmfFile` → 恢复可见与视图。
  **全程不 `SaveAs`/`MoveToDrawing`/`OpenDrawing`，用户图零修改**，
  加工道次不受影响（仅固有"调整视图后可用"现象）

---

### 7.6 屏幕刷新被宏关掉且未恢复 → 窗口标题/画面"卡"在旧的临时档案名上

**现象：** 跑完"自动化生产排版"或"重新生成标签"后，AlphaCAM 主窗口标题一直是
`3D 5-轴鉋花机专业版: regen_<订单>_<Timer>`（临时副本名），画面也不更新，
**看起来像打开的还是临时档案**。

**真相：** 文档其实早就换回真档案了。用 COM 读一下就知道：

```python
app = win32com.client.GetActiveObject("aroutaps.Application")
d = app.ActiveDrawing
d.FullName          # D:\2016\NC\<订单>\<订单>.ard   ← 真档案
d.Modified          # False
d.ScreenUpdating    # False   ← 病根
```

`ScreenUpdating = False` 时 AlphaCAM **不重绘任何东西**，所以标题栏与画面都停在
最后一次重绘时的状态，与真实的活动文档无关。

**根因：** `Make.m_CreateAlphaCAMDrawingsOfSheets` 开头置
`ActiveDrawing.ScreenUpdating = False`、`Frame.ProjectBarUpdating = False`，
但收尾处的两行恢复语句**被注释掉了**（`Make.bas` 的 3991/3992；同段的
`App.DisableUndo = False`、`QuickShading = blnQuickShade` 都有恢复，唯独这两项漏了）。

**修复：** 取消那两行注释，并在调用方兜底：

```vba
ActiveDrawing.ScreenUpdating = True
Frame.ProjectBarUpdating = True
ActiveDrawing.Redraw
```

**现场急救（不必重启 AlphaCAM）：** 连 COM 设 `ActiveDrawing.ScreenUpdating = True`
→ `Redraw()` → `ZoomAll()`，标题栏立刻恢复正常。

**教训：** 判断"当前打开的是哪个档案"**不要看标题栏**，要读 `ActiveDrawing.FullName`；
标题栏可能因屏幕被锁定而严重滞后。

---

### 7.7 VBA 边枚举边增删 → 同一件号多个几何 → 标签"同一位置重复涂黑"

**现象：** 多个零件位置调整过若干次后，重新生成门板标签，会出现**两张标签涂黑在同一位置**
（例：订单 `9-11纳百川`，`_2` 与 `_3` 的 hatch 位置重合）。

**取证手段（可复用）：** 标签是 EMF 矢量图，用 .NET 渲染成 PNG 就能直接用眼睛比对：
```powershell
Add-Type -AssemblyName System.Drawing
$img = [System.Drawing.Image]::FromFile($emf); $bmp = New-Object System.Drawing.Bitmap($img.Width,$img.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp); $g.DrawImage($img,0,0,$img.Width,$img.Height)
$bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png)
```
把多张标签的**同一区域**裁出来纵向拼一张，就能看出"被涂黑的是不是同一条"。

**根因（`Make.bas` 的 `m_ExportDoorLabelEMFs`，三处缺陷）：**
1. **边枚举边删**：去重循环在 `For Each ActiveDrawing.Geometries` **内部**调 `SheetPath.Delete`。
   VBA 的 `For Each` 在枚举中被改动会跳项 → 同一 `DEF_ATT_NEST_DOOR_COUNT`
   （`LicomUKljo_alphadoor_nest_door_count`）**残留多个几何**。
2. **只删最后一个 hatch**：`Set ps = App.ActiveDrawing.HatchPath(...)` 会**覆盖引用**，
   一轮里匹配到多个几何时前面的 hatch 无人删除 → **残留到下一张标签的同一位置**。
   （`Drawing.HatchPath` 返回 `Paths`，是**真的往图纸里加了几何**，必须显式删除。）
3. **`ps` 每轮不重置**：某件号没有匹配几何时，会对**已删除**对象重复 `Delete`。

**改法：** 几何先**快照进 `Collection` 再增删**（绝不在 `For Each` 里 Delete/Add）；
每轮产生的 hatch **全部收集**，在 `SaveEmfFile` 之后**逐个删光**。

**通用教训：**
- VBA 里**任何** `For Each` 循环体内都不要 `Delete`/`Add` 被枚举的集合 —— 先快照。
- 会往图纸里加几何的 API（`HatchPath` 等）**返回的是句柄**，用 `Set x = ...` 覆盖前
  先确认前一个是否已删；多个时要用 `Collection` 收全。
- 部署后若发现模块里 `.Add` 变成 `.add`（或其它成员名大小写变化），那是 **VBA 编译时的
  成员名归一化，不是错误** —— 比对运行中代码与仓库文件时**必须大小写不敏感**，
  否则会误判成"部署失败"（本次就误报过一次 `RESULT=FAIL`）。

**注意：** 同一订单明细数量 >1 时，排版上会有多个实例共用同一个**印在板上的件号**，
因此那几张标签"看起来一样"是正常的；本条的 bug 特指**同一位置残留重复涂黑**。

---

### 7.8 `Frame.ReadTextFile` 读**不存在的行号** → 模态报错框 + 刷屏记事本（2026-09-11 我踩的）

想确认"自动化生产排版"挂在哪个菜单下，为了"多查几项"写了个探针，对
`CDM.ctx` 连读了 `(1,1) (3,1) (7,1) (8,1) (12,1) (13,1) (14,1) (15,1) (16,1)` —— 其中
**7/8/12/13/14/16 六行根本不存在**。结果：

- AlphaCAM 弹出**模态框**并卡住界面：
  `ERROR IN READING TEXT FILE ...\CDM\CDM.ctx` /
  `ERRORS WILL BE WRITTEN TO ...\CDM\CDM.err` / `确定`
- 同时**连开 6 个记事本**显示 `CDM.err`（一个错误一个），全是垃圾窗口。
- `CDM.err` 内容（766 字节，正好 6 条）：
  `TEXT MESSAGE(S) NOT FOUND :-` + `$7 / $8 / $12 / $13 / $14 / $16`

**关键：`CDM.ctx` 文件本身完全正常**（26 KB，2015/11/23），路径也没错 ——
`ReadTextFile` 是按 **(行, 列)** 取**文本资源槽**，槽位不存在就报错，
报错信息却写成"读文件失败"，极容易被误判成文件损坏。

**教训（务必遵守）：**
- **绝不盲读 CTX 行号**。要用哪条资源，先在**已知存在的行**里确认，或改成一次读整块后解析。
- CTX 是**稀疏**的：本次实测真正有内容的只有 `(1,1)="CDM橱柜"`、`(3,1)="CDM"`，
  连 `(15,1)="确定(&O)"` 这种看着该有的槽也是碰运气。
- 读 CTX 前**先用 `On Error` 包住**，并且**不要**把"读不到"当成"文件坏了"。
- 报错框是**模态**的、会卡住 AlphaCAM 的界面；但**不影响 COM**（`run_vba_line`/`App.Run`
  期间仍可执行），所以"探针返回 ok"**不能**证明没有弹窗 —— 本次就是探针 ok 但屏幕上有框。
- AlphaCAM 自己的错误处理会 `ShellExecute` 打开 `CDM.err` 记事本，**同一错误重复触发会叠一堆窗口**，
  事后要清理（本次 6 个）。

**清理手法**（无 pywin32 依赖，ctypes 即可）：枚举顶层窗口 → 按标题匹配
→ `PostMessageW(hwnd, WM_CLOSE, 0, 0)`。见 `tools/win_ctl.py --close-title "CDM.err"`。

**顺带一个环境事实：** AlphaCAM 的进程名是 **`Acam.exe`**（PID 会变，主窗口类名
`AlphaCAM_3DMILL`，标题 `3D 5-轴鉋花机专业版`）—— 用 `Get-Process | ? ProcessName -match 'alphacam'`
**过滤不到它**，别据此判断"AlphaCAM 没在跑"。

### 7.9 `Path` / `Element` 的几个数据陷阱（2026-09-13，全部实测）

#### (a) `MinXL/MaxXL/Length` **含 rapid 段**；`GetFeedExtent` 才是不含的

第 2 条及以后的刀路，**开头会带一个从"上一件收尾点"出发的 rapid 定位段**。实测：

```
path2: MinXL~MaxXL = [0,0]-[1005,300]      ← 把前一件也框进来了
       GetFeedExtent = [405,0]-[1005,300]  ← 真实加工范围
       Length        = 2205 = 405(rapid) + 1800(真实周长)
```

**规避：** 判断"件多大/在哪/多长"一律用 `GetFeedExtent`（取不到时再回落包围盒），
或**逐元素累加、跳过 `IsRapid`**。
**本次踩坑：** 用 `MinXL~MaxXL` 算件面积 → 第 2 件被算成 301500mm²（实为 180000）→ 该降速的没降速。
**复制刀路几何时也必须跳过 rapid**（`小条先切` 与 `modRamp` 的 `If Not elem.IsRapid Then` 就是这个原因），
否则复制出的几何会多出一条横穿的直线。

#### (b) `Path.SetStartPoint(X, Y)` 只对**闭合路径**有效，且点必须在路径上

文档原文只有一句：*"Set the start point for the path, **if path is closed**"*，
参数是"新起点的局部坐标" —— **没有任何"自动吸附到路径上"的承诺**。
官方示例传的也都是路径上的点（`Geo.SetStartPoint 50, 100` 是 0..100 矩形的上边中点）。

⇒ 传一个落在路径外的点属**未定义行为**。本次实测到一个真实反例：
`modRamp` v1.x 用"沿边整边跳一次"求点，100×600 的竖条会算出 `(50, -300)`
（在包围盒外，靠 `SetStartPoint` 自身兜住才没出事）。v2.0 已改为**直接算目标边的中点**。

#### (c) `Drawing.GetNestInformation` 在**无排版的图纸上直接抛异常**

报错是 `现在图档内无排版`。**单件图、测试图都属于这种情形**，不是错误。
必须按 §4.7 的模板包起来：

```vba
Set Ni = Nothing
On Error Resume Next
Set Ni = ActiveDrawing.GetNestInformation   ' 无排版会失败，吞掉
On Error GoTo EH                            ' ← 勿用 On Error GoTo 0
If Ni Is Nothing Then ' 退回"图纸范围中心"等兜底
```

**本次踩坑：** `modRamp` v1.x 没包 → **在单件图上整个功能失败**（弹红框），
而它自己的 `SetGeoStartToSheetSide` 里明明写了兜底分支，永远走不到。v2.0.1 修。

#### (d) 手工重建刀路时的正确写法（文档实证）

- `MillData.ManualToolPath(X, Y, Z)` 给**起点**，`MillManualToolPath.Add3DLine(X, Y, Z)`
  是**进给直线**（起点=上一段终点，Z=**终点** Z）。
- **从 `Z = 0`（板面）起步再沿路径下降**，不要把 `ManualToolPath` 的 Z 直接设成切深 ——
  那等于在起点**直插**一整刀。
- 逐段 Z 用 `Add3DLine` 表达即可（3D 进给直线，天然形成斜坡）。

#### (e) 常用但反编译文件里看不到的成员（用 §4.8 的 `dir()` 查）

`Element.Length`、`StartZL/EndZL`、`StartZG/EndZG`、`LeadIn/LeadOut`；
`App.LicomdatPath`（拼刀具路径，别硬编码盘符）、`App.CreateTool`、`App.CreateLeadData(3D)`；
`Drawing.CreateRectangle(x1,y1,x2,y2)`、`Path.SetMaterial(厚, Z)`。

#### (f) `MillData.ManualToolPath` 重建后，**加工方式（子工序）名会变成「手动输入」**（2026-09-13 实测）

实测：测试图里原本是 `精加工   刀具 2   FLAT - 5MM` 的子工序，其刀路被 `modRamp`
用 `ManualToolPath` 重建之后，工程里能看到的加工方式变成了
**`手动输入   刀具 2   FLAT - 5MM`** —— **原来的「精加工」不再出现**
（未确证是被改名还是另建了子工序，但现象明确）。

**后果（凡是按"加工方式名"筛刀路的自动化都会受影响）：**
- `modRamp` 自己的 `methodName` 参数，**第二次跑就匹配不上了** —— 得选「手动输入」才命中。
- 一般化教训：**`ManualToolPath` 不是"原地替换"**。它会把刀路落到一个「手动输入」子工序里，
  所以要动刀路结构的插件，都要重新确认它对**子工序名/工序结构**的影响。

**2026-09-13 v2.1.0 的修复情况（实测确认）：**

| 项 | 状态 |
|---|---|
| **加工道次 `OpNo`** | ✅ **已回填**：`Finish` 返回新建的 `Paths` → 逐条 `OpNo = 原道次` → 收尾 `Operations.OrderAll`。实测基线 `opNo=1` → 处理后仍是 `1` |
| **幂等标记 `CCC_RampDone`** | ✅ **已修复**：原先写在被删除的旧刀路上（丢失）→ 现在写在 `Finish` 返回的**新**刀路上。实测第二次执行报「候选=2, 已处理过=2」，刀路元素数 `[210,211]` **未变** |
| **子工序名（加工方式）** | ❌ **仍会变成「手动输入」**（实测 `精加工` → `手动输入`）—— 道次保住了，但名字没保住。按"加工方式"筛选的流程仍要按「手动输入」去找

**2026-09-14 补证：不是"改名"，而是【为每条新刀路各建了一个子工序】，且名字【改不回来】。**

- 加工道次窗口实测：处理前 `Op 1` 下 **1 个子项**（`精加工 刀具 2 FLAT - 5MM`）；
  `modRamp` 处理 4 个件之后变成 **4 个子项**，每个都叫 `手动输入 刀具 2 FLAT - 5MM`。
  所以 `Finish` **不是原地替换**，是**另建子工序**（原先那句"未确证"到此确证）。
- `SubOperation.Name` 是 **只读** —— `tempacamapi\Objects\SubOperation\SubOperationProperties.htm`：
  > `Name - (String) The name of this sub-operation, as in the Operation List dialog box (read-only)`

  所以**事后**无法把「手动输入」改回「精加工」。
- **但名字可以在【建之前】定好** —— ✅ **2026-09-14 v2.2.0 已解决**：
  建手工刀路前给 MillData 设带参属性
  `md.Attribute("LicomUKDMBOperationName") = "精加工"`，新子工序就叫
  `精加工   刀具 2   FLAT - 5MM`。详见 §7.9(i)。
- 影响可控：`OpNo` 已回填（见上表），刀路顺序由 `Operations.OrderAll` + 插件自己的排序决定，
  NC 输出仍按道次 `1` 分组。
- 剩下的结构性差异：**一个子工序变成多个**（每件一个，名字相同）——
  这是 `ManualToolPath` 的固有行为，改不掉；若一定要保住子工序结构，唯一方向是
  **别用 `ManualToolPath` 重建**：改成 AlphaCAM **原生斜坡参数**（`MillData.AutoZ` /
  `AutoZRampAngle`，见 `开料小板件防松动算法方案.md` §11）—— 那是"改参数重算刀路"，工序结构天然保留。

#### (g) `MillManualToolPath.Finish` 会把**收尾的抬刀 rapid 单列成一条 1 元素刀路**（2026-09-14 实测）

`modRamp` v2.1.1 在轮廓切完后补一条 `Add3DRapid(起点X, 起点Y, 20)`，
`Finish` 返回的是**两条** `Path`：

| Path | 元素数 | 内容 |
|---|---|---|
| `ToolPaths(1)` | 210 / 211 | 斜坡 + 轮廓（末元素 `EndZL = -18`）|
| `ToolPaths(2)` | **1** | `IsRapid=True`，`Z -18 → 20`，XY 与上一条的末点相同 |

**后果（验收脚本的坑）**：我第一版验收只检查"每条刀路末元素是不是 Z20 的 rapid"，
于是把**轮廓刀路**（末元素 Z=-18）判成"抬刀失败"，其实功能是好的 ——
**差点去修一个不存在的 bug**。

**正确判据**：把**同一子工序内所有刀路按顺序拼接**，看整体末尾元素；
或断言"子工序内恰有 1 条 Z20 抬刀 rapid，且它是最后一条刀路的唯一元素"。
设备视角两者完全等价（抬刀本来就是一个独立的快速移动段）。
**顺带**：跨子工序交接时，下一条子工序的**首元素**是一条 `Z 20 → 0` 的 rapid，
可用来交叉验证"上一件的抬刀真的生效了"。

**顺带 2**：`GetFeedExtent` 对这种 1 元素 rapid 刀路返回 `False`（没有进给段），
`MinXL==MaxXL`；算包围盒/中心时必须跳过这种空进给刀路（`If Not blnExt`）。

#### (h) `Drawing.OrderManual(Ordered)` 能重排加工顺序，但会**重派生刀路之间的连接 rapid**（2026-09-14 实测）

文档原文：

> `drw.OrderManual (Ordered)` —— The parameter Ordered contains some or all of the geometries or
> tool paths (not both) in the drawing. **The paths will be ordered to match the order of the paths
> in this collection.**

`Ordered` 用 `drw.CreatePathCollection()` 建，再逐个 `.Add` 路径。

**实测结论（4 件测试台，逐元素"指纹"对比）**：

- **刀路本体不会被改写**：斜坡段元素数、逐段 Z 剖面、斜坡锚定关系**逐元素不变**；
- 被重新派生的只有**刀路之间的连接 rapid**：
  - 被排到**首位**的那条会**去掉**引导 rapid（212 → 211 元素）——
    这正是 AlphaCAM 对"首条刀路"的固有约定（图纸里第一件本来就没有引导 rapid）；
  - 排到后面的每条会**补上**一条从上一件终点出发的引导 rapid（4 → 5 元素）；
- **自定义属性会存活**：`CCC_RampDone` 仍在 → 幂等标记不会被重排弄丢
  （实测重排后第二次执行报 `候选=6, 已处理过=6`，结构不变）；
- **子工序按新的刀路顺序重新分组**（`sub1` 变成新首位那一组），子工序名不变。

**坑（我第一次实验就踩了）**：目标顺序如果把"轮廓刀路"和它配对的"抬刀 rapid 刀路"**拆开**
（我第一版把抬刀排到了最后），AlphaCAM 为了自洽会去补/挪 rapid，结果 element 分布变得无法解释
（轮廓少 1 个元素、抬刀涨成 2 个元素、多出一条 `Z=0` 的横穿 rapid），
**看起来像"刀路被改坏了"**。

> **规矩**：重排时**必须保持"轮廓 + 其配对的收尾 rapid"成组不拆**；
> 并且只在**顺序确实需要变化**时才调用（`dst(k) Is drw.ToolPaths(k)` 逐条比较对象同一性），
> 不要无条件调用 —— 无谓的调用会白白改写未匹配刀路的引导 rapid。

#### (i) 想让手工刀路不叫"手动输入"：`MillData` 的带参属性 `LicomUKDMBOperationName`（2026-09-14 实测）

文档（`MillData.ManualToolPath` 页脚）：

> If the Attribute with name **"LicomUKDMBOperationName"** is set for the MillData object
> it will be used as **the name of the operation** in the operation list and the NC code.

实测（v2.2.0 落地）：

```vba
Set md = App.CreateMillData
md.FinalDepth = -18
md.Attribute("LicomUKDMBOperationName") = "精加工"     ' ← 必须用带参属性写法
Set mtp = md.ManualToolPath(sx, sy, 0#)
```

结果：新建子工序名 = **`精加工   刀具 2   FLAT - 5MM`**（`刀具 …` 部分由 AlphaCAM 自动附加）
—— 于是 §7.9(f) 里"名字丢掉、只能叫手动输入"的老问题**从根上解决**，
加工道次窗口与 NC 注释都回到原样。

**写法务必注意**（详见 §7.13）：`md.SetAttribute "名", 值` 在 VBA 里是**编译错误**。

### 7.10 加工道次窗口（ProjectBar）不刷新 → 用 `Frame.ProjectBarUpdating`（2026-09-14 实测）

**现象**（用户反馈）：插件跑完，**加工道次窗口还是旧的**（`Op 1` 下仍显示 1 个子项），
要点一下窗口 / 重新选中才能看到新建的刀路。用户原话："op刷新失败加工道次窗口没有更新"。

**根因**：`modRamp` 只恢复了 `Drawing.ScreenUpdating`（§7.6），**没管 ProjectBar**。
`Frame.ProjectBar` 就是**加工道次窗口**。

**ACAMAPI 原文**（`Frame.ProjectBarUpdating`）：

> Set to `False` to stop the project bar being updated ...
> Set to `True` when the macro has finished adding paths, the project bar will then be updated.

**修法**：用 `False` / `True` 把**整个修改过程**包起来，**错误分支里也必须恢复**
（先存 `Err` 再动别的，见 §4.4）：

```vba
Frame.ProjectBarUpdating = False            ' 处理前
On Error GoTo ErrHandler
...  增删刀路 ...
drw.Operations.OrderAll
Frame.ProjectBarUpdating = True             ' ← 这一句才触发窗口刷新
drw.ScreenUpdating = True
Exit Sub
ErrHandler:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description    ' 先存 Err, 见 §4.4
    Frame.ProjectBarUpdating = True          ' 出错也要恢复, 否则窗口永久冻结
    drw.ScreenUpdating = True
```

**验收**：光读 `Frame.ProjectBarUpdating = True` 只是**间接证据**（那是我们自己写的值）。
**直接证据是看窗口**：处理前 `Op 1` 下 1 行 `精加工…`，处理后 **4 行** `手动输入…`（§7.9(f) 的截图核对）。
取图用 `see` 抓 AlphaCAM **主窗口句柄**即可（`list_windows` 里标题如 `3D 5-轴鉋花机专业版`），
**不需要切前台**，也就不会像 §1.2 那样把按键送到别的程序里去。

### 7.11 VBA 工程处于 `[中断]`(break) 状态时，VBE 对象模型给的是**陈旧且误导**的结果（2026-09-14 实测，重要）

**现象**：注入的临时模块有编译错误 → 弹"编译错误: 方法和数据成员未找到" → 关掉弹窗后：

- `VBComponents.Remove(mod)` **返回成功，但组件仍在**，`VBComponents.Count` 也不变；
- 遍历 `VBProjects(i).Name` 会报"**该工程已被保护，不能执行操作**"——
  而 `ActiveVBProject.Protection = 0`，**根本没被保护**（这条假消息会把你引向完全错误的方向）。

**根因**：此时 VBA IDE 标题是 `... - CCC功能 [中断] - [modRamp (代码)]`，
`VBE.ActiveVBProject.Mode = 1`（`vbext_vm_Break`）。**中断态下 VBE 对象模型的状态是陈旧的**，
结构修改可能被静默忽略（本次实际是"改动生效了但列表没刷新"，两种表现都见过）。

**纯 COM 复位（不用键盘，避免 §1.2 把按键送到别的窗口）**：

```python
vbe = app.VBE
for i in range(1, vbe.CommandBars.Count + 1):
    bar = vbe.CommandBars.Item(i)
    for j in range(1, bar.Controls.Count + 1):
        c = bar.Controls.Item(j)
        if int(c.ID) == 228:        # 标准/调试工具栏的「重新设置(&R)」= Reset
            c.Execute()
```

复位后 `Mode` 变成 2（设计态），组件列表立刻正确。可复用脚本：`tmp/reset_vba.py`。

**操作规程（现在是硬要求 —— 跑完任何注入式 VBA 测试都要做）**：

1. 复核 `CCC功能` 组件数回到 **12**（正常值）；
2. 发现 `MCP_*` 残留 → **先 Reset 再删**（中断态删不掉）。

### 7.12 跨线程复用同一个 COM 对象 → 报 `<unknown>.Name` 这种没头没脑的错（2026-09-14）

写"带超时的 `app.Run`"时，我在主线程 `GetActiveObject` 拿到 `app`，却在 worker 线程里调用它；
结果 `app.Run(...)` 抛异常，`str(e)` 就是 **`<unknown>.Name`** —— 既不是 `Run` 的问题，
也不是 `Name` 的问题。把 `GetActiveObject` / `AddFromString` / `Run` / `Remove`
**全部挪进 worker 线程**（各自 `pythoncom.CoInitialize()`）后立刻正常。

**判据**：报错文本里出现 `<unknown>.xxx`，**先怀疑跨线程/跨套间使用 COM 对象**，
不要顺着那个 `.xxx` 去查成员。可复用脚本：`tmp/vba_try.py`
（自带超时 + 弹窗清场 + 中断态复位 + 清残留，用来安全试跑 VBA 片段）。

### 7.13 带参属性(parameterized property)在 VBA 里只能写 `对象.属性("名") = 值`（2026-09-14）

想设 `MillData` 的 `"LicomUKDMBOperationName"` 时：

| 写法 | 结果 |
|---|---|
| `md.SetAttribute "LicomUKDMBOperationName", "精加工"` | ❌ VBA **编译错误**：方法和数据成员未找到 |
| `md.Attribute("LicomUKDMBOperationName") = "精加工"` | ✅ 正确 |

原因：typelib 里它是 **id 1003 的 propget + propput**（同一个带参属性），pywin32 生成包装时
把 put 侧单独命名成 `SetAttribute`（所以 `dir()` 里能看到这个名字），
**但 VBA 编译器不认这个拼出来的名字**，只认属性语法。
`Path.Attribute` 一直是这个写法（`tp.Attribute("CCC_RampDone") = 1` ✓），
只是 `MillData` 上第一次用。

**怎么快速定位这类问题**：`dir()` 只说明"有这么个名字"，**不代表 VBA 认**。
权威依据是 typelib 里的 `_prop_map_get_` / `_prop_map_put_`：
`%LOCALAPPDATA%\Temp\gen_py\3.12\<typelib>\Xxx.py`。看到
"# The method X is actually a property, but must be used as a method to correctly pass the
arguments"，就说明 **VBA 侧要用属性写法**。

**附带一条**：编译错误弹窗会**阻塞调用方**（§4.2），所以试跑 VBA 一定要
**后台线程盯着弹窗并点掉**（`tmp/vba_try.py` 的 `DialogWatcher`），
否则脚本会僵在那里。

---

## 8. AlphaDOOR（CDM）门板机制与数据库（本项目核心）

### 8.1 门板构成

CDM 生成门板 = **外部几何**（宽高矩形，CDM 自动创建）+ **样式几何**（运行门样式宏
`AdoorMain` 返回的 `PathsToReturn`）+ **门类型刀路**（`AD_DOOR_PATHS` 记录，生成时
应用到对应几何）。

### 8.2 原轮廓矩形（外部几何）生成过程

`CDM` 项目 `Make` 模块 `mbln_Style_Make_930`（约 7692 行起）：

```vba
If Not .IgnoreOuterGeometry Then
    '..create the door perimeter
    Set pthOut = ActiveDrawing.CreateRectangle(0, 0, RequiredData.Width, RequiredData.Length)
    Set dll = CreateObject("StdAlpha.ShareClass")
    dll.jc pthOut, RequiredData.UserVariables(46), ...(49)   ' 用户变量46-49参与矩形造型
    dll.diamond pthOut, .CornerRadius                        ' 四角圆角
    pthOut.Group = 0                                         ' 矩形组号 = 0
    lngGeoNumber = lngGeoNumber + 1
    pthOut.Attribute(DEF_ATT_GEOMETRY_NUMBER) = CStr(lngGeoNumber)  ' 几何编号 = 1
    Call m_SetDetailAttributes(Door, pthOut)
End If
```

- 矩形 **Group = 0**，几何编号属性 = **1**（注释 "first outside pass should always be 1"）。
- `IgnoreOuterGeometry` 是门类型属性对话框的复选框（`&IgnoreOuterGeometry`，
  与 Width/Length/CornerRadius 同级，见 `UserStyleTestMain.gbln_CreateINI` 生成的 .ini 格式）。

### 8.3 刀路 ↔ 几何关联机制（`Make.mbln_MakeMachining`，4824 行起）

```vba
If .GroupID <> 0 Then
    Set pthsToCut = mpths_PathsInGroup(ActiveDrawing.Geometries, .GroupID)
    ' 按 Group 号选几何
Else
    ' GroupID=0：按"几何编号属性"选几何（函数名有误导性，实为
    ' pthPath.Attribute(DEF_ATT_GEOMETRY_NUMBER) = .PathOffsetFrom）
    If Not Door.IgnoreOuterGeometry Then
        Set pthsToCut = mpths_PathsNotInGroup(ActiveDrawing.Geometries, .PathOffsetFrom)
    Else
        Set pthsToCut = Nothing   ' ← IgnoreOuterGeometry=True 时 GroupID=0 刀路被强制跳过！
    End If
End If
```

**结论（本项目方案 A 的依据）：**
- `GroupID ≠ 0` 的刀路按组关联宏返回的几何 → **全部刀路应设非 0 GroupID**；
- `GroupID = 0` 的刀路关联外部几何（矩形），且勾选 `IgnoreOuterGeometry` 后会被跳过；
- 要让刀具落在自定义图形上：宏几何设固定组（如梯形=1、内偏移=2），并把
  `AD_DOOR_PATHS.GroupID` 改成对应组号。

### 8.4 读取 CDM.mdb（门类型/刀路数据库）

- 64 位 Python 无 ACE/Jet OLEDB 驱动（`未找到提供程序`）；**用 32 位 AlphaCAM VBA +
  DAO 后期绑定**读取：

```vba
Set dbe = CreateObject("DAO.DBEngine.36")
Set db = dbe.OpenDatabase("D:\2016\LICOMDAT\CDM Data\CDM.mdb", True, True)  ' 只读
```

- 关键表：`AD_DOOR_TYPES`（门类型：TypeID/Width/Length/CornerRadius/
  `IgnoreOuterGeometry`/UserVariableString...）、`AD_DOOR_PATHS`（刀路：
  PathNumber/`GroupID`/`PathOffsetFrom`/`PathOffsetSide`/`PathOffsetValue`/
  MachiningMethod/ToolName...）、`AD_USER_STYLES`（样式宏 .arb 路径）。
- **Jet DAO 的 `LIKE` 通配符默认是 `*` 不是 `%`**（ANSI-89 SQL）：
  `WHERE UserStyleName LIKE '*梯形*'`，用 `%` 匹配不到。
- 输出文件用 GBK 编码写（VBA `Print` 按系统 ANSI 代码页），Python 侧 `decode('gbk')` 读。

### 8.5 直接改数据库的风险

- CDM 运行时会用**内存中的门类型数据**覆盖数据库（用户改过 UI 后写回），
  直接 `UPDATE AD_DOOR_TYPES/AD_DOOR_PATHS` 可能被覆盖——UI 操作更可靠；
- 改库前**必须备份**（复制 `CDM.mdb`，1.2GB）；
- 改库后需重启 AlphaDOOR / 重新打开门类型才生效。

---

## 附：速查表

| 问题 | 一句话答案 |
|---|---|
| ProgID 是什么 | `aroutaps.Application` |
| 宏名格式 | `项目名.模块名.宏名` |
| 新宏 Run 不了 | 平台宏名表限制，改已有宏的代码体 |
| 模块读出的行尾 | `\r\n`；Python 写文件用 `newline=''` |
| `AddFromString` 行尾 | CRLF / LF 均可 |
| 删宏怎么删 | 定位到 `End Sub` 整块删 |
| `Dim` 放哪 | 过程顶部，勿放循环/条件块内 |
| MsgBox | 会阻塞**调用方**，但**不阻塞**别进程的 COM；验证代码里别写，改用写文件 |
| MsgBox 关不掉怎么办 | `WM_COMMAND/IDOK` 无效 → 对「确定」按钮发 `BM_CLICK`（§4.2） |
| 自动化跑会弹窗的插件 | 调用放**后台** + 独立进程**按标题**监视并记录窗口文本（§5.5） |
| `Run` 参数上限 | self + 宏名 + 9 个；更多参数就**注入临时模块**调无参过程（§2.6） |
| 仓库文本 ≠ 运行版文本 | VBA 会**重写长小数字面量**（→科学计数法）；改运行时计算如 `Pi = 4 * Atn(1)`（§3.5） |
| 读回校验怎么比 | 先**大小写归一化**再逐行 diff，否则淹没在 `.Count`→`.count` 之类差异里（§3.5） |
| `'function' object has no attribute ...` | **方法漏了括号**（Python 不能像 VBA 那样省略无参括号）（§4.8） |
| 想知道对象有哪些成员 | `gencache.EnsureDispatch(...)` 后 `dir()`；反编译 `.vb` 里的 `_VtblGap` 不代表"没有"（§4.8） |
| 刀路 `Length`/`bbox` 偏大 | 含 **rapid** 定位段；用 `GetFeedExtent`，或逐元素累加时跳过 `IsRapid`（§7.9a） |
| 无排版图纸取 NestInformation | 直接抛「现在图档内无排版」；按 §4.7 模板包起来，勿让它拖垮整个功能（§7.9c） |
| `SetStartPoint` 传路径外的点 | 文档只保证**闭合路径**、参数应在路径上；越界属未定义行为（§7.9b） |
| Open 后文件被锁 | 宏里忘 `Close`；空宏执行无参 `Close` 释放 |
| EH 里 Err 变 0 | `On Error GoTo 0` 清空 Err，先存变量再处理 |
| Drawing 几何数 | 无 `Count` 属性，用 `GetFirstGeo()` 遍历 |
| 圆弧镜像 | CW 参数取反 |
| Offset 内/外 | Left/Right 相对行进方向；镜像路径要换侧 |
| 刀路落错几何 | 宏几何用固定组号，与 `AD_DOOR_PATHS.GroupID` 对应 |
| 矩形上的刀路消失 | GroupID=0 刀路在 IgnoreOuterGeometry=True 时被跳过 |
| 原轮廓矩形在哪 | `Make.mbln_Style_Make_930`：CreateRectangle + Group=0 + 几何编号1 |
| 读 CDM.mdb | AlphaCAM VBA + `DAO.DBEngine.36`（32位）；Jet LIKE 用 `*` |
| 项目受保护 | CDM 等，跳过即可；解锁后可读 `Make`/`UserStyleTestMain` |
| Run 报"编译时遇到错误" | `APC.ApcHost.7`；动态生成的**模块名/过程名以下划线开头**（必须字母开头） |
| 新模块残留"模块N" | `module.Name` 赋值失败（下划线开头模块名非法），且清理按原名找不到 |
| 启动报"取得选项ID失败" | `CDM.arb` 的 `Licom/OptionID` 流丢失（崩溃损坏）；用 olefile 检查，从 backup 恢复并重装代码 |




