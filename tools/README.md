# tools/ —— CDM / CCC 模块的部署与诊断工具

这些脚本原先散在仓库根目录、以 `_` 前缀被 `.gitignore` 忽略（一次性补丁用）。
其中 9 个是可复用的运维工具，按 `.gitignore` 里写的约定移到这里并纳入版本管理。

所有脚本都假定 **本目录的上一级 = 仓库根**，可执行路径均相对仓库根：
`python tools/<name>.py ...`

## 部署闭环（改 VBA 模块必须走这条链）

| 步骤 | 工具 | 作用 |
|---|---|---|
| 1 | `running_snapshot.py [tag] [Project] [comp1,comp2]` | 部署前把**运行中**的组件导出到 `backup/`，作为回滚点与审计基线。`Project` 默认 `CDM`（默认组件集 Make/modAutoImportNest/frmAutoNest/Events）；其它工程省略组件名则导出**全部**组件 |
| 2 | `component_deploy.py audit <组件> <仓库相对路径> <基线 glob> [Project]` | 比对「运行中 / 仓库 / 基线」。**只有运行版等于仓库或等于基线时才允许写**，防止覆盖掉别人在 VBA 编辑器里的手改 |
| 3 | `component_deploy.py deploy ...` | `DeleteLines` + `AddFromString` 整体替换，**读回校验**，校验失败自动回滚 |
| 4 | `deploy_probe.py`（CDM） / `ccc_probe.py`（CCC） | 断言新版标记在位、旧代码已消失，并用 `App.Run` 触发**全工程编译**探针 |

比较一律**大小写不敏感**：VBA 会把成员名归一化（`.Add` → `.add`），不这样做会误报部署失败。

> ⚠️ **不要写长小数字面量**。VBA 保存时会把 `0.0174532925199433` 改写成
> `1.74532925199433E-02`，仓库文本与运行文本就永远对不上（`deploy` 的读回校验会失败）。
> 用运行时计算代替（例：`Pi = 4 * Atn(1)`）。

### 一次典型的 CCC 部署

```bash
python tools/running_snapshot.py rampv2 "CCC功能" modRamp
python tools/component_deploy.py audit  modRamp "CCC功能/modRamp.bas" "backup/alphacam_running_modRamp_rampv2_*.bas" "CCC功能"
python tools/component_deploy.py deploy modRamp "CCC功能/modRamp.bas" "backup/alphacam_running_modRamp_rampv2_*.bas" "CCC功能"
python tools/ccc_probe.py
```

## 诊断工具

| 工具 | 用途 |
|---|---|
| `gbk_read.py <file> <start> [end]` | 按行打印 **GBK** 源码（`CDM功能/*.bas` 是 GBK+LF，read 类工具读不了） |
| `db_query.py [JobName]` | 查 `CDM.mdb`：本项目 Python 是 64 位而 Jet 4.0 只有 32 位，**只能**借 AlphaCAM 的 VBA 走 DAO/ADO，这是唯一通路 |
| `db_schema_probe.py [JobName]` | 探测/补建 `AD_REPORT_DATA` 字段并列出本单各行（PK/DetailID/件序号/UID） |
| `nest_dump.py` | 导出当前排版图的结构：每板每件的路径数、`DetailID`、件序号、UID |
| `match_replay.py` | 用上面两份真实 dump **复演**标签行配对算法，验证不会再出现重复标签 |
| `win_ctl.py` | 枚举顶层窗口（PID/类名/可见性）、置前、按标题关闭 —— AlphaCAM 弹模态框或错误处理刷屏记事本时用它清场 |

## 弹窗 / 模态框（实测规律，2026-09-13）

自动化调用 VBA（`App.Run` / 注入临时模块）时，被调用过程里的 `MsgBox` 会**卡住调用方**，
但从**其它进程**发来的 COM 调用**照样能进**。两条实测结论：

| 结论 | 证据 |
|---|---|
| **AlphaCAM 的模态框不阻塞 COM** | 回执框还挂在屏幕上时，另一个进程的 COM 调用正常返回（与 `SKILL.md` 里 `ReadTextFile` 那条一致） |
| **`WM_COMMAND`/`IDOK` 关不掉 VBA 的 `MsgBox`** | 对 `#32770` 窗口连发 8 次 `PostMessage(WM_COMMAND, IDOK)` 无效；改用**对子按钮 `SendMessage(BM_CLICK)`** 一次即关 |

`win_ctl.py --dismiss <pid>` 目前走的是 `WM_COMMAND/IDOK` —— **对 AlphaCAM 的 VBA `MsgBox` 无效**，
需要改成找 `Button` 子窗口发 `BM_CLICK`（见 `tmp/dismiss_hard.py` 的写法）。
另外 `--dismiss` 只按窗口类 `#32770` 过滤，**会误伤其它程序**（实测误点了 360压缩 的解压进度窗），
建议一律**按标题精确匹配**。

### 跑会弹窗的自动化时的正确姿势

1. **放后台**：长任务用后台作业跑，别让前台调用被回执框卡死（`python -u` 保证日志实时）。
2. **配一个独立监视进程**：轮询目标标题的窗口，**记录其文本**（回执框内容往往是唯一的
   结果摘要，例：`候选/受理/已应用/跳过/微连接` 计数），再点掉它。
   `tmp/watch_popups.py` 是这个模式的参考实现。
3. 关窗只按标题匹配，**绝不按窗口类匹配**。

## 注意

- 这些脚本会**写**运行中的 AlphaCAM/CDM 工程与数据库，属于操作类工具，跑之前先 `running_snapshot.py`。
- `tmp/` 下是诊断产物（dump、PNG、日志），随时可删。
- `backup/CDM_snapshot_*/` 是 CDM 工程全量导出（59 个组件，含仓库未镜像的类如
  `CRouterReportData`），排错时很有用。
- **API 语义直接查本地文档**，不要外推：`<AlphaCAM 安装目录>\tempacamapi\` 是解压好的
  ACAMAPI 参考（862 个 .htm）。`tmp/readdoc.py <页面名>` 可转成纯文本打印。
  本次就是靠它纠正了 `StockZ` 的误解、确认了 `SetStartPoint` 只对闭合路径有效。
