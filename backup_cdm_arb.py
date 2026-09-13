# -*- coding: utf-8 -*-
"""一键备份 AlphaCAM 的 .arb 插件资源包（CDM / CCC功能 等）。

用法:
    python backup_cdm_arb.py                        # 默认备份 CDM.arb
    python backup_cdm_arb.py --keep 10              # 只保留最近 10 份
    python backup_cdm_arb.py --arb "<...>.arb"      # 备份任意 .arb（如 CCC功能）
    python backup_cdm_arb.py --arb "<...>.arb" --tag CCC

背景:
    .arb 是 AlphaCAM 的插件资源包（OLE 复合文档），含 VBA 工程源码、窗体与配置。
    AlphaCAM **运行期间就在写它**（2026-09-13 实测：AlphaCAM 还在跑，`CCC功能.arb` 里
    已经能搜到刚部署的新代码），但：
      - **mtime 可能一直冻结**，所以【不能用 mtime 判断"是否已落盘"】；
        要判断版本请**搜内容**，而且正反两面都查（新代码独有语句应命中、
        旧代码独有语句应 0 命中）—— `.arb` 的源码保留流里会留**陈旧碎片**；
      - 文件在运行期间可能被独占，复制会失败（此时先退出 AlphaCAM）。
    CDM.arb 另有已知损坏模式：`Licom/OptionID` 流丢失
    （启动报"取得选项ID失败 / 无法打开CDM"）。本脚本会校验该流。

⚠️ 本脚本只备份【磁盘上的文件】。要备份【当前运行中】的工程源码（不依赖落盘时机），
   用 tools/running_snapshot.py <tag> <工程名> 导出各组件 —— 两者互补，建议都做。
"""
import datetime
import glob
import os
import shutil
import sys

try:
    import olefile
except ImportError:
    print("需要 olefile：pip install olefile")
    sys.exit(1)

DEFAULT_ARB = r"C:\Program Files (x86)\Vero Software\Alphacam 2016 R1\StartUp\CDM\CDM.arb"
KNOWN_ARBS = {
    "CDM": DEFAULT_ARB,
    # CCC功能 是本项目自己的插件，放在 VBMacros\Startup 下（2026-09-13 找到）
    "CCC": r"D:\2016\LICOMDIR\VBMacros\Startup\CCC功能.arb",
}

BACKUP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "backup")
KEEP = 10

argv = sys.argv[1:]
ARB = DEFAULT_ARB
TAG = None
for i, a in enumerate(argv):
    if a == "--keep" and i + 1 < len(argv):
        try:
            KEEP = int(argv[i + 1])
        except ValueError:
            pass
    elif a == "--arb" and i + 1 < len(argv):
        ARB = argv[i + 1]
    elif a == "--tag" and i + 1 < len(argv):
        TAG = argv[i + 1]
    elif a in KNOWN_ARBS:
        ARB = KNOWN_ARBS[a]
        TAG = TAG or a


def acam_running() -> bool:
    out = os.popen('tasklist /FI "IMAGENAME eq Acam.exe" /FO CSV /NH').read()
    return "Acam.exe" in out


def main():
    arb = os.path.abspath(ARB)
    base = os.path.basename(arb)
    name = TAG or base
    if not os.path.exists(arb):
        print("未找到 .arb：", arb)
        return 1
    print("目标: %s  (%d 字节, 最后写入 %s)"
          % (arb, os.path.getsize(arb),
             datetime.datetime.fromtimestamp(os.path.getmtime(arb)).strftime("%Y-%m-%d %H:%M:%S")))

    if acam_running():
        print("警告：AlphaCAM 正在运行。")
        print("  - 运行期间它可能独占该文件，复制会失败；")
        print("  - 文件通常是【较新】的（AlphaCAM 运行中就会写），但 mtime 可能冻结，")
        print("    所以别用 mtime 判断版本 —— 要判断就搜内容（正反两面都查）；")
        print("  - 想备份运行中的工程源码，另用 tools/running_snapshot.py（与本次备份互补）。")
        print("继续备份...")

    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    os.makedirs(BACKUP_DIR, exist_ok=True)
    dst = os.path.join(BACKUP_DIR, "%s_%s.bak" % (base, stamp))

    try:
        shutil.copy2(arb, dst)
    except PermissionError:
        print("复制失败：文件被锁定（AlphaCAM 占用）。请先关闭 AlphaCAM 再备份。")
        return 1

    # OLE 完整性校验
    try:
        ole = olefile.OleFileIO(dst)
        streams = ole.listdir()
        n = len(streams)
        has_opt = ole.exists("Licom/OptionID")
        has_proj = any("/".join(s).startswith("vao") for s in streams)
        ole.close()
    except Exception as e:
        n, has_opt, has_proj = 0, False, False
        print("OLE 校验失败:", e)

    print("已备份 ->", dst)
    print("大小: %d 字节 | 流数量: %d | vao/VBA 工程: %s | Licom/OptionID: %s"
          % (os.path.getsize(dst), n,
             "[有]" if has_proj else "[无!]",
             "[有]" if has_opt else "[无]"))
    if name == "CDM":
        if has_opt:
            print("CDM.arb 关键流 Licom/OptionID: [OK 存在]")
        else:
            print("!! CDM.arb 缺 Licom/OptionID —— 该备份不可用（启动会报「取得选项ID失败」）")
    elif not has_opt:
        print("说明: 自定义插件（%s）通常没有 Licom/OptionID，这不是损坏。" % name)
    if not has_proj:
        print("!! 注意: 没找到 vao/ 下的 VBA 工程流，这个 .arb 可能不含源码。")

    # 清理旧备份（按同一 .arb 的基名各自保留 KEEP 份）
    pat = os.path.join(BACKUP_DIR, "%s_*.bak" % base)
    files = sorted(glob.glob(pat))
    if len(files) > KEEP:
        for f in files[:-KEEP]:
            try:
                os.remove(f)
                print("清理旧备份:", os.path.basename(f))
            except OSError:
                pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
