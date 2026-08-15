# -*- coding: utf-8 -*-
"""一键备份 AlphaCAM CDM.arb（当前可用版本）。

用法:
    python backup_cdm_arb.py            # 备份到 backup/CDM.arb_<时间戳>.bak
    python backup_cdm_arb.py --keep 10  # 只保留最近 10 份

背景:
    CDM.arb 是 AlphaCAM 的 CDM 插件资源包（OLE 复合文档），含 VBA 工程源码、
    窗体与 Licom/OptionID 配置。AlphaCAM 退出/保存时会把内存 VBA 工程写回该文件，
    若期间崩溃可能损坏（OptionID 流丢失 → 启动报"取得选项ID失败"）。
    定期备份可快速恢复。
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

CDM_ARB = r"C:\Program Files (x86)\Vero Software\Alphacam 2016 R1\StartUp\CDM\CDM.arb"
BACKUP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "backup")
KEEP = 10
for i, a in enumerate(sys.argv[1:]):
    if a == "--keep" and i + 2 < len(sys.argv):
        try:
            KEEP = int(sys.argv[i + 2])
        except ValueError:
            pass


def acam_running() -> bool:
    out = os.popen('tasklist /FI "IMAGENAME eq Acam.exe" /FO CSV /NH').read()
    return "Acam.exe" in out


def main():
    if not os.path.exists(CDM_ARB):
        print("未找到 CDM.arb：", CDM_ARB)
        sys.exit(1)

    if acam_running():
        print("警告：AlphaCAM 正在运行。")
        print("  - 复制可能拿到旧版本（AlphaCAM 内存工程未保存的最新改动不在文件里）；")
        print("  - 建议先让 AlphaCAM 保存（或退出）后再备份，确保含最新代码。")
        print("继续备份（文件可能被锁定或过期）...")

    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = os.path.join(BACKUP_DIR, "CDM.arb_%s.bak" % stamp)
    os.makedirs(BACKUP_DIR, exist_ok=True)

    try:
        shutil.copy2(CDM_ARB, dst)
    except PermissionError:
        print("复制失败：文件被锁定（AlphaCAM 占用）。请先关闭 AlphaCAM 再备份。")
        sys.exit(1)

    # 校验 OLE 完整性
    try:
        ole = olefile.OleFileIO(dst)
        ok = ole.exists("Licom/OptionID")
        n = len(ole.listdir())
        ole.close()
    except Exception as e:
        ok, n = False, 0
        print("OLE 校验失败:", e)

    print("已备份 ->", dst)
    print("大小: %d 字节 | 流数量: %d | Licom/OptionID: %s" % (
        os.path.getsize(dst), n, "[OK 存在]" if ok else "[缺失 - 该备份不可用!]"))

    # 清理旧备份（保留最近 KEEP 份）
    pat = os.path.join(BACKUP_DIR, "CDM.arb_*.bak")
    files = sorted(glob.glob(pat))
    if len(files) > KEEP:
        for f in files[:-KEEP]:
            try:
                os.remove(f)
                print("清理旧备份:", f)
            except OSError:
                pass


if __name__ == "__main__":
    main()
