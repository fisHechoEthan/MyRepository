#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
递归删除"已成功解压"的压缩包 —— 默认直接删除版

行为：
  1. 扫描所有 zip/rar（严格排除 docx/xlsx/pptx 等伪压缩包）
  2. 只删【能找到同名非空解压目录】的压缩包
  3. 校验通过（未匹配=0 且无空目录）→ 直接删除
  4. 校验不通过 → 拒绝删除，打印原因，等人工处理
  5. 只删文件，绝不删任何目录
"""

import os
import sys
import logging
from pathlib import Path

# ================= 配置区 =================
TARGET_DIR = r"D:\download\vs code\projects\材料包"   # ← 改成你的实际路径
LOG_FILE = "clean_log.txt"

# 是否允许"跳过校验强制删除"（默认 False = 校验不过就不删）
FORCE_DELETE_WHEN_FAIL = False
# =========================================

ARCHIVE_EXTS = [".rar", ".zip"]

PSEUDO_ARCHIVE_EXTS = {
    ".docx", ".docm", ".dotx", ".dotm",
    ".xlsx", ".xlsm", ".xltx", ".xltm",
    ".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm",
    ".odt", ".ods", ".odp", ".odg",
    ".jar", ".apk", ".epub", ".whl", ".nupkg", ".vsix", ".ipa",
    ".xpi", ".crx",
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)


def is_archive(path: Path) -> bool:
    suffix = path.suffix.lower()
    if suffix in PSEUDO_ARCHIVE_EXTS:
        return False
    if suffix not in ARCHIVE_EXTS:
        return False
    try:
        with open(path, "rb") as f:
            head = f.read(8)
    except Exception:
        return False
    if head[:4] in (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08"):
        return True
    if head[:7] == b"Rar!\x1a\x07\x00":
        return True
    if head[:8] == b"Rar!\x1a\x07\x01\x00":
        return True
    return False


def extract_dir_of(archive: Path) -> Path:
    base = archive.stem if archive.suffix else archive.name + "_out"
    return archive.parent / base


def human_size(p: Path) -> str:
    try:
        n = p.stat().st_size
        for u in ["B", "KB", "MB", "GB"]:
            if n < 1024:
                return f"{n:.1f}{u}"
            n /= 1024
        return f"{n:.1f}TB"
    except Exception:
        return "?"


def count_files_in_dir(d: Path) -> int:
    n = 0
    for _, _, files in os.walk(d):
        n += len(files)
    return n


def find_all_archives(root: Path):
    result = []
    for dirpath, _, filenames in os.walk(root):
        for fn in filenames:
            p = Path(dirpath) / fn
            if is_archive(p):
                result.append(p)
    return result


def main():
    # 支持命令行传路径覆盖配置
    if len(sys.argv) > 1:
        target = sys.argv[1]
    else:
        target = TARGET_DIR

    root = Path(target).resolve()
    if not root.exists():
        log.error(f"❌ 目录不存在：{root}")
        sys.exit(1)

    log.info(f"扫描根目录：{root}")
    log.info("=" * 70)

    archives = find_all_archives(root)

    to_delete = []   # (压缩包, 目录, 文件数)
    keep = []        # (压缩包, 期望目录)

    for arc in archives:
        tgt = extract_dir_of(arc)
        if tgt.exists() and tgt.is_dir():
            nfiles = count_files_in_dir(tgt)
            to_delete.append((arc, tgt, nfiles))
        else:
            keep.append((arc, tgt))

    # ==================== 统计报告 ====================
    log.info("")
    log.info("╔══════════════════════════════════════════════╗")
    log.info("║              数 量 统 计 报 告                ║")
    log.info("╚══════════════════════════════════════════════╝")
    log.info(f"  识别到的压缩包总数       : {len(archives)}")
    log.info(f"  ├─ 已找到同名解压目录     : {len(to_delete)}")
    log.info(f"  └─ 未找到同名解压目录     : {len(keep)}")

    total_files_in_dirs = sum(n for _, _, n in to_delete)
    log.info(f"  待删压缩包对应的目录文件数 : {total_files_in_dirs}")

    zip_n = sum(1 for a, _, _ in to_delete if a.suffix.lower() == ".zip")
    rar_n = sum(1 for a, _, _ in to_delete if a.suffix.lower() == ".rar")
    log.info(f"  待删中 .zip : {zip_n}    .rar : {rar_n}")

    empty_dirs = [(a, t) for a, t, n in to_delete if n == 0]
    if empty_dirs:
        log.warning(f"  ⚠️ 有 {len(empty_dirs)} 个压缩包对应的目录是空的（0 个文件）")
        for a, t in empty_dirs[:10]:
            log.warning(f"      {a.name}  ->  {t.name}/")

    log.info("═" * 50)

    # ==================== 结论 ====================
    no_match_ok = (len(keep) == 0)
    no_empty_ok = (len(empty_dirs) == 0)
    all_ok = no_match_ok and no_empty_ok

    if len(archives) == 0:
        log.info("✅ 没有发现任何压缩包，无需处理")
        return

    if all_ok:
        log.info(f"✅ 结论：全部 {len(archives)} 个压缩包都能匹配到非空解压目录")
    else:
        if not no_match_ok:
            log.info(f"⚠️ 有 {len(keep)} 个压缩包没找到对应解压目录")
        if not no_empty_ok:
            log.info(f"⚠️ 有 {len(empty_dirs)} 个压缩包对应的目录为空")
    log.info("═" * 50)

    # ==================== 明细 ====================
    if keep:
        log.info("")
        log.info(f"【将保留的压缩包】（{len(keep)} 个，未找到同名解压目录）")
        for arc, tgt in keep:
            log.info(f"  {arc}   [{human_size(arc)}]")
            log.info(f"      期望目录不存在：{tgt}")

    log.info("")
    log.info("=" * 70)

    # ==================== 执行 ====================
    if not all_ok and not FORCE_DELETE_WHEN_FAIL:
        log.info("🛑 校验未通过，拒绝删除")
        if not no_match_ok:
            log.info(f"   原因：{len(keep)} 个压缩包未找到同名解压目录")
            log.info(f"   请人工核对方才【将保留的压缩包】列表")
        if not no_empty_ok:
            log.info(f"   原因：{len(empty_dirs)} 个目录为空，解压可能失败")
        log.info("")
        log.info("   如确认要强制删除，请把脚本里 FORCE_DELETE_WHEN_FAIL 改为 True")
        return

    if not all_ok and FORCE_DELETE_WHEN_FAIL:
        log.warning("🔧 FORCE_DELETE_WHEN_FAIL=True，跳过校验强制删除")

    log.info(f"🚀 开始删除 {len(to_delete)} 个压缩包...")
    deleted, failed = 0, 0
    for arc, _, _ in to_delete:
        try:
            arc.unlink()
            deleted += 1
            log.info(f"  ✅ 已删除 {arc.name}")
        except Exception as e:
            failed += 1
            log.error(f"  ❌ 删除失败 {arc.name}：{e}")

    log.info("=" * 70)
    log.info(f"删除完成：成功 {deleted}，失败 {failed}，保留 {len(keep)}")


if __name__ == "__main__":
    main()