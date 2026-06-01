"""
safe_move.py — 安全移動/整理檔案工具
=====================================
核心原則：
  1. 先備份 ZIP，再做任何事
  2. 先 DRY-RUN 顯示計畫，確認後才執行
  3. 只用 COPY+VERIFY+DELETE，絕不用 rename/move
  4. 目的地有同名檔案 → 直接跳過（絕不覆蓋）
  5. 每一步都寫 LOG
"""

import os
import shutil
import hashlib
import zipfile
import datetime
import logging
from pathlib import Path


# ============================================================
# LOG 設定
# ============================================================
def setup_logger(log_path: Path) -> logging.Logger:
    logger = logging.getLogger("safe_move")
    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s",
                            datefmt="%Y-%m-%d %H:%M:%S")
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setFormatter(fmt)
    ch = logging.StreamHandler()
    ch.setFormatter(fmt)
    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger


# ============================================================
# 核心：備份整個資料夾為 ZIP
# ============================================================
def backup_folder(folder: Path, backup_dir: Path, logger: logging.Logger) -> Path:
    """把 folder 整個壓縮成 ZIP，返回 ZIP 路徑"""
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_name = f"BACKUP_{folder.name}_{ts}.zip"
    zip_path = backup_dir / zip_name

    backup_dir.mkdir(parents=True, exist_ok=True)
    logger.info(f"開始備份: {folder} → {zip_path}")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        count = 0
        for root, dirs, files in os.walk(folder):
            # ★ 跳過備份資料夾本身，防止無限遞迴
            dirs[:] = [d for d in dirs
                       if Path(root, d).resolve() != backup_dir.resolve()]
            for fname in files:
                src = Path(root) / fname
                # ★ 跳過 ZIP 本身（防止寫入自己）
                if src.resolve() == zip_path.resolve():
                    continue
                arcname = src.relative_to(folder)
                zf.write(src, arcname)
                count += 1
        logger.info(f"備份完成：{count} 個檔案 → {zip_path}")

    return zip_path


# ============================================================
# 核心：安全複製（COPY + 驗證 MD5 + 刪除原始）
# ============================================================
def md5(path: Path) -> str:
    h = hashlib.md5()
    with open(path, "rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()


def safe_move_file(src: Path, dst_dir: Path, logger: logging.Logger,
                   dry_run: bool = False) -> str:
    """
    安全移動一個檔案到 dst_dir。
    返回: "moved" / "skipped" / "error"
    """
    dst_dir.mkdir(parents=True, exist_ok=True)
    dst = dst_dir / src.name

    # ★ 目的地有同名檔案 → 跳過，絕不覆蓋
    if dst.exists():
        logger.warning(f"[SKIP] 目的地已存在，略過: {dst}")
        return "skipped"

    if dry_run:
        logger.info(f"[DRY-RUN] 會移動: {src.name}  →  {dst_dir.name}/")
        return "dry"

    try:
        # Step 1: 複製
        shutil.copy2(src, dst)
        logger.debug(f"[COPY] {src.name} → {dst_dir.name}/")

        # Step 2: 驗證 MD5
        src_md5 = md5(src)
        dst_md5 = md5(dst)
        if src_md5 != dst_md5:
            dst.unlink()  # 複製失敗，刪掉損壞的目的地
            logger.error(f"[ERROR] MD5 不符，已刪除目的地: {src.name}")
            return "error"

        # Step 3: 確認複製成功後才刪除原始
        src.unlink()
        logger.info(f"[OK] {src.name} → {dst_dir.name}/")
        return "moved"

    except Exception as e:
        logger.error(f"[ERROR] 移動失敗 {src.name}: {e}")
        # 如果目的地已存在（部分寫入），清除它
        if dst.exists():
            try:
                dst.unlink()
            except Exception:
                pass
        return "error"


# ============================================================
# 主執行：整理 COA 資料夾
# ============================================================
def organize_coa(coa_root: str, dry_run: bool = True):
    root = Path(coa_root)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    # LOG 存放位置（放在 CODE資料）
    code_dir = Path(r"C:\Users\admin\OneDrive\桌面\CODE資料")
    log_dir = code_dir / "_COA_LOG"
    log_dir.mkdir(exist_ok=True)
    log_path = log_dir / f"safe_move_{ts}.log"
    logger = setup_logger(log_path)

    if dry_run:
        logger.info("=" * 60)
        logger.info("★★★  DRY-RUN 模式（只顯示計畫，不實際移動）★★★")
        logger.info("=" * 60)
    else:
        logger.info("=" * 60)
        logger.info("★★★  正式執行模式 ★★★")
        logger.info("=" * 60)

        # ★ 先備份，再做任何事（備份放在 CODE資料，避免遞迴）
        code_dir = Path(r"C:\Users\admin\OneDrive\桌面\CODE資料")
        backup_dir = code_dir / "_COA_BACKUP"
        backup_folder(root, backup_dir, logger)

    # ============================================================
    # 分類規則
    # ============================================================
    rules = {
        # 松勝COA
        "松勝COA": [
            {
                "dst": "松勝COA/鼻壓條鼻線COA",
                "keywords": ["鼻線", "鼻壓條", "PW-"],
            },
            {
                "dst": "松勝COA/安全測試報告",
                "keywords": ["細胞毒", "皮膚", "甲醛", "螢光", "SGS",
                             "偶氮", "MSDS", "M62-"],
            },
            {
                "dst": "松勝COA/外層布COA",
                "keywords": [],   # 預設（其他全部）
                "default": True,
            },
        ],
        # 福綿COA
        "福綿COA": [
            {
                "dst": "福綿COA/印刷布COA",
                "keywords": ["印刷布", "sPET", "抹茶", "灰紫藍", "40G", "乾燥"],
            },
            {
                "dst": "福綿COA/安全測試報告",
                "keywords": ["細胞毒", "皮膚", "螢光", "偶氮", "甲醛",
                             "Full Men", "Full_Men", "ISO", "KE_2018",
                             "TWNC", "纖維"],
            },
            {
                "dst": "福綿COA/親膚布COA",
                "keywords": [],
                "default": True,
            },
        ],
    }

    stats = {"moved": 0, "skipped": 0, "error": 0, "dry": 0}

    for src_folder_name, rule_list in rules.items():
        src_folder = root / src_folder_name
        if not src_folder.exists():
            logger.warning(f"資料夾不存在，略過: {src_folder}")
            continue

        files = [f for f in src_folder.iterdir() if f.is_file()]
        logger.info(f"\n處理 {src_folder_name}（{len(files)} 個檔案）")

        for f in files:
            matched = False
            for rule in rule_list:
                if rule.get("default"):
                    continue  # 預設規則最後再用
                if any(kw in f.name for kw in rule["keywords"]):
                    dst_dir = root / rule["dst"]
                    result = safe_move_file(f, dst_dir, logger, dry_run)
                    stats[result] += 1
                    matched = True
                    break

            if not matched:
                # 使用預設規則
                for rule in rule_list:
                    if rule.get("default"):
                        dst_dir = root / rule["dst"]
                        result = safe_move_file(f, dst_dir, logger, dry_run)
                        stats[result] += 1
                        break

    logger.info("\n" + "=" * 60)
    logger.info(f"完成！移動:{stats['moved']}  跳過:{stats['skipped']}  "
                f"錯誤:{stats['error']}  預覽:{stats['dry']}")
    logger.info(f"LOG 存放於: {log_path}")
    logger.info("=" * 60)

    return stats


# ============================================================
# 進入點
# ============================================================
if __name__ == "__main__":
    import sys

    COA_ROOT = r"C:\Users\admin\OneDrive\桌面\COA"

    # 預設 DRY-RUN，加 --go 才真正執行
    is_dry_run = "--go" not in sys.argv

    if is_dry_run:
        print("\n" + "="*60)
        print("  DRY-RUN 模式：只顯示計畫，不移動任何檔案")
        print("  確認計畫無誤後，執行：python safe_move.py --go")
        print("="*60 + "\n")
    else:
        print("\n" + "="*60)
        print("  正式執行：先備份 ZIP，再移動檔案")
        print("="*60 + "\n")

    organize_coa(COA_ROOT, dry_run=is_dry_run)
