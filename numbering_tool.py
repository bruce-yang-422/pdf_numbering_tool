#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF Numbering Tool

支援：
- 批次處理
- 檔名排序
- 使用者選擇單一或全部 PDF
- 兩種編號模式：每檔案從 1 開始 / 所有檔案連續編號

Authors: 楊翔志 & AI Collective
Studio: tranquility-base
Email: bruce.yichai20250505@gmail.com
版本: 1.0 (2025-11-14)
"""

import sys
import logging
from pathlib import Path
from io import BytesIO
from datetime import datetime
from reportlab.pdfgen import canvas
from pypdf import PdfReader, PdfWriter


# -------------------------------------------------------------
# 自我定位：獲取程式所在目錄（不依賴當前工作目錄）
# -------------------------------------------------------------
def get_script_dir():
    """
    獲取程式所在目錄的絕對路徑
    支援：
    - 直接執行 Python 腳本
    - 打包成 EXE（PyInstaller）
    - 從任何位置執行
    """
    if getattr(sys, 'frozen', False):
        # 打包成 EXE 的情況（PyInstaller）
        # sys.executable 是 EXE 的路徑
        return Path(sys.executable).parent.resolve()
    else:
        # 直接執行 Python 腳本的情況
        # __file__ 是當前腳本的路徑
        return Path(__file__).parent.resolve()


# 獲取程式所在目錄（全域變數，只計算一次）
SCRIPT_DIR = get_script_dir()


# -------------------------------------------------------------
# 設定 Logging
# -------------------------------------------------------------
LOG_FILE = SCRIPT_DIR / "numbering_tool.txt"


def setup_logging():
    """設定 logging，輸出到固定檔名並覆蓋舊檔案"""
    # 設定 logging 格式
    log_format = "%(asctime)s - %(levelname)s - %(message)s"
    date_format = "%Y-%m-%d %H:%M:%S"
    
    # 設定 logging，使用 'w' 模式覆蓋舊檔案
    logging.basicConfig(
        level=logging.INFO,
        format=log_format,
        datefmt=date_format,
        handlers=[
            logging.FileHandler(str(LOG_FILE), mode='w', encoding='utf-8'),  # 覆蓋模式
            logging.StreamHandler(sys.stdout)  # 同時輸出到控制台
        ]
    )
    
    logger = logging.getLogger(__name__)
    logger.info("=" * 60)
    logger.info("PDF Numbering Tool 開始執行")
    logger.info(f"執行時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 60)
    
    return logger


# -------------------------------------------------------------
# 讀取設定檔 coords.env
# -------------------------------------------------------------
def load_config(config_path=None, logger=None):
    """讀取設定檔 coords.env"""
    if config_path is None:
        # 使用程式所在目錄的 coords.env
        config_path = SCRIPT_DIR / "coords.env"
    else:
        # 如果是相對路徑，轉換為基於程式目錄的絕對路徑
        config_path = SCRIPT_DIR / config_path
    
    config = {}

    if not config_path.exists():
        error_msg = f"錯誤：找不到設定檔 {config_path}"
        if logger:
            logger.error(error_msg)
        print(error_msg)
        sys.exit(1)

    if logger:
        logger.info(f"讀取設定檔：{config_path}")

    with open(str(config_path), "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()

            if not line or line.startswith("#"):
                continue

            if "=" in line:
                key, value = line.split("=", 1)
                key, value = key.strip(), value.strip()

                if key in ["X1", "Y1", "X2", "Y2", "DIGITS", "PAD"]:
                    try:
                        config[key] = int(value)
                    except:
                        config[key] = 0

                elif key in ["DRAW_BOX", "DRAW_CIRCLE"]:
                    config[key] = 1 if value == "1" else 0

    required = ["X1", "Y1", "X2", "Y2", "DIGITS", "PAD", "DRAW_BOX", "DRAW_CIRCLE"]
    for key in required:
        if key not in config:
            config[key] = 0

    if logger:
        logger.info(f"設定載入完成：編號位數={config['DIGITS']}, 位置1=({config['X1']}, {config['Y1']}), "
                   f"位置2=({config['X2']}, {config['Y2']}), 方框={config['DRAW_BOX']}, 圓框={config['DRAW_CIRCLE']}")

    return config


# -------------------------------------------------------------
# 檔名排序
# -------------------------------------------------------------
def extract_prefix_sort_key(filename):
    """提取檔名前綴作為排序鍵，優先順序：純數字 > 字母開頭 > 其他"""
    name = filename.stem
    # 提取第一個前綴（以 _ 或 - 分隔）
    prefix = name.split("_")[0].split("-")[0]

    # 純數字（如日期 20251031）
    if prefix.isdigit():
        return (0, int(prefix))
    # 字母開頭
    if prefix and prefix[0].isalpha():
        return (1, prefix.lower())
    # 其他情況
    return (2, prefix.lower() if prefix else "")


def find_all_pdfs_with_selection(input_dir=None, logger=None):
    """尋找所有 PDF 並讓使用者選擇要處理的檔案"""
    if input_dir is None:
        # 使用程式所在目錄的 input 資料夾
        input_path = SCRIPT_DIR / "input"
    else:
        # 如果是相對路徑，轉換為基於程式目錄的絕對路徑
        input_path = SCRIPT_DIR / input_dir

    if not input_path.exists():
        error_msg = f"錯誤：找不到資料夾 {input_path}"
        if logger:
            logger.error(error_msg)
        print(error_msg)
        sys.exit(1)

    pdf_files = list(input_path.glob("*.pdf"))
    if not pdf_files:
        error_msg = "錯誤：資料夾內沒有 PDF"
        if logger:
            logger.error(error_msg)
        print(error_msg)
        sys.exit(1)

    # 按檔名前綴排序
    pdf_files = sorted(pdf_files, key=extract_prefix_sort_key)

    if logger:
        logger.info(f"在 {input_path} 資料夾中找到 {len(pdf_files)} 個 PDF 檔案")

    # 如果只有一個檔案，直接返回，跳過選擇步驟
    if len(pdf_files) == 1:
        print(f"\n找到 1 個 PDF：{pdf_files[0].name}")
        if logger:
            logger.info(f"僅找到一個檔案，自動選擇：{pdf_files[0].name}")
        return pdf_files

    # 多個檔案時才需要選擇
    print(f"\n找到 {len(pdf_files)} 個 PDF：")
    for idx, pdf in enumerate(pdf_files, 1):
        print(f"  {idx:>2}) {pdf.name}")
        if logger:
            logger.info(f"  {idx:>2}) {pdf.name}")

    print()
    try:
        choice = input("請輸入要處理的序號（或 ALL 全部）：").strip().upper()
    except KeyboardInterrupt:
        if logger:
            logger.warning("使用者中斷程式")
        print("\n\n程式已取消")
        sys.exit(0)

    selected_files = []
    if choice == "ALL" or choice == "":
        selected_files = pdf_files
        if logger:
            logger.info("使用者選擇處理全部檔案")
    else:
        try:
            num = int(choice)
            if 1 <= num <= len(pdf_files):
                selected_files = [pdf_files[num - 1]]
                if logger:
                    logger.info(f"使用者選擇處理第 {num} 個檔案：{pdf_files[num - 1].name}")
            else:
                if logger:
                    logger.warning(f"輸入無效：{choice}，將處理全部檔案")
                print("錯誤：輸入無效，將處理全部檔案")
                selected_files = pdf_files
        except ValueError:
            if logger:
                logger.warning(f"輸入無效：{choice}，將處理全部檔案")
            print("錯誤：輸入無效，將處理全部檔案")
            selected_files = pdf_files

    return selected_files


# -------------------------------------------------------------
# 編號
# -------------------------------------------------------------
def format_number(number, digits):
    return str(number).zfill(digits)


def create_number_overlay(number1, number2, config, page_width, page_height):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=(page_width, page_height))

    font_size = 12
    c.setFont("Helvetica-Bold", font_size)

    h = font_size

    # --- 第一個編號 ---
    x1, y1 = config["X1"], config["Y1"]
    t1 = format_number(number1, config["DIGITS"])
    w1 = c.stringWidth(t1, "Helvetica-Bold", font_size)

    if config["DRAW_BOX"]:
        c.rect(x1 - config["PAD"], y1 - config["PAD"], w1 + config["PAD"] * 2, h + config["PAD"] * 2)
    elif config["DRAW_CIRCLE"]:
        radius = max(w1, h) / 2 + config["PAD"]
        c.circle(x1 + w1 / 2, y1 + h / 2, radius)

    c.drawString(x1, y1, t1)

    # --- 第二個編號 ---
    x2, y2 = config["X2"], config["Y2"]
    t2 = format_number(number2, config["DIGITS"])
    w2 = c.stringWidth(t2, "Helvetica-Bold", font_size)

    if config["DRAW_BOX"]:
        c.rect(x2 - config["PAD"], y2 - config["PAD"], w2 + config["PAD"] * 2, h + config["PAD"] * 2)
    elif config["DRAW_CIRCLE"]:
        radius = max(w2, h) / 2 + config["PAD"]
        c.circle(x2 + w2 / 2, y2 + h / 2, radius)

    c.drawString(x2, y2, t2)

    c.save()
    buffer.seek(0)
    return buffer


# -------------------------------------------------------------
# 處理 PDF
# -------------------------------------------------------------
def process_pdf(input_pdf_path, output_pdf_path, start_number, config, logger=None):
    """處理 PDF，在每頁加入編號"""
    if logger:
        logger.info(f"開始處理：{input_pdf_path.name} -> {output_pdf_path.name}")
        logger.info(f"起始編號：{start_number}")

    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()

    current_number = start_number
    total_pages = len(reader.pages)

    if logger:
        logger.info(f"PDF 總頁數：{total_pages}")

    print(f"  共 {total_pages} 頁")

    for page_index, page in enumerate(reader.pages, 1):
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)

        num1 = current_number
        num2 = current_number + 1
        current_number += 2

        num1_str = format_number(num1, config['DIGITS'])
        num2_str = format_number(num2, config['DIGITS'])
        
        print(f"    → 第 {page_index} 頁：{num1_str} / {num2_str}")
        
        if logger:
            logger.info(f"  第 {page_index}/{total_pages} 頁：編號 {num1_str} / {num2_str}")

        try:
            overlay_buf = create_number_overlay(num1, num2, config, page_width, page_height)
            overlay_reader = PdfReader(overlay_buf)
            overlay_page = overlay_reader.pages[0]

            page.merge_page(overlay_page)
            writer.add_page(page)
        except Exception as e:
            error_msg = f"處理第 {page_index} 頁時發生錯誤：{str(e)}"
            if logger:
                logger.error(error_msg)
            raise Exception(error_msg)

    # 確保輸出目錄存在（基於程式目錄）
    output_dir = Path(output_pdf_path).parent
    output_dir.mkdir(parents=True, exist_ok=True)
    with open(str(output_pdf_path), "wb") as f:
        writer.write(f)

    if logger:
        logger.info(f"完成處理：{output_pdf_path.name}，編號範圍 {format_number(start_number, config['DIGITS'])} ~ {format_number(current_number - 1, config['DIGITS'])}")

    return current_number  # 回傳下一份 PDF 的起始值


# -------------------------------------------------------------
# 主程式
# -------------------------------------------------------------
def main():
    # 設定 logging
    logger = setup_logging()
    
    print("=" * 50)
    print("PDF Numbering Tool")
    print("=" * 50)

    # 使用自我定位的路徑（不依賴當前工作目錄）
    config = load_config(None, logger)  # None 表示使用預設路徑（程式目錄下的 coords.env）
    print("✓ 已載入設定\n")

    pdf_list = find_all_pdfs_with_selection(None, logger)  # None 表示使用預設路徑（程式目錄下的 input）
    print()

    # --- 選擇編號模式 ---
    # 如果只有一個檔案，跳過選擇步驟（兩種模式效果相同）
    if len(pdf_list) == 1:
        mode = "1"
        if logger:
            logger.info("僅有一個檔案，自動使用預設編號模式（每個檔案從相同起始編號開始）")
    else:
        print("編號模式：")
        print("  1) 每個檔案皆從相同起始編號開始（預設）")
        print("  2) 所有檔案連續編號（上一份接下一份）")
        try:
            mode = input("請選擇 1 或 2（預設 1）：").strip()
        except KeyboardInterrupt:
            logger.warning("使用者中斷程式")
            print("\n\n程式已取消")
            sys.exit(0)

        if mode not in ["1", "2"]:
            mode = "1"
            print("使用預設模式：每個檔案從相同起始編號開始")
    
    mode_name = "每個檔案從相同起始編號開始" if mode == "1" else "所有檔案連續編號"
    logger.info(f"編號模式：{mode_name}")

    # --- 輸入起始編號 ---
    while True:
        try:
            start_input = input("請輸入起始編號（預設 1）：").strip()
            base_start = int(start_input) if start_input else 1
            if base_start < 1:
                print("起始編號必須大於 0")
                continue
            break
        except ValueError:
            print("請輸入有效整數")
        except KeyboardInterrupt:
            logger.warning("使用者中斷程式")
            print("\n\n程式已取消")
            sys.exit(0)

    logger.info(f"起始編號：{base_start}")
    print()

    # --- 執行處理 ---
    next_number = base_start
    total_files = len(pdf_list)
    
    logger.info(f"開始處理 {total_files} 個 PDF 檔案")

    success_count = 0
    fail_count = 0

    for idx, pdf_path in enumerate(pdf_list, 1):
        print(f"\n[{idx}/{total_files}] 處理：{pdf_path.name}")

        if mode == "1":
            start_number = base_start
        else:
            start_number = next_number

        # 生成輸出檔名（避免覆蓋，使用程式目錄下的 output 資料夾）
        output_pdf_path = SCRIPT_DIR / "output" / f"{pdf_path.stem}_numbered.pdf"
        
        try:
            next_number = process_pdf(pdf_path, output_pdf_path, start_number, config, logger)
            print(f"  ✓ 完成：{output_pdf_path.name}")
            success_count += 1
        except Exception as e:
            error_msg = f"處理 {pdf_path.name} 時發生問題：{str(e)}"
            print(f"  ✗ 錯誤：{error_msg}")
            logger.error(error_msg, exc_info=True)
            fail_count += 1
            continue

    # 總結
    summary = f"\n處理完成：成功 {success_count} 個，失敗 {fail_count} 個"
    logger.info("=" * 60)
    logger.info(summary)
    logger.info(f"Log 檔案已儲存至：{LOG_FILE}")
    logger.info(f"程式所在目錄：{SCRIPT_DIR}")
    logger.info("=" * 60)

    print("\n" + "=" * 50)
    print("全部完成！結果已輸出到 output 資料夾。")
    print(f"Log 檔案：{LOG_FILE}")
    print("=" * 50 + "\n")


if __name__ == "__main__":
    main()
