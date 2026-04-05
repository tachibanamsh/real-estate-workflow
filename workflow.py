"""
不動産販促資料生成ワークフロー メインスクリプト

使用方法:
  python workflow.py             # 全ステップ実行
  python workflow.py --step 1    # Step 1のみ
  python workflow.py --step 1 2  # Step 1, 2のみ
  python workflow.py --step 3 4  # Step 3, 4のみ（intermediatesが揃っている場合）

ステップ:
  1: PDFからテキスト抽出 → outputs/intermediates/text/
  2: 画像加工（Gemini）  → outputs/intermediates/images/
  3: PowerPoint生成      → outputs/*.pptx
  4: PDF変換             → outputs/*.pdf
"""

import argparse
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR / "src"))

from step1_extract_text import run as run_step1
from step2_process_images import run as run_step2
from step3_generate_pptx import run as run_step3
from step4_export_pdf import run as run_step4


def main():
    parser = argparse.ArgumentParser(description="不動産販促資料生成ワークフロー")
    parser.add_argument(
        "--step",
        nargs="*",
        type=int,
        choices=[1, 2, 3, 4],
        help="実行するステップ番号（省略時は全ステップ実行）",
    )
    args = parser.parse_args()

    steps = set(args.step) if args.step else {1, 2, 3, 4}

    pptx_path = None

    if 1 in steps:
        print()
        run_step1()

    if 2 in steps:
        print()
        run_step2()

    if 3 in steps:
        print()
        pptx_path = run_step3()

    if 4 in steps:
        print()
        run_step4(pptx_path)

    print("\n全ステップ完了")


if __name__ == "__main__":
    main()
