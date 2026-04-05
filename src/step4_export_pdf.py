"""
Step 4: PowerPointをPDFに変換する
Mac環境前提: LibreOffice CLI を使用

前提: LibreOfficeがインストールされていること
  brew install --cask libreoffice
"""

import subprocess
import shutil
from pathlib import Path
import yaml

SCRIPT_DIR = Path(__file__).parent
REPO_ROOT = SCRIPT_DIR.parent

LIBREOFFICE_CANDIDATES = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/local/bin/libreoffice",
    "/opt/homebrew/bin/libreoffice",
]


def load_config() -> dict:
    with open(REPO_ROOT / "config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def find_libreoffice() -> str:
    for path in LIBREOFFICE_CANDIDATES:
        if Path(path).exists():
            return path
    # PATHから探す
    found = shutil.which("libreoffice") or shutil.which("soffice")
    if found:
        return found
    raise FileNotFoundError(
        "LibreOfficeが見つかりません。以下でインストールしてください:\n"
        "  brew install --cask libreoffice"
    )


def convert_to_pdf(pptx_path: Path, output_dir: Path) -> Path:
    """LibreOffice CLIを使ってPPTXをPDFに変換する"""
    soffice = find_libreoffice()
    print(f"LibreOffice: {soffice}")

    cmd = [
        soffice,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(output_dir),
        str(pptx_path),
    ]
    print(f"実行: {' '.join(cmd)}")

    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

    if result.returncode != 0:
        raise RuntimeError(
            f"PDF変換失敗 (returncode={result.returncode})\n"
            f"stdout: {result.stdout}\n"
            f"stderr: {result.stderr}"
        )

    print(result.stdout)
    pdf_path = output_dir / (pptx_path.stem + ".pdf")
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF出力ファイルが見つかりません: {pdf_path}")

    return pdf_path


def run(pptx_path: Path = None) -> Path:
    config = load_config()
    output_dir = REPO_ROOT / config["paths"]["outputs"]

    if pptx_path is None:
        slug = config["property"]["slug"]
        pptx_path = output_dir / f"{slug}.pptx"

    if not pptx_path.exists():
        raise FileNotFoundError(f"PPTXが見つかりません: {pptx_path}")

    print("=== Step 4: PDF変換 ===")
    pdf_path = convert_to_pdf(pptx_path, output_dir)
    print(f"PDF保存: {pdf_path}")
    print("=== Step 4 完了 ===")
    return pdf_path


if __name__ == "__main__":
    run()
