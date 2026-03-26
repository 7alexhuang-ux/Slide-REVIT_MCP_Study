#!/usr/bin/env python3
"""
Project Showcase — PPTX → PDF 轉換器

支援兩種轉換方式（依可用性自動選擇）：
1. PowerPoint COM 自動化（Windows + PowerPoint 已安裝）
2. LibreOffice headless 模式（跨平台）

用法:
    python convert_to_pdf.py input.pptx [--output output.pdf]
"""

import argparse
import subprocess
import sys
from pathlib import Path


def convert_via_powerpoint(pptx_path: Path, pdf_path: Path) -> bool:
    """透過 PowerPoint COM 自動化轉換（僅 Windows）"""
    try:
        import comtypes.client
    except ImportError:
        try:
            import win32com.client
            return _convert_win32com(pptx_path, pdf_path)
        except ImportError:
            return False

    try:
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1

        pptx_abs = str(pptx_path.resolve())
        pdf_abs = str(pdf_path.resolve())

        presentation = powerpoint.Presentations.Open(pptx_abs, WithWindow=False)
        # 32 = ppSaveAsPDF
        presentation.SaveAs(pdf_abs, 32)
        presentation.Close()
        powerpoint.Quit()
        return True
    except Exception as e:
        print(f'PowerPoint COM 轉換失敗: {e}', file=sys.stderr)
        return False


def _convert_win32com(pptx_path: Path, pdf_path: Path) -> bool:
    """透過 win32com 轉換"""
    try:
        import win32com.client

        powerpoint = win32com.client.Dispatch("PowerPoint.Application")

        pptx_abs = str(pptx_path.resolve())
        pdf_abs = str(pdf_path.resolve())

        presentation = powerpoint.Presentations.Open(pptx_abs, WithWindow=False)
        presentation.SaveAs(pdf_abs, 32)  # 32 = ppSaveAsPDF
        presentation.Close()
        powerpoint.Quit()
        return True
    except Exception as e:
        print(f'win32com 轉換失敗: {e}', file=sys.stderr)
        return False


def convert_via_libreoffice(pptx_path: Path, pdf_path: Path) -> bool:
    """透過 LibreOffice headless 模式轉換"""
    # 嘗試常見的 LibreOffice 路徑
    soffice_paths = [
        'soffice',
        'libreoffice',
        '/usr/bin/soffice',
        '/usr/bin/libreoffice',
        'C:/Program Files/LibreOffice/program/soffice.exe',
        'C:/Program Files (x86)/LibreOffice/program/soffice.exe',
    ]

    soffice = None
    for path in soffice_paths:
        try:
            result = subprocess.run(
                [path, '--version'],
                capture_output=True, timeout=10,
            )
            if result.returncode == 0:
                soffice = path
                break
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue

    if not soffice:
        return False

    try:
        output_dir = str(pdf_path.parent.resolve())
        subprocess.run(
            [
                soffice,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', output_dir,
                str(pptx_path.resolve()),
            ],
            capture_output=True,
            timeout=120,
            check=True,
        )

        # LibreOffice 會用原檔名 + .pdf，需要確認/重新命名
        auto_name = pptx_path.with_suffix('.pdf')
        if auto_name.name != pdf_path.name and auto_name.exists():
            auto_name.rename(pdf_path)

        return pdf_path.exists()
    except Exception as e:
        print(f'LibreOffice 轉換失敗: {e}', file=sys.stderr)
        return False


def convert_to_pdf(pptx_path: str, output_path: str = None):
    """主函式：PPTX → PDF"""
    pptx_file = Path(pptx_path)
    if not pptx_file.exists():
        print(f'錯誤：找不到檔案 {pptx_path}', file=sys.stderr)
        sys.exit(1)

    if output_path:
        pdf_file = Path(output_path)
    else:
        pdf_file = pptx_file.with_suffix('.pdf')

    print(f'正在轉換: {pptx_file.name} → {pdf_file.name}')

    # 方法一：PowerPoint COM（Windows）
    if sys.platform == 'win32':
        print('  嘗試 PowerPoint COM 自動化...')
        if convert_via_powerpoint(pptx_file, pdf_file):
            print(f'✅ 已生成 PDF：{pdf_file}')
            return str(pdf_file)
        print('  PowerPoint 不可用，嘗試 LibreOffice...')

    # 方法二：LibreOffice
    print('  嘗試 LibreOffice headless...')
    if convert_via_libreoffice(pptx_file, pdf_file):
        print(f'✅ 已生成 PDF：{pdf_file}')
        return str(pdf_file)

    # 都失敗
    print('❌ PDF 轉換失敗。請確認已安裝以下其一：', file=sys.stderr)
    print('   - Microsoft PowerPoint（Windows）', file=sys.stderr)
    print('   - LibreOffice（跨平台，https://www.libreoffice.org/）', file=sys.stderr)
    sys.exit(1)


def main():
    parser = argparse.ArgumentParser(
        description='Project Showcase — PPTX → PDF 轉換',
    )
    parser.add_argument('input', help='PPTX 檔案路徑')
    parser.add_argument('--output', '-o', help='輸出 PDF 路徑')
    args = parser.parse_args()
    convert_to_pdf(args.input, args.output)


if __name__ == '__main__':
    main()
