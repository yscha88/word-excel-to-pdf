import locale
from tqdm import tqdm
import subprocess
from pathlib import Path
import platform
import shutil

def get_system_language():
    lang = locale.getdefaultlocale()[0]
    if lang.startswith("ko"):
        return "ko"
    elif lang.startswith("ja"):
        return "ja"
    elif lang.startswith("zh"):
        return "zh"
    else:
        return "en"

LANG = get_system_language()

def get_soffice_command():
    cmd = shutil.which("soffice")
    if cmd:
        return cmd

    system = platform.system()
    if system == "Darwin":  # macOS
        return "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    elif system == "Windows":
        return "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
    elif system == "Linux":
        possible_paths = [
            "/usr/bin/soffice",
            "/usr/lib/libreoffice/program/soffice",
            "/snap/bin/soffice"
        ]
        for path in possible_paths:
            if Path(path).exists():
                return path
        return "/usr/bin/soffice"
    else:
        raise EnvironmentError("Unsupported OS: please install LibreOffice and add it to PATH.")

def convert_to_pdf(input_path: Path, input_root: Path, output_root: Path):
    ext = input_path.suffix.lower()
    if ext not in [".docx", ".xlsx"]:
        msg = {
            "en": "[SKIP] Unsupported file:",
            "ko": "[SKIP] 지원하지 않는 파일:",
            "ja": "[SKIP] 未対応のファイル:",
            "zh": "[SKIP] 不支持的文件："
        }[LANG]
        print(f"{msg} {input_path}")
        return

    rel_path = input_path.relative_to(input_root)
    output_path = output_root / rel_path.with_suffix(".pdf")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # LibreOffice CLI conversion using soffice
    try:
        subprocess.run([
            get_soffice_command(),
            "--headless",
            "--convert-to", "pdf:writer_pdf_Export:SelectPdfVersion=1",
            "--outdir", str(output_path.parent),
            str(input_path.resolve())
        ], check=True)
        msg = {
            "en": "[OK] Converted:",
            "ko": "[OK] 변환 완료:",
            "ja": "[OK] 変換完了:",
            "zh": "[OK] 转换完成："
        }[LANG]
        print(f"{msg} {output_path}")
    except subprocess.CalledProcessError as e:
        msg = {
            "en": "[ERROR] LibreOffice conversion failed:",
            "ko": "[ERROR] LibreOffice 변환 실패:",
            "ja": "[ERROR] LibreOffice の変換に失敗しました:",
            "zh": "[ERROR] LibreOffice 转换失败："
        }[LANG]
        print(f"{msg} {input_path} - {e}")

def main(input_dir="./docs", output_dir="./pdf_output"):
    input_root = Path(input_dir)
    output_root = Path(output_dir)

    if not input_root.exists():
        msg = {
            "en": "[ERROR] Input folder does not exist:",
            "ko": "[ERROR] 입력 폴더가 존재하지 않습니다:",
            "ja": "[ERROR] 入力フォルダが存在しません:",
            "zh": "[ERROR] 输入文件夹不存在："
        }[LANG]
        print(f"{msg} {input_root}")
        return

    files = sorted(list(input_root.rglob("*.docx")) + list(input_root.rglob("*.xlsx")))
    if not files:
        msg = {
            "en": "[INFO] No files to convert.",
            "ko": "[INFO] 변환할 파일이 없습니다.",
            "ja": "[INFO] 変換対象のファイルがありません。",
            "zh": "[INFO] 没有可转换的文件。"
        }[LANG]
        print(msg)
        return

    success, failure = 0, 0
    desc_map = {
        "en": "📄 Converting",
        "ko": "📄 변환 진행 중",
        "ja": "📄 変換中",
        "zh": "📄 正在转换"
    }
    unit_map = {
        "en": "file",
        "ko": "파일",
        "ja": "ファイル",
        "zh": "个文件"
    }

    for file in tqdm(files, desc=desc_map[LANG], unit=unit_map[LANG]):
        try:
            convert_to_pdf(file, input_root, output_root)
            success += 1
        except Exception as e:
            err_msg = {
                "en": "[ERROR] Conversion failed:",
                "ko": "[ERROR] 변환 실패:",
                "ja": "[ERROR] 変換に失敗しました:",
                "zh": "[ERROR] 转换失败："
            }[LANG]
            print(f"{err_msg} {file} - {e}")
            failure += 1

    msg = {
        "en": "✅ Converted:",
        "ko": "✅ 변환 완료:",
        "ja": "✅ 変換完了：",
        "zh": "✅ 转换完成："
    }[LANG]
    print(f"\n{msg} {success}")
    if failure:
        msg = {
            "en": f"❌ Failed: {failure} file(s). See above logs.",
            "ko": f"❌ 실패: {failure}건 (위 로그 참고)",
            "ja": f"❌ 失敗: {failure}件（上記ログを参照）",
            "zh": f"❌ 失败：{failure}个（请参考以上日志）"
        }[LANG]
        print(msg)

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Convert .docx/.xlsx to PDF/A")
    parser.add_argument("--input", default="./docs", help="Input directory")
    parser.add_argument("--output", default="./pdf_output", help="Output directory")
    args = parser.parse_args()
    main(args.input, args.output)

    # python main.py --input ./docs --output ./pdf_output