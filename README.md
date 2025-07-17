# 📄 docx-xlsx-to-pdf

Convert `.docx` and `.xlsx` files into **PDF/A** format in batch, preserving folder structure.  
This tool uses **LibreOffice CLI** for true headless conversion and is suitable for document archiving and AI processing.

---

## 🌐 Available Languages / 사용 가능한 언어 / 利用可能な言語 / 可用语言

- [🇺🇸 English](#-english)
- [🇰🇷 한국어](#-한국어)
- [🇯🇵 日本語](#-日本語)
- [🇨🇳 中文](#-中文)

---

## 🇺🇸 English

### ✅ Features
- 🧠 Converts to **PDF/A-1** (ISO archive format)
- 📁 Preserves original folder structure
- ⚡ Headless via `soffice` (LibreOffice CLI)
- 📊 Progress bar with `tqdm`
- ✅ Supports `.docx` and `.xlsx`

### 🔧 Requirements

```bash
pip install -r requirements.txt
```

Install LibreOffice:

- **macOS**:
  ```bash
  brew install --cask libreoffice
  export PATH="/Applications/LibreOffice.app/Contents/MacOS:$PATH"
  ```
- **Windows**:
  - Download from https://www.libreoffice.org/download
  - Add LibreOffice path (e.g., `C:\Program Files\LibreOffice\program`) to system PATH
- **Linux (Debian/Ubuntu)**:
  ```bash
  sudo apt install libreoffice
  ```

### 🚀 Usage

```bash
python main.py --input ./docs --output ./pdf_output
```

---

## 🇰🇷 한국어

### ✅ 기능
- 🧠 PDF/A-1 형식으로 변환
- 📁 원본 폴더 구조 유지
- ⚡ LibreOffice CLI(`soffice`) 기반 Headless 실행
- 📊 tqdm 진행률 표시
- ✅ `.docx`, `.xlsx` 지원

### 🔧 설치 방법

```bash
pip install -r requirements.txt
```

LibreOffice 설치:

- **macOS**:
  ```bash
  brew install --cask libreoffice
  export PATH="/Applications/LibreOffice.app/Contents/MacOS:$PATH"
  ```
- **Windows**:
  - https://www.libreoffice.org/download 에서 다운로드
  - `C:\Program Files\LibreOffice\program` 경로를 시스템 PATH에 추가
- **Linux (Debian/Ubuntu)**:
  ```bash
  sudo apt install libreoffice
  ```

### 🚀 사용 방법

```bash
python main.py --input ./docs --output ./pdf_output
```

---

## 🇯🇵 日本語

### ✅ 機能
- 🧠 PDF/A-1（ISOアーカイブ形式）に変換
- 📁 元のフォルダー構造を保持
- ⚡ LibreOffice CLI（`soffice`）でヘッドレス実行
- 📊 `tqdm` による進捗バー
- ✅ `.docx`、`.xlsx` に対応

### 🔧 インストール手順

```bash
pip install -r requirements.txt
```

LibreOffice のインストール:

- **macOS**:
  ```bash
  brew install --cask libreoffice
  export PATH="/Applications/LibreOffice.app/Contents/MacOS:$PATH"
  ```
- **Windows**:
  - https://www.libreoffice.org/download からダウンロード
  - `C:\Program Files\LibreOffice\program` を PATH に追加
- **Linux (Debian/Ubuntu)**:
  ```bash
  sudo apt install libreoffice
  ```

### 🚀 使い方

```bash
python main.py --input ./docs --output ./pdf_output
```

---

## 🇨🇳 中文

### ✅ 功能
- 🧠 转换为 PDF/A-1（ISO 归档格式）
- 📁 保留原始文件夹结构
- ⚡ 使用 LibreOffice CLI（`soffice`）真正无头执行
- 📊 `tqdm` 显示进度条
- ✅ 支持 `.docx` 和 `.xlsx`

### 🔧 安装方法

```bash
pip install -r requirements.txt
```

安装 LibreOffice：

- **macOS**:
  ```bash
  brew install --cask libreoffice
  export PATH="/Applications/LibreOffice.app/Contents/MacOS:$PATH"
  ```
- **Windows**:
  - 从 https://www.libreoffice.org/download 下载并安装
  - 将 `C:\Program Files\LibreOffice\program` 添加到系统 PATH
- **Linux (Debian/Ubuntu)**:
  ```bash
  sudo apt install libreoffice
  ```

### 🚀 使用方法

```bash
python main.py --input ./docs --output ./pdf_output
```
