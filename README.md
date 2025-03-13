# DOCX Splitter & Merger

A lightweight CLI Python script used for seamlessly splitting and merging `.docx` files while preserving media and formatting.

## Features
- **Split** a `.docx` file at page breaks.
- **Merge** multiple `.docx` files into one.
- **Preserves** tables, images, and text styles.
- **Handles edge cases** (empty files, non-`.docx` inputs, etc.).
- **Simple command-line interface (CLI).**

## Installation
Requires **Python 3.x** and `python-docx`. Install with:
```sh
pip install python-docx
```

## Usage
### Split a DOCX File
```sh
python docx_tool.py split input.docx
```
Creates `split_part_1.docx`, `split_part_2.docx`, etc.

### Merge Multiple DOCX Files
```sh
python docx_tool.py merge file1.docx file2.docx --output merged.docx
```
Creates `merged.docx` with page breaks between documents.

## Edge Cases Handled
- Skips empty files.
- Ignores non-`.docx` files when merging.
- Ensures formatting, tables, and images are retained.

## Quick Start
1. Clone the repo:
   ```sh
   git clone https://github.com/tobias-sava/docx-split-merge.git && cd docx-split-merge
   ```
2. Install dependencies:
   ```sh
   pip install python-docx
   ```
3. Run the script:
   ```sh
   python docx_tool.py split input.docx  # To split
   python docx_tool.py merge file1.docx file2.docx --output merged.docx  # To merge
   ```

## License
This project is under the **MIT License**. Use, modify, and distribute freely.

