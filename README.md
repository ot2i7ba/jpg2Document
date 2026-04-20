# jpg2Document
**jpg2Document** is a Python script that generates a Microsoft Word document from a template by inserting a table of compressed, landscape-oriented images in place of a specified placeholder. It is especially useful for creating reports, presentations, or documentation that require structured visual content.

> [!NOTE]
> Only landscape-oriented images (width > height) are processed. Portrait and square images are skipped automatically.

> [!WARNING]
> This script is in active development. While it has been tested for core functionality, complex Word templates (e.g., heavily formatted placeholders or nested tables) may require additional testing.

## Table of Contents
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Command-Line Arguments](#command-line-arguments)
- [Examples](#examples)
- [Releases](#releases)
- [Changes](#changes)
- [License](#license)

## Features
- **Template-Based Document Generation:** Replaces a plain-text placeholder in a `.docx` template with a two-column image table.
- **Image Compression and Scaling:** Downscales images to a configurable maximum pixel width and compresses them as JPEG before insertion.
- **Landscape Filtering:** Only processes images where width > height. Portrait, square, and unreadable files are skipped and counted in the summary.
- **Parallel Processing:** Compresses images concurrently using multiple CPU cores (`--workers`), configurable or auto-detected.
- **Configurable Resampling:** Choose the downscaling filter (`lanczos`, `bicubic`, `bilinear`, `nearest`) via `--resample`.
- **Overwrite Protection:** The output file is never silently overwritten — use `--force` to permit it.
- **Robust Error Handling:** All configuration errors are collected and reported together before processing begins. Unexpected image errors are caught per-file without aborting the run.
- **Progress Spinner:** Displays a console spinner during processing; automatically suppressed when `--verbose` is active.
- **Document Metadata:** Sets author and comments properties on the generated document.

## Requirements
- **Python:** 3.7 or higher
- **Python Packages:**
  - [Pillow](https://python-pillow.org/) `==9.3.0`
  - [python-docx](https://python-docx.readthedocs.io/en/latest/) `==0.8.11`

## Installation

1. **Clone the repository**
   ```sh
   git clone https://github.com/ot2i7ba/jpg2Document.git
   cd jpg2Document
   ```

2. **Install dependencies**
   ```sh
   pip install -r requirements.txt
   ```

## Usage

1. **Prepare your template**
   Create a `.docx` file containing the placeholder text `<<jpg2Document>>` at the position where the image table should be inserted. The placeholder must be plain, unstyled text within a single paragraph.

2. **Place your images**
   Put the landscape images (`.jpg`, `.jpeg`, `.png`) you want to include in the working directory or specify a path with `--input_dir`.

3. **Run the script**
   ```sh
   python jpg2Document.py
   ```

## Command-Line Arguments

### input / output
| Argument | Default | Description |
|---|---|---|
| `--template` | `jpg2Document.docx` | Path to the Word template |
| `--output` | `pictures.docx` | Path for the output file |
| `--input_dir` | current directory | Directory containing source images |
| `--extensions` | `.jpg,.jpeg,.png` | Comma-separated list of file extensions to process |
| `--placeholder` | `<<jpg2Document>>` | Placeholder text to replace in the template |
| `--force` | off | Overwrite output file if it already exists |

### image options
| Argument | Default | Description |
|---|---|---|
| `--image_width` | `9.2` | Width (cm) per image in the table |
| `--gap_width` | `0.05` | Gap (cm) between the two image columns |
| `--max_px` | `1200` | Maximum image width in pixels before downscaling |
| `--jpeg_quality` | `80` | JPEG compression quality (1–95) |
| `--resample` | `lanczos` | Downscaling filter: `lanczos`, `bicubic`, `bilinear`, `nearest` |

### document metadata
| Argument | Default | Description |
|---|---|---|
| `--doc_author` | `jpg2Document` | Author property of the output document |
| `--doc_comments` | `jpg2Document by ot2i7ba` | Comments property of the output document |

### runtime
| Argument | Default | Description |
|---|---|---|
| `--workers` | CPU count | Parallel worker processes for image compression |
| `--clear` | off | Clear the console before saving output |
| `--verbose` | off | Enable debug logging (also suppresses spinner) |

## Examples

### Default run
```sh
python jpg2Document.py
```
Reads `jpg2Document.docx` as template, processes all landscape `.jpg`/`.jpeg`/`.png` images in the current directory, and writes `pictures.docx`.

### Custom template and output
```sh
python jpg2Document.py \
  --template report_template.docx \
  --output annual_report.docx \
  --placeholder "<<AnnualImages>>" \
  --input_dir ./photos \
  --image_width 8.0 \
  --gap_width 0.1 \
  --max_px 1000 \
  --jpeg_quality 75 \
  --doc_author "Jane Doe" \
  --force
```

### Fast run with all cores, bilinear resampling
```sh
python jpg2Document.py --resample bilinear --verbose
```

## Releases
A compiled version (no Python required) is available in the **[Releases](https://github.com/ot2i7ba/jpg2Document/releases)** section on GitHub.

___

## Changes

## v0.1.0
- **Critical fix:** Implemented the missing `process_images()` function — the script was not runnable in v0.0.5.
- **Parallel image compression:** Added `ProcessPoolExecutor` support via `--workers` (defaults to CPU count).
- **Config validation:** All parameters are validated before processing begins; all errors are reported together.
- **Overwrite protection:** Output file is never silently overwritten; `--force` flag required.
- **Resample filter:** Added `--resample` option (`lanczos`, `bicubic`, `bilinear`, `nearest`).
- **Spinner integrated:** Progress spinner is now active during compression; suppressed in `--verbose` mode.
- **`--clear` is now opt-in:** Previously the console was cleared by default; now requires explicit `--clear`.
- **Run-level placeholder replacement:** Preserves paragraph character formatting when replacing the placeholder; falls back gracefully if the placeholder spans multiple runs.
- **Robust border handling:** `_remove_table_borders` now replaces any existing `tblBorders` element instead of appending a duplicate.
- **Broad exception handling in workers:** Image processing errors (including `MemoryError`, `DecompressionBombError`) are caught per-file and reported as skipped, never abort the run.
- **`freeze_support()` placement:** Moved to `if __name__ == "__main__"` guard for correct frozen-executable support.
- **Dead code removed:** `DEFAULT_INPUT_DIR`, unused `UnidentifiedImageError` import.

## v0.0.2
- Refactored main process into `Jpg2Dokument` class.
- Spinner converted to context manager.
- Replaced `os.path` with `pathlib`.
- Added `tempfile.TemporaryDirectory()` for automatic cleanup.

## v0.0.1
- Initial release.

___

## License
This project is licensed under the **[MIT License](https://github.com/ot2i7ba/jpg2Document/blob/main/LICENSE)**.

## Contributing
Contributions are welcome. Please fork the repository and submit a pull request for review.

## Disclaimer
This project is provided without warranties. See the license for terms of use and limitations of liability.

# Conclusion
This script has been tailored to fit my specific professional needs, and while it may seem like a small tool, it has a significant impact on my workflow. jpg2Document is a valuable tool for automating the creation of Word documents that incorporate multiple images in a structured and visually appealing manner. By handling image compression, scaling, and placement, it simplifies the process of compiling images into reports or presentations. Greetings to my dear colleagues who avoid scripts like the plague and think that consoles and Bash are some sort of dark magic – the [compiled](https://github.com/ot2i7ba/jpg2Document/releases) version will spare you the console kung-fu and hopefully be a helpful tool for you as well. 😉