# jpg2Document.py
# Copyright (c) 2024 ot2i7ba
# https://github.com/ot2i7ba/
# This code is licensed under the MIT License (see LICENSE for details).

"""
jpg2Document.py v0.0.2 - 2025-02-07
Generate a Word document from a template by inserting a table of compressed,
landscape-oriented images in place of a specified placeholder. If the placeholder
is not found, the script aborts immediately (no images are processed).
"""

# Python standard library imports
import sys
import itertools
import time
import threading
import subprocess
import shutil
import argparse
import tempfile
from math import ceil
from typing import List, Tuple
from pathlib import Path

# Third-party library imports
from PIL import Image, UnidentifiedImageError
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm
from docx.table import Table
from docx.text.paragraph import Paragraph  # Extended typing

# Default values
DEFAULT_TEMPLATE = "jpg2Document.docx"            # Default Word template name
DEFAULT_OUTPUT = "pictures.docx"                  # Default output (generated) document name
DEFAULT_PLACEHOLDER = "<<jpg2Document>>"          # Placeholder text to be replaced in the template
DEFAULT_IMAGE_WIDTH_CM = 9.2                      # Width in centimeters for each inserted image
DEFAULT_GAP_WIDTH_CM = 0.05                       # Gap in centimeters between image columns
DEFAULT_MAX_IMAGE_WIDTH_PX = 1200                 # Maximum width in pixels before scaling images down
DEFAULT_JPEG_QUALITY = 80                         # JPEG compression quality (0-100)
DEFAULT_DOC_AUTHOR = "jpg2Document"               # Document property: author
DEFAULT_DOC_COMMENTS = "jpg2Document by ot2i7ba"  # Document property: comments
DEFAULT_INPUT_DIR = Path.cwd()                    # Default current directory
DEFAULT_EXTENSIONS = ".jpg,.jpeg,.png"            # Allowed image file extensions

def clear_console() -> None:
    """Clear the console screen using subprocess.call."""
    command = 'cls' if sys.platform.startswith('win') else 'clear'
    subprocess.call(command, shell=True)

def print_blank_line() -> None:
    """Print a blank line."""
    print()

class Spinner:
    """
    Object-oriented spinner for displaying a progress indicator.
    Usable with start() and stop(), and as a context manager.
    """
    def __init__(self, task_name: str) -> None:
        self.task_name = task_name
        self._stop_event = threading.Event()
        self._lock = threading.Lock()
        self._spinner_cycle = itertools.cycle(['-', '/', '-', '\\'])
        self._thread = threading.Thread(target=self._spin, daemon=True)

    def _spin(self) -> None:
        while not self._stop_event.is_set():
            with self._lock:
                sys.stdout.write(f'\rProcessing {self.task_name} ... {next(self._spinner_cycle)}')
                sys.stdout.flush()
            time.sleep(0.1)
        with self._lock:
            sys.stdout.write('\r')
            sys.stdout.flush()

    def start(self) -> None:
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        self._thread.join()

    def __enter__(self) -> "Spinner":
        self.start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.stop()

def remove_table_borders(table: Table) -> None:
    """Remove all visible borders from the given docx 'table'."""
    tbl = table._tbl
    if tbl.tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.append(tblPr)
    else:
        tblPr = tbl.tblPr

    tblBorders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_el = OxmlElement(f'w:{edge}')
        edge_el.set(qn('w:val'), 'none')
        tblBorders.append(edge_el)
    tblPr.append(tblBorders)

def compress_and_scale_image(
    input_path: str,
    output_path: str,
    max_width_px: int,
    quality: int
) -> bool:
    """
    Attempt to open 'input_path' as an image, resize it to 'max_width_px'
    if it is wider, and save it as a JPEG to 'output_path' with the specified quality.

    Returns:
        True if the image was successfully opened and saved, False otherwise.
    """
    try:
        with Image.open(input_path) as img:
            # Only accept landscape images (width > height).
            if img.width <= img.height:
                return False

            if img.width > max_width_px:
                ratio = max_width_px / float(img.width)
                new_height = int(img.height * ratio)
                img = img.resize((max_width_px, new_height), Image.LANCZOS)

            img.save(output_path, 'JPEG', optimize=True, quality=quality)
        return True

    except (OSError, UnidentifiedImageError) as err:
        print(f"\nSkipping '{input_path}': could not open or process the image ({err}).")
        return False

def insert_table_after_paragraph(
    document: Document,
    paragraph: Paragraph,
    rows: int,
    cols: int
) -> Table:
    """
    Create a new table with (rows x cols) cells, then insert it immediately
    after the given paragraph in the docx document.

    Returns:
        A docx.table.Table object representing the newly inserted table.
    """
    tmp_table = document.add_table(rows=rows, cols=cols)
    tbl_element = tmp_table._tbl
    tmp_table._element.getparent().remove(tmp_table._element)
    paragraph._p.addnext(tbl_element)

    return Table(tbl_element, paragraph._parent)

def create_table_with_images(
    document: Document,
    paragraph: Paragraph,
    image_paths: List[str],
    image_width_cm: float,
    gap_width_cm: float
) -> None:
    """
    Build a 3-column table of images (left image, narrow gap, right image),
    inserting the table immediately after the given paragraph. Each pair of images
    is followed by an empty row.
    """
    total_images = len(image_paths)
    pairs = ceil(total_images / 2)
    rows_needed = pairs * 2
    cols = 3

    table = insert_table_after_paragraph(document, paragraph, rows_needed, cols)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    remove_table_borders(table)

    for row in table.rows:
        row.cells[0].width = Cm(image_width_cm)
        row.cells[1].width = Cm(gap_width_cm)
        row.cells[2].width = Cm(image_width_cm)

    idx_image = 0
    for pair_idx in range(pairs):
        row_images = table.rows[pair_idx * 2]
        # The following row remains empty as a separator.
        _ = table.rows[pair_idx * 2 + 1]

        if idx_image < total_images:
            cell_left = row_images.cells[0]
            p_left = cell_left.paragraphs[0]
            p_left.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_left.add_run().add_picture(image_paths[idx_image], width=Cm(image_width_cm))
            idx_image += 1

        if idx_image < total_images:
            cell_right = row_images.cells[2]
            p_right = cell_right.paragraphs[0]
            p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p_right.add_run().add_picture(image_paths[idx_image], width=Cm(image_width_cm))
            idx_image += 1

def get_image_files(input_dir: Path, valid_ext: Tuple[str, ...]) -> List[Path]:
    """
    Return a sorted list of image file Paths from the specified directory that match the given extensions.
    """
    # Search the directory (only files in the current directory)
    files = [p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() in valid_ext]
    return sorted(files, key=lambda p: p.name.lower())

def process_images(
    input_dir: Path,
    compressed_dir: Path,
    valid_ext: Tuple[str, ...],
    max_width_px: int,
    jpeg_quality: int
) -> List[str]:
    """
    Process all images in the given input directory that have valid extensions.
    Compress and scale images into the compressed directory.
    
    Returns:
        A list of file paths (as strings) of the successfully processed images.
    """
    image_files = get_image_files(input_dir, valid_ext)
    compressed_paths: List[str] = []
    for i, img_path in enumerate(image_files):
        out_path = compressed_dir / f"compressed_{i}.jpg"
        if compress_and_scale_image(str(img_path), str(out_path), max_width_px, jpeg_quality):
            compressed_paths.append(str(out_path))
    return compressed_paths

def parse_arguments() -> argparse.Namespace:
    """
    Parse optional command-line arguments. If none are provided, default values are used.
    """
    parser = argparse.ArgumentParser(
        description="Insert compressed landscape images into a Word template at a specified placeholder.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )

    parser.add_argument("--template", type=str, default=None,
                        help="Path to the Word template (.docx).")
    parser.add_argument("--output", type=str, default=None,
                        help="Path for the output Word file.")
    parser.add_argument("--placeholder", type=str, default=None,
                        help="Placeholder text in the template to be replaced.")
    parser.add_argument("--input_dir", type=str, default=str(DEFAULT_INPUT_DIR),
                        help="Directory to search for image files.")
    parser.add_argument("--extensions", type=str, default=DEFAULT_EXTENSIONS,
                        help="Comma-separated list of valid image file extensions.")
    parser.add_argument("--image_width", type=float, default=None,
                        help="Width in cm for each inserted image.")
    parser.add_argument("--gap_width", type=float, default=None,
                        help="Width in cm for the gap between images.")
    parser.add_argument("--max_px", type=int, default=None,
                        help="Maximum width in pixels before image is scaled down.")
    parser.add_argument("--jpeg_quality", type=int, default=None,
                        help="JPEG compression quality (0-100).")
    parser.add_argument("--doc_author", type=str, default=None,
                        help="Value for the document's author property.")
    parser.add_argument("--doc_comments", type=str, default=None,
                        help="Value for the document's comments property.")

    return parser.parse_args()

def validate_parameters(args: argparse.Namespace) -> Tuple[Tuple[str, ...], float, float, int, int]:
    """
    Validate and normalize input parameters.
    
    Returns:
        A tuple containing:
          - valid_extensions: Tuple[str, ...]
          - image_width_cm: float
          - gap_width_cm: float
          - max_image_width_px: int
          - jpeg_quality: int
    """
    # Validate JPEG quality
    quality = args.jpeg_quality if args.jpeg_quality is not None else DEFAULT_JPEG_QUALITY
    if not (0 <= quality <= 100):
        print("jpeg_quality must be between 0 and 100.")
        sys.exit(1)
    
    # Validate width values
    img_width = args.image_width if args.image_width is not None else DEFAULT_IMAGE_WIDTH_CM
    gap_width = args.gap_width if args.gap_width is not None else DEFAULT_GAP_WIDTH_CM
    if img_width <= 0 or gap_width < 0:
        print("image_width must be > 0 and gap_width must be >= 0.")
        sys.exit(1)
    
    # Validate max_px
    max_px = args.max_px if args.max_px is not None else DEFAULT_MAX_IMAGE_WIDTH_PX
    if max_px <= 0:
        print("max_px must be a positive integer.")
        sys.exit(1)
    
    # Validate the input directory
    input_path = Path(args.input_dir)
    if not input_path.is_dir():
        print(f"Input directory '{input_path}' does not exist or is not a directory.")
        sys.exit(1)
    
    # Process file extensions (e.g., ".jpg,.jpeg,.png")
    ext_list = [ext.strip().lower() for ext in args.extensions.split(',') if ext.strip()]
    if not ext_list:
        print("At least one valid file extension must be provided.")
        sys.exit(1)
    
    return (tuple(ext_list), img_width, gap_width, max_px, quality)

class jpg2Document:
    """
    Encapsulates the entire process:
      - Reading and validating parameters,
      - Opening the template,
      - Processing and compressing images,
      - Inserting images into the template,
      - Setting document properties and saving.
    """
    def __init__(self, args: argparse.Namespace) -> None:
        self.args = args
        self.valid_ext, self.image_width_cm, self.gap_width_cm, self.max_image_width_px, self.jpeg_quality = validate_parameters(args)
        self.docx_template = args.template or DEFAULT_TEMPLATE
        self.docx_output = args.output or DEFAULT_OUTPUT
        self.placeholder_text = args.placeholder or DEFAULT_PLACEHOLDER
        self.doc_author = args.doc_author or DEFAULT_DOC_AUTHOR
        self.doc_comments = args.doc_comments or DEFAULT_DOC_COMMENTS
        self.input_dir = Path(args.input_dir)

    def run(self) -> None:
        # Check if the template exists
        template_path = Path(self.docx_template)
        if not template_path.exists():
            print(f"Template '{template_path}' was not found. Aborting.")
            sys.exit(1)

        try:
            document = Document(str(template_path))
        except Exception as err:
            print(f"Failed to open template '{template_path}': {err}")
            sys.exit(1)

        # Check if the placeholder is present.
        if not any(self.placeholder_text in paragraph.text for paragraph in document.paragraphs):
            print(f"Placeholder '{self.placeholder_text}' was not found. Aborting.")
            sys.exit(1)

        # Retrieve image files from the input directory
        image_files = get_image_files(self.input_dir, self.valid_ext)
        if not image_files:
            print("No image files found in the input directory. Aborting.")
            sys.exit(1)

        # Use a temporary directory to store the compressed images.
        with tempfile.TemporaryDirectory() as temp_dir:
            compressed_dir = Path(temp_dir)
            with Spinner("images") as spinner:
                compressed_paths = process_images(self.input_dir, compressed_dir, self.valid_ext, self.max_image_width_px, self.jpeg_quality)
            print()  # Line break after the spinner

            if not compressed_paths:
                print("No valid landscape images could be processed. Aborting.")
                sys.exit(1)

            # Replace the placeholder in the document with the table of images.
            for paragraph in document.paragraphs:
                if self.placeholder_text in paragraph.text:
                    paragraph.text = paragraph.text.replace(self.placeholder_text, "")
                    create_table_with_images(document, paragraph, compressed_paths, self.image_width_cm, self.gap_width_cm)
                    break

        print_blank_line()
        print(f"Placeholder '{self.placeholder_text}' found and replaced.")

        # Set document properties.
        document.core_properties.author = self.doc_author
        document.core_properties.comments = self.doc_comments

        try:
            document.save(self.docx_output)
        except Exception as err:
            print(f"Failed to save document '{self.docx_output}': {err}")
            sys.exit(1)
        print(f"Document saved as: {self.docx_output}")
        print_blank_line()
        print("Time saved once again. Thanks, ot2i7ba!")
        print("Check the result and adjust it if necessary!")

def main() -> None:
    args = parse_arguments()
    app = jpg2Document(args)
    app.run()

if __name__ == "__main__":
    clear_console()
    try:
        main()
    except KeyboardInterrupt:
        print_blank_line()
        print("\nProcess interrupted by user. Exiting gracefully...")
        sys.exit(0)
