# PyCharmOneNoteExport

Extracts text from Microsoft OneNote `.one` files without requiring the OneNote desktop application.

## Requirements

- Python 3.10+
- [`aspose-note`](https://pypi.org/project/aspose-note/)

```
pip install aspose-note
```

## Usage

### As a script

Edit the `ONE_FILE` path at the bottom of `onenote_extractor.py` and run:

```
python onenote_extractor.py
```

### As a module

```python
import onenote_extractor as one

doc = one.load(r"D:\OneNote\Paspoort.one")
pages = one.extract_pages(doc, section_name="Paspoort")
one.print_pages(pages)
```

### Output format

Each page is printed as:

```
========================================================================
  [SectionName] > Page Title
------------------------------------------------------------------------
  Line one of page text
  Line two of page text
```

## API

| Function | Description |
|---|---|
| `load(one_file)` | Opens a `.one` file and returns an `aspose.note.Document` |
| `extract_pages(doc, section_name)` | Returns a list of dicts with `section`, `page`, and `text` keys |
| `get_page_title(page)` | Reads the title of a single page; returns `"(untitled)"` if absent |
| `get_page_text(page)` | Concatenates all `RichText` nodes on a page into a single string |
| `print_pages(pages)` | Pretty-prints the list returned by `extract_pages` |

## Interactive exploration

Open `notebook.ipynb` in PyCharm or Jupyter to run the extraction step by step.

To debug with breakpoints, set them in `onenote_extractor.py` and run it directly via PyCharm's debugger (`Shift+F9`).
