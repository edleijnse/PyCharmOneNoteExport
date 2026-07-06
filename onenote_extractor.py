import aspose.note as asn


def load(one_file: str) -> asn.Document:
    return asn.Document(one_file)


def get_page_text(page: asn.Page) -> str:
    lines = []
    for rt in page.GetChildNodes(asn.RichText):
        text = rt.Text.strip()
        if text:
            lines.append(text)
    return "\n".join(lines)


def get_page_title(page: asn.Page) -> str:
    try:
        return page.Title.TitleText.Text.strip()

    except Exception:
        pass
    return "(untitled)"


def extract_pages(doc: asn.Document, section_name: str = "") -> list[dict]:
    """Return one dict per page with section, page name, level, and text."""
    if not section_name:
        section_name = "Section"

    results = []
    for page in doc:
        results.append({
            "section": section_name,
            "page": get_page_title(page),
            # "level": page.level,
            "text": get_page_text(page),
        })
    return results


def print_pages(pages: list[dict]) -> None:
    pad = "  "
    for entry in pages:
        # indent = pad * (entry["level"] - 1)
        indent = pad
        print("=" * 72)
        print(f"{indent}[{entry['section']}] > {entry['page']}")
        print("-" * 72)
        if entry["text"]:
            for line in entry["text"].splitlines():
                print(f"{indent}  {line}")
        else:
            print(f"{indent}  (no text content)")
        print()


def write_pages(pages: list[dict], output_file: str) -> None:
    pad = "  "
    with open(output_file, "w", encoding="utf-8") as f:
        for entry in pages:
            # indent = pad * (entry["level"] - 1)
            indent = pad
            f.write("=" * 72 + "\n")
            f.write(f"{indent}[{entry['section']}] > {entry['page']}\n")
            f.write("-" * 72 + "\n")
            if entry["text"]:
                for line in entry["text"].splitlines():
                    f.write(f"{indent}  {line}\n")
            else:
                f.write(f"{indent}  (no text content)\n")
            f.write("\n")


if __name__ == "__main__":
    import os

    ONE_FILE = r"D:\OneNote\Paspoort.one"
    ONE_FILE_OUTPUT = r"D:\OneNote\Paspoort.one.text"
    section_name = os.path.splitext(os.path.basename(ONE_FILE))[0]

    doc = load(ONE_FILE)
    pages = extract_pages(doc, section_name)
    print_pages(pages)
    write_pages(pages, ONE_FILE_OUTPUT)
