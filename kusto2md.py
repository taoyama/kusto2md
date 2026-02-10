"""
Kusto Explorer Clipboard to Markdown Converter

Reads the HTML Format from the Windows clipboard (which contains the full
query + results as Kusto Explorer copies them), and converts to markdown.

Usage:
  1. Copy from Kusto Explorer (query + results).
  2. Run this script.
  3. The markdown is in your clipboard, ready to paste.
"""

import ctypes
import ctypes.wintypes
import re
import sys
from html.parser import HTMLParser
from html import unescape

# --- Win32 Clipboard Access ---
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

user32.GetClipboardData.restype = ctypes.c_void_p
kernel32.GlobalLock.restype = ctypes.c_void_p
kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
kernel32.GlobalSize.restype = ctypes.c_size_t
kernel32.GlobalSize.argtypes = [ctypes.c_void_p]

try:
    import pyperclip
except ImportError:
    print("Installing pyperclip...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyperclip"])
    import pyperclip

URL_RE = re.compile(r'https?://[^\s\)\]>"]+')


def get_html_from_clipboard() -> str | None:
    """Read the 'HTML Format' data from the Windows clipboard."""
    fmt = user32.RegisterClipboardFormatW("HTML Format")
    if not fmt:
        return None

    if not user32.OpenClipboard(0):
        return None
    try:
        handle = user32.GetClipboardData(fmt)
        if not handle:
            return None
        size = kernel32.GlobalSize(handle)
        if not size:
            return None
        ptr = kernel32.GlobalLock(handle)
        if not ptr:
            return None
        try:
            data = ctypes.string_at(ptr, size)
            return data.decode("utf-8", errors="replace").rstrip("\x00")
        finally:
            kernel32.GlobalUnlock(handle)
    finally:
        user32.CloseClipboard()


# --- HTML Parsing Helpers ---

class TextExtractor(HTMLParser):
    """Extract visible text from HTML, converting <p> boundaries to newlines."""

    def __init__(self):
        super().__init__()
        self.lines = []
        self._current_line = []
        self._skip = False

    def handle_starttag(self, tag, attrs):
        if tag == "style":
            self._skip = True
        if tag in ("p", "br"):
            self._flush_line()

    def handle_endtag(self, tag):
        if tag == "style":
            self._skip = False

    def handle_data(self, data):
        if not self._skip:
            self._current_line.append(data)

    def handle_entityref(self, name):
        self.handle_data(unescape(f"&{name};"))

    def handle_charref(self, name):
        self.handle_data(unescape(f"&#{name};"))

    def _flush_line(self):
        text = "".join(self._current_line)
        if text:
            self.lines.append(text)
        self._current_line = []

    def get_text(self) -> str:
        self._flush_line()
        return "\n".join(self.lines)


class TableExtractor(HTMLParser):
    """Extract rows/cells from an HTML <table>."""

    def __init__(self):
        super().__init__()
        self.rows: list[list[str]] = []
        self._in_td = False
        self._current_cell: list[str] = []
        self._current_row: list[str] = []

    def handle_starttag(self, tag, attrs):
        if tag == "td":
            self._in_td = True
            self._current_cell = []
        elif tag == "tr":
            self._current_row = []

    def handle_endtag(self, tag):
        if tag == "td":
            self._in_td = False
            self._current_row.append("".join(self._current_cell).strip())
        elif tag == "tr":
            if self._current_row:
                self.rows.append(self._current_row)

    def handle_data(self, data):
        if self._in_td:
            self._current_cell.append(data)

    def handle_entityref(self, name):
        self.handle_data(unescape(f"&{name};"))

    def handle_charref(self, name):
        self.handle_data(unescape(f"&#{name};"))


def extract_execute_links(html: str) -> list[tuple[str, str]]:
    """Extract (label, url) pairs from the Execute: [...] line."""

    class LinkExtractor(HTMLParser):
        def __init__(self):
            super().__init__()
            self.links = []
            self._current_href = None
            self._current_text = []

        def handle_starttag(self, tag, attrs):
            if tag == "a":
                attrs_dict = dict(attrs)
                self._current_href = attrs_dict.get("href")
                self._current_text = []

        def handle_endtag(self, tag):
            if tag == "a" and self._current_href:
                label = "".join(self._current_text).strip()
                self.links.append((label, unescape(self._current_href)))
                self._current_href = None

        def handle_data(self, data):
            if self._current_href is not None:
                self._current_text.append(data)

    # Get the part before <div data-type="query">
    query_div_idx = html.find('<div data-type="query">')
    header_html = html[:query_div_idx] if query_div_idx != -1 else html[:2000]

    extractor = LinkExtractor()
    extractor.feed(header_html)
    return extractor.links


def extract_query(html: str) -> str | None:
    """Extract the KQL query text from <div data-type="query">...</div>."""
    match = re.search(r'<div data-type="query">(.*?)</div>', html, re.DOTALL)
    if not match:
        return None
    query_html = match.group(1)
    extractor = TextExtractor()
    extractor.feed(query_html)
    text = extractor.get_text()
    # Clean up: replace non-breaking spaces with regular spaces
    text = text.replace("\xa0", " ")
    return text.strip() if text.strip() else None


def extract_table(html: str) -> list[list[str]] | None:
    """Extract table rows from the HTML."""
    table_match = re.search(r'<table[^>]*>(.*?)</table>', html, re.DOTALL)
    if not table_match:
        return None
    extractor = TableExtractor()
    extractor.feed(table_match.group(1))
    # Clean up: replace non-breaking spaces
    for row in extractor.rows:
        for i in range(len(row)):
            row[i] = row[i].replace("\xa0", " ")
    return extractor.rows if extractor.rows else None


def extract_cluster_url(html: str) -> str | None:
    """Extract the cluster connection URL (not a deep-link with query= param)."""
    query_div_idx = html.find('<div data-type="query">')
    header_html = html[:query_div_idx] if query_div_idx != -1 else html[:3000]
    matches = re.findall(r'href="(https?://[^"]+)"', header_html)
    for url in matches:
        url = unescape(url)
        if "query=" not in url:
            return url
    return None


# --- Markdown Conversion ---

def linkify_url(url: str) -> str:
    """Turn a raw URL into a short clickable markdown link."""
    stripped = url.rstrip("/")
    parts = stripped.split("/")
    label = "/".join(parts[-2:]) if len(parts) >= 2 else stripped
    if len(label) > 60:
        label = label[:57] + "..."
    return f"[{label}]({url})"


def rows_to_markdown(rows: list[list[str]]) -> str:
    """Convert table rows to a markdown table. Auto-linkify URL cells."""
    if not rows:
        return ""
    max_cols = max(len(r) for r in rows)
    for row in rows:
        while len(row) < max_cols:
            row.append("")

    # Linkify URLs in data cells (skip header row 0)
    for r in range(1, len(rows)):
        for c in range(max_cols):
            cell = rows[r][c]
            if URL_RE.search(cell):
                rows[r][c] = URL_RE.sub(lambda m: linkify_url(m.group(0)), cell)

    col_widths = [max(len(rows[r][c]) for r in range(len(rows))) for c in range(max_cols)]
    col_widths = [max(w, 3) for w in col_widths]

    def fmt(row):
        return "| " + " | ".join(cell.ljust(col_widths[i]) for i, cell in enumerate(row)) + " |"

    header = fmt(rows[0])
    sep = "| " + " | ".join("-" * w for w in col_widths) + " |"
    data = [fmt(row) for row in rows[1:]]
    return "\n".join([header, sep] + data)


def build_markdown(execute_links, cluster_url, query, table_rows):
    """Assemble the full markdown output."""
    parts = []

    # Header with execute links
    if execute_links or cluster_url:
        parts.append("### Query\n")
        if cluster_url:
            parts.append(f"> **Cluster:** {cluster_url}\n")
        if execute_links:
            link_items = []
            for label, url in execute_links:
                if "query=" in url:  # deep-links only
                    link_items.append(f"[{label}]({url})")
            if link_items:
                parts.append("> **Open in:** " + " | ".join(link_items) + "\n")

    # Query code block
    if query:
        if not execute_links and not cluster_url:
            parts.append("### Query\n")
        parts.append(f"```kql\n{query}\n```")

    # Results table
    if table_rows:
        parts.append("\n### Results\n")
        parts.append(rows_to_markdown(table_rows))

    return "\n".join(parts)


def main():
    # Try HTML Format first (has query + results)
    html = get_html_from_clipboard()

    if html and "<div data-type=" in html:
        print("Found Kusto Explorer HTML clipboard data.")
        execute_links = extract_execute_links(html)
        cluster_url = extract_cluster_url(html)
        query = extract_query(html)
        table_rows = extract_table(html)

        markdown = build_markdown(execute_links, cluster_url, query, table_rows)
    else:
        # Fallback: plain text (tab-separated results only)
        print("No Kusto HTML found, falling back to plain text clipboard.")
        text = pyperclip.paste()
        if not text or not text.strip():
            print("Clipboard is empty.")
            return

        # Simple TSV conversion
        lines = [l for l in text.strip().splitlines() if l.strip()]
        rows = [l.split("\t") for l in lines]
        markdown = rows_to_markdown(rows) if rows else ""

    if not markdown:
        print("No data to convert.")
        return

    pyperclip.copy(markdown)
    print("\n--- Markdown output ---")
    print(markdown)
    print(f"\n✓ Markdown copied to clipboard! ({len(markdown)} chars)")


if __name__ == "__main__":
    main()
