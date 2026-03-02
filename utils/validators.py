def parse_page_ranges(text, max_pages):
    pages = set()
    for part in text.split(","):
        part = part.strip()
        if "-" in part:
            start, end = map(int, part.split("-"))
            pages.update(range(start-1, min(end, max_pages)))
        else:
            pages.add(int(part)-1)
    return {p for p in pages if 0 <= p < max_pages}