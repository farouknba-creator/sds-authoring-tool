def extract_sections_from_pdf(pdf_path):
    """Extract sections from a GHS‑compliant SDS PDF. Handles:
       - SECTION 1: Identification
       - Section 1. Identification
       - 1. Identification
       - 1. HAZARDS IDENTIFICATION
       and similar variations.
    """
    import re
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
    except Exception as e:
        return [f"Error extracting PDF: {e}"]

    if not text.strip():
        return ["[No extractable text – PDF may be scanned.]"]

    # Normalise line endings
    text = text.replace('\r\n', '\n').replace('\r', '\n')

    # Pattern to match GHS section headings, capturing the full heading line
    # This will match:
    #   SECTION 1: Identification
    #   Section 1. Identification
    #   1. Identification
    #   2. HAZARDS IDENTIFICATION
    #   SECTION 2 – Hazards identification
    #   etc.
    heading_pattern = re.compile(
        r'^(?:SECTION\s+)?(\d{1,2})\s*[\.\:\)\-]?\s*(.*?)\s*$',
        re.IGNORECASE | re.MULTILINE
    )

    # Find all matches with their positions
    matches = list(re.finditer(heading_pattern, text))
    if not matches:
        # No headings found – return the whole text as one block
        return [text.strip()]

    sections = []
    # First section starts at the first heading; discard any preamble before it
    for i, match in enumerate(matches):
        section_num = match.group(1)
        title = match.group(2).strip()
        start = match.start()
        end = matches[i+1].start() if i+1 < len(matches) else len(text)
        body = text[start:end].strip()
        # Create a clean label for the section
        label = f"Section {section_num} – {title}" if title else f"Section {section_num}"
        # The section content is the whole block from the heading onward
        sections.append(body)

    # If there is any text before the first heading, we could optionally add it as a
    # "Preamble" section, but SDS have supplier info there that you might want to skip.
    # We'll skip it for cleanliness.
    return sections
