# omml2latex

Convert OMML (Office Math Markup Language) elements from MS Office documents
into KaTeX-compatible LaTeX strings. No external dependencies.

## Installation

```bash
pip install omml2latex
```

## Quick Start

MS Office documents (`.pptx`, `.docx`) store math equations internally as
OMML XML elements. Pass an `m:oMath` element to `convert_omml` to get a
LaTeX string:

```python
import xml.etree.ElementTree as ET
from omml2latex import convert_omml

xml_string = """
<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <m:r>
        <m:rPr><m:scr m:val="double-struck"/><m:sty m:val="b"/></m:rPr>
        <m:t>Rα∞</m:t>
    </m:r>
</m:oMath>
"""

root = ET.fromstring(xml_string)
print(convert_omml(root))
# $\mathbf{\mathbb{R \alpha \infty }}$
```

## API

### `convert_omml(node)`

| Parameter | Type | Description |
|-----------|------|-------------|
| `node` | `xml.etree.ElementTree.Element` | An `m:oMath`, `m:oMathPara`, or `m:mathPr` element |

Returns a KaTeX-compatible LaTeX string:
- `m:oMath` → `$...$` (inline math)
- `m:oMathPara` → `$$...$$` (display math)
- `m:mathPr` → `""` (global math settings, no output)

## Extracting Equations from a PPTX or DOCX File

OOXML files (`.pptx`, `.docx`) are ZIP archives containing XML files inside.
To extract all equations from a file, open it with `zipfile` and search for
`m:oMath` elements:

```python
import zipfile
import xml.etree.ElementTree as ET
from omml2latex import convert_omml

MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

with zipfile.ZipFile("presentation.pptx") as zf:
    for name in zf.namelist():
        if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
            tree = ET.fromstring(zf.read(name))
            for node in tree.iter(f"{{{MATH_NS}}}oMath"):
                print(convert_omml(node))
```

For `.docx` files, replace `ppt/slides/slide` with `word/document`:

```python
with zipfile.ZipFile("document.docx") as zf:
    tree = ET.fromstring(zf.read("word/document.xml"))
    for node in tree.iter(f"{{{MATH_NS}}}oMath"):
        print(convert_omml(node))
```

## Features

- **ECMA-376 based**: built from the official OMML schema definitions
- **Recursive descent parser**: handles nested structures (fractions, matrices, accents, large operators, etc.)
- **Unicode math mapping**: automatically maps Unicode Mathematical Alphabet characters (𝑓, 𝜋, 𝒜, …) to their LaTeX equivalents

## License

Apache 2.0 — Copyright 2026 INSEONG LEE
