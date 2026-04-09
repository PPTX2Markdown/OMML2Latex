"""Public package interface for OMML to LaTeX conversion."""

from ._parser import convert_omml

__version__ = "0.1.1"
__all__ = [
    "convert_omml",
    "main",
]


def main() -> None:
    """CLI entry point — omml2latex [options] input."""
    import argparse
    import sys
    import zipfile
    import xml.etree.ElementTree as ET
    from pathlib import Path

    parser = argparse.ArgumentParser(
        description="Extract and convert OMML equations to LaTeX."
    )
    parser.add_argument("input", help="Input file (.pptx, .docx)")
    parser.add_argument("-o", "--output", help="Output file (default: stdout)")
    args = parser.parse_args()

    path = Path(args.input)
    suffix = path.suffix.lower()

    if suffix not in (".pptx", ".docx"):
        print(f"Unsupported file type: {path.suffix}", file=sys.stderr)
        sys.exit(1)

    def _strip_ns(tag: str) -> str:
        return tag.split("}")[-1] if "}" in tag else tag

    def _extract_from_para(para_node: ET.Element) -> list[str]:
        """단락(a:p) 노드에서 OMML 수식을 추출해 LaTeX 문자열 목록으로 반환합니다."""
        math_nodes: list[ET.Element] = []

        for child in para_node:
            local = _strip_ns(child.tag)
            # AlternateContent 래퍼 처리
            if local == "AlternateContent":
                choice = next((c for c in child if _strip_ns(c.tag) == "Choice"), None)
                if choice is not None:
                    for sp in choice:
                        if _strip_ns(sp.tag) == "sp":
                            for txBody in sp:
                                if _strip_ns(txBody.tag) == "txBody":
                                    for p in txBody:
                                        if _strip_ns(p.tag) == "p":
                                            return _extract_from_para(p)
            # a14:m 래퍼 처리
            if "}" in child.tag and _strip_ns(child.tag) == "m":
                for math_el in child:
                    if _strip_ns(math_el.tag) in ("oMath", "oMathPara"):
                        math_nodes.append(math_el)

        # 단락 직계 자식으로 oMath / oMathPara가 있는 경우
        for child in para_node:
            if _strip_ns(child.tag) in ("oMath", "oMathPara"):
                math_nodes.append(child)

        return [convert_omml(m) for m in math_nodes]

    equations: list[str] = []
    with zipfile.ZipFile(path) as zf:
        if suffix == ".pptx":
            slide_names = sorted(
                (n for n in zf.namelist()
                 if n.startswith("ppt/slides/slide") and n.endswith(".xml")),
                key=lambda x: int(x.split("slide")[-1].split(".xml")[0]),
            )
            xml_names = slide_names
        else:
            xml_names = ["word/document.xml"]

        for name in xml_names:
            try:
                tree = ET.fromstring(zf.read(name))
            except Exception:
                continue
            for el in tree.iter():
                if _strip_ns(el.tag) == "p":
                    equations.extend(_extract_from_para(el))

    if not equations:
        print("No equations found.", file=sys.stderr)
        sys.exit(1)

    output = "\n".join(equations)
    if args.output:
        Path(args.output).write_text(output, encoding="utf-8")
    else:
        print(output)
