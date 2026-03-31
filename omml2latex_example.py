import xml.etree.ElementTree as ET
import zipfile
import sys
from omml2latex import convert_omml

def _strip_ns(tag):
    """XML 태그에서 네임스페이스를 제거하고 로컬 태그 이름만 반환합니다.
       예: '{http://...}oMath' -> 'oMath' """
    if not tag:
        return ""
    if "}" in tag:
        return tag.split("}")[-1]
    return tag

def extract_math_from_para(para_node: ET.Element):
    """
    pptx2md와 동일한 방식으로 단락(a:p) 엘리먼트 구문 트리에서 OMML 수식을 추출합니다.
    추출된 수식 엘리먼트를 LaTeX 문자열로 변환하여 리스트로 반환합니다.
    """
    math_nodes = []

    # 단락 직계 자식들을 순회하며 AlternateContent 등 래퍼(wrapper)를 확인합니다.
    for child in para_node:
        local = _strip_ns(child.tag)
        if local == "AlternateContent":
            choice = None
            for c in child:
                if _strip_ns(c.tag) == "Choice":
                    choice = c
                    break
            if choice is not None:
                for sp in choice:
                    if _strip_ns(sp.tag) == "sp":
                        for txBody in sp:
                            if _strip_ns(txBody.tag) == "txBody":
                                for p in txBody:
                                    if _strip_ns(p.tag) == "p":
                                        # 발견된 단락(p)에 대해 재귀적으로 수식을 추출
                                        return extract_math_from_para(p)

        # a14:m 래퍼 태그 안에 수식(m:oMath)이 들어있는 경우를 처리합니다.
        if "}" in child.tag:
            ns, tag_local = child.tag.split("}", 1)
            ns = ns.lstrip("{")
            if tag_local == "m":
                # a14:m 노드 내부에서 m:oMathPara 또는 m:oMath를 검색
                for math_el in child:
                    if _strip_ns(math_el.tag) in ("oMath", "oMathPara"):
                        math_nodes.append(math_el)

    # AlternateContent 외에도 단락 자식으로 oMath / oMathPara가 직접 존재하는지 확인
    for child in para_node:
        if _strip_ns(child.tag) in ("oMath", "oMathPara"):
            math_nodes.append(child)

    if not math_nodes:
        return []

    # omml2latex 호출하여 수식 노드를 LaTeX 변환
    parts = []
    for m in math_nodes:
        parts.append(convert_omml(m))
    return parts

def extract_math_to_md(pptx_path, output_md_path):
    """
    PPTX 파일에서 OMML 수식 노드를 추출하여 Markdown 파일로 저장합니다.
    외부 의존성 없이 PPTX의 슬라이드 XML 파일들을 직접 파싱합니다.
    """
    try:
        with zipfile.ZipFile(pptx_path, 'r') as zf:
            # zip 파일 내에서 ppt/slides/slideN.xml 파일들의 목록을 가져옵니다
            slide_files = [f for f in zf.namelist() if f.startswith('ppt/slides/slide')]
            # 슬라이드 번호순으로 자연 정렬 (예: slide1.xml, slide2.xml, ...)
            slide_files.sort(key=lambda x: int(x.split('slide')[-1].split('.xml')[0]))
            
            with open(output_md_path, 'w', encoding='utf-8') as md_file:
                md_file.write(f"# Extracted Math from `{pptx_path}`\n\n")
                
                for slide_file in slide_files:
                    slide_num = slide_file.split('slide')[-1].split('.xml')[0]
                    md_file.write(f"## Slide {slide_num}\n\n")
                    
                    xml_content = zf.read(slide_file)
                    root = ET.fromstring(xml_content)
                    
                    # 해당 슬라이드에서 단락(Paragraph, a:p) 태그를 전부 찾습니다
                    paragraphs = []
                    for el in root.iter():
                        if _strip_ns(el.tag) == "p":
                            paragraphs.append(el)
                    
                    found_math = False
                    for p in paragraphs:
                        math_parts = extract_math_from_para(p)
                        if math_parts:
                            found_math = True
                            for latex in math_parts:
                                # 수식 문자열의 시작 부분 기호를 확인해 블록/인라인 구문 삽입
                                if latex.startswith("$$"):
                                    md_file.write(f"**Block Math:**\n\n{latex}\n\n")
                                else:
                                    md_file.write(f"**Inline Math:** \n\n{latex}\n\n")
                            
                    if not found_math:
                        md_file.write("*No math found on this slide.*\n\n")

        print(f"✅ Successfully extracted math to: {output_md_path}")

    except Exception as e:
        print(f"Error parsing PPTX: {e}")

def main():
    if len(sys.argv) < 2:
        print("사용법: python omml2latex_example.py <pptx_파일_경로> [출력할_md_파일_경로]")
        sys.exit(1)
        
    pptx_path = sys.argv[1]
    
    if len(sys.argv) > 2:
        output_md_path = sys.argv[2]
    else:
        # Generate default output name
        output_md_path = pptx_path.replace('.pptx', '_math_extracted.md')
        
    print(f"Extracting math from: {pptx_path}")
    extract_math_to_md(pptx_path, output_md_path)

if __name__ == "__main__":
    main()
