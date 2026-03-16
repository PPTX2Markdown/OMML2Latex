# OMML2LaTeX

OMML (Office Math Markup Language) 수식을 LaTeX 코드로 변환해주는 파이썬 스크립트(`omml2latex.py`)입니다.

## 개요

이 프로젝트는 MS Office의 수식 형식인 OMML을 널리 쓰이는 LaTeX 형식으로 직관적으로 변환하기 위해 작성되었습니다.

ECMA-376 문서를 참고하여 OMML의 구조를 파악하였으며, 프로젝트 내부의 `shared-math-strict.rnc`, `shared-math-strict.xsd`, `shared-math-transitional.xsd` 등 스키마 정의 파일(RNC, XSD)을 기반으로 OMML 요소들을 처리합니다.

각 OMML 노드를 순회하며 구문을 분석하는 방식은 **Recursive Descent Parser(재귀적 하향 구문 분석기)**의 형태를 띠도록 코드를 작성하였습니다.

## 주요 특징

- **ECMA-376 표준 기반**: 공식 스키마와 문서를 참조하여 견고한 변환 규칙 적용
- **재귀적 하향 파서 구조**: OMML의 계층적 트리 구조(`omath`, `f`(분수), `r`(런), `t`(텍스트) 등)를 재귀적으로 순회하여 LaTeX 문자열로 조합
- **유니코드 심볼 매핑 지원**: 그리스 문자 및 다양한 수학 기호(예: `α` -> `\alpha`)를 적절한 LaTeX 매크로로 자동 치환

## 파일 구조

- `omml2latex.py`: OMML 파싱 및 LaTeX 변환 핵심 로직
- `shared-math-strict.rnc` / `shared-math-strict.xsd` / `shared-math-transitional.xsd`: 파서 작성 시 참고한 ECMA-376 OMML 스키마 문서들

## 사용 방법

이 스크립트는 MS Office 문서(`.docx`, `.pptx` 등) 내부에 포함된 XML 형식의 수식 데이터(OMML)를 추출하여 변환할 때 사용할 수 있습니다. MS Office 문서는 실제로는 ZIP 압축 파일(OOXML 규격)이므로, `zipfile` 모듈 등을 이용해 내부 XML(`slide1.xml`, `document.xml` 등)을 읽어들인 뒤 루트부터 순회하며 변환합니다.

다음은 파워포인트(`.pptx`) 파일에서 슬라이드 내부의 수식을 찾아 LaTeX로 변환하는 개념적인 사용 예시입니다.

```python
import zipfile
import xml.etree.ElementTree as ET
from omml2latex import parse_omml_to_latex, get_tag

def extract_math(element):
    """
    재귀적으로 XML 트리를 순회하면서 수식 컴포넌트(OMML)를 찾아 변환합니다.
    """
    tag = get_tag(element)

    # 1. 수식 문단 (oMathPara - 블록 수식) 처리
    if tag == "oMathPara":
        latex = parse_omml_to_latex(element)
        print("블록 수식:\n", latex)
        return  # 자식 oMath 탐색 방지를 위해 반환

    # 2. 인라인 수식 (oMath) 처리
    elif tag == "oMath":
        latex = parse_omml_to_latex(element)
        print("인라인 수식:\n", f"$${latex}$$")
        return

    # 수식 요소가 아니면 자식 노드들을 계속 탐색
    for child in element:
        extract_math(child)

# .pptx 파일 압축 해제 및 파싱
with zipfile.ZipFile("mathematicalExpression.pptx", 'r') as z:
    # ppt 내의 슬라이드 원본 xml 파일만 필터링
    slides = [f for f in z.namelist() if f.startswith("ppt/slides/slide") and f.endswith(".xml")]

    for slide in sorted(slides):
        xml_data = z.read(slide)
        root = ET.fromstring(xml_data)

        # 수식 추출 및 LaTeX 변환 실행
        extract_math(root)
```
