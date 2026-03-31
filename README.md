# OMML2LaTeX

OMML (Office Math Markup Language) 수식을 LaTeX 코드로 변환해주는 설치 가능한 Python 패키지입니다.

## 개요

이 프로젝트는 MS Office의 수식 형식인 OMML을 널리 쓰이는 LaTeX 형식으로 직관적으로 변환하기 위해 작성되었습니다.

ECMA-376 문서를 참고하여 OMML의 구조를 파악하였으며, 프로젝트 내부의 `shared-math-strict.rnc`, `shared-math-strict.xsd`, `shared-math-transitional.xsd` 등 스키마 정의 파일(RNC, XSD)을 기반으로 OMML 요소들을 처리합니다.

각 OMML 노드를 순회하며 구문을 분석하는 방식은 **Recursive Descent Parser(재귀적 하향 구문 분석기)**의 형태를 띠도록 코드를 작성하였습니다.

## 주요 특징

- **ECMA-376 표준 기반**: 공식 스키마와 문서를 참조하여 견고한 변환 규칙 적용
- **재귀적 하향 파서 구조**: OMML의 계층적 트리 구조(`omath`, `f`(분수), `r`(런), `t`(텍스트) 등)를 재귀적으로 순회하여 LaTeX 문자열로 조합
- **유니코드 심볼 매핑 지원**: 그리스 문자 및 다양한 수학 기호(예: `α` -> `\alpha`)를 적절한 LaTeX 매크로로 자동 치환

## 패키지 구조

- `omml2latex/`: import 가능한 패키지 본체
- `omml2latex/__init__.py`: 얇은 공개 API 레이어
- `omml2latex/_parser.py`: OMML 파싱 및 LaTeX 변환 핵심 구현
- `pyproject.toml`: `pip` / `uv` 설치용 메타데이터
- `shared-math-strict.rnc` / `shared-math-strict.xsd` / `shared-math-transitional.xsd`: 파서 작성 시 참고한 ECMA-376 OMML 스키마 문서들

## 설치

로컬 개발용 editable 설치:

```bash
uv pip install -e .
```

일반 설치:

```bash
pip install .
```

Git dependency 설치:

```bash
pip install "git+https://github.com/<OWNER>/OMML2Latex.git"
```

```toml
[project]
dependencies = [
  "omml2latex @ git+https://github.com/<OWNER>/OMML2Latex.git",
]
```

`uv`를 쓰는 프로젝트에서도 동일하게 Git dependency 형태로 추가할 수 있습니다.

## 사용 방법

기본 사용 방식은 다른 Python 프로젝트에서 라이브러리로 import해서 OMML XML 문자열이나 XML element를 LaTeX로 변환하는 것입니다.

문자열 XML이 있다면:

```python
from omml2latex import parse_omml_xml

latex = parse_omml_xml(
    '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
    "<m:r><m:t>x</m:t></m:r>"
    "</m:oMath>"
)
```

이미 `xml.etree.ElementTree.Element` 형태의 OMML 노드를 갖고 있다면:

```python
from omml2latex import parse_omml_to_latex

latex = parse_omml_to_latex(omml_element)
```

## PPTX 연동 예시

MS Office 문서(`.docx`, `.pptx` 등) 내부의 OMML을 직접 순회하며 변환할 수도 있습니다.

```python
import zipfile
import xml.etree.ElementTree as ET

from omml2latex import parse_omml_to_latex


def extract_math(element):
    tag = element.tag.split("}")[-1]

    if tag == "oMathPara":
        print(parse_omml_to_latex(element))
        return

    if tag == "oMath":
        print(parse_omml_to_latex(element))
        return

    for child in element:
        extract_math(child)


with zipfile.ZipFile("mathematicalExpression.pptx", "r") as zf:
    slides = [
        name
        for name in zf.namelist()
        if name.startswith("ppt/slides/slide") and name.endswith(".xml")
    ]

    for slide in sorted(slides):
        root = ET.fromstring(zf.read(slide))
        extract_math(root)
```

별도 예제 스크립트로는 `convert_pptx_math.py`를 참고할 수 있습니다.
이 스크립트는 패키지의 일부 public interface는 아니고, 샘플/실험용 유틸리티입니다.
```

## Public API

- `parse_omml_to_latex(node)`: `xml.etree.ElementTree.Element` 입력을 LaTeX 문자열로 변환
- `convert_omml_to_latex(node)`: 위 함수의 alias
- `parse_omml_xml(xml)`: OMML XML 문자열을 바로 변환
