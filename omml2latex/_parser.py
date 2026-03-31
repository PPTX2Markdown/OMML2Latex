"""Convert OMML (Office Math Markup Language) elements into LaTeX strings."""

from __future__ import annotations

import xml.etree.ElementTree as ET

__version__ = "0.1.0"
__all__ = [
    "parse_omml_to_latex",
    "convert_omml_to_latex",
    "parse_omml_xml",
    "get_tag",
    "get_child",
    "get_children",
    "get_val",
]

# ==============================================================================
# Helper / Utility
# ==============================================================================
def get_tag(node):
    return node.tag.split('}')[-1] if node is not None else ""

def get_child(node, tag):
    if node is None: return None
    for child in node:
        if get_tag(child) == tag: return child
    return None

def get_children(node, tag):
    if node is None: return []
    return [child for child in node if get_tag(child) == tag]

def get_val(node, attr='val'):
    if node is None: return None
    # 속성명에 네임스페이스가 붙어있을 경우를 대비해 순회하며 속성명 검사
    # transitional과 strict 스키마 모두에서 val 속성은 동일한 이름이지만, 네임스페이스가 다르므로 (ECMA-376 참고)
    for k, v in node.attrib.items():
        if k == attr or k.endswith(f"}}{attr}"):
            return v
    return None

# ==============================================================================
# Unicode Mapping & Cleanup Table
# ==============================================================================
UNICODE_TO_LATEX = {
    'α': '\\alpha', 'β': '\\beta', 'γ': '\\gamma', 'δ': '\\delta', 'ε': '\\epsilon',
    'ζ': '\\zeta', 'η': '\\eta', 'θ': '\\theta', 'ι': '\\iota', 'κ': '\\kappa',
    'λ': '\\lambda', 'μ': '\\mu', 'ν': '\\nu', 'ξ': '\\xi', 'ο': '\\omicron',
    'π': '\\pi', 'ρ': '\\rho', 'σ': '\\sigma', 'τ': '\\tau', 'υ': '\\upsilon',
    'φ': '\\phi', 'χ': '\\chi', 'ψ': '\\psi', 'ω': '\\omega',
    'Δ': '\\Delta', 'Γ': '\\Gamma', 'Θ': '\\Theta', 'Λ': '\\Lambda', 'Ξ': '\\Xi',
    'Π': '\\Pi', 'Σ': '\\Sigma', 'Φ': '\\Phi', 'Ψ': '\\Psi', 'Ω': '\\Omega',
    
    '∞': '\\infty', '≈': '\\approx', '≠': '\\neq', '≤': '\\leq', '≥': '\\geq',
    '×': '\\times', '÷': '\\div', '±': '\\pm', '·': '\\cdot', '°': '\\circ',
    '∂': '\\partial', '∇': '\\nabla', 
    '∈': '\\in', '∉': '\\notin', '⊂': '\\subset', '⊃': '\\supset', '∪': '\\cup', '∩': '\\cap',
    '∧': '\\land', '∨': '\\vee', '∀': '\\forall', '∃': '\\exists', '∄': '\\nexists',
    '∅': '\\emptyset', '←': '\\leftarrow', '→': '\\rightarrow', '↔': '\\leftrightarrow',
    '⇒': '\\Rightarrow', '⇐': '\\Leftarrow', '⇔': '\\Leftrightarrow',
}

def cleanup_text(text):
    if not text: return ""
    invisibles = ['\u2061', '\u2062', '\u2063', '\u2064']
    for ch in invisibles:
        text = text.replace(ch, '')
    # Word에서 띄어쓰기로 사용되는 Non-breaking space(\xa0)를 일반 공백으로 변환
    text = text.replace('\xa0', ' ')
    for k, v in UNICODE_TO_LATEX.items():
        text = text.replace(k, f" {v} ")
    return text

# ==============================================================================
# Simple Types (ST) - Data Extraction
# RNC 파일의 기본 데이터 타입들입니다. 추출은 1:1로 전부 진행하되, 
# LaTeX 생태계와 맞지 않는 부분은 CT(Complex Type) 단계에서 무시([IGNORE])
# ==============================================================================
def parse_m_ST_Integer255(val): 
    # [IGNORE: Word UI 렌더링 전용] 1~255 정수. 주로 "수동 줄바꿈 시 화면의 어느 x좌표에 정렬할 것인가(alnAt)" 같은 GUI 전용 상태값. 
    # 왜 무시하는가? -> LaTeX는 환경(\begin{align} 등) 내에서 '&' 기호를 통해 정렬을 매우 수학적인 구조로 자동 처리하기 때문에 물리적인 인덱스는 의미가 없다.
    return int(val) if val else None

def parse_m_ST_Integer2(val):
    # [IGNORE: LaTeX 자동화 영역] -2~2 정수. "첨자 크기(argSz)를 기본치보다 강제로 더 키울까 줄일까"를 결정.
    # 왜 무시하는가? -> LaTeX의 TeX 엔진은 첨자 중첩 깊이에 따라 크기(\scriptstyle, \scriptscriptstyle)를 자체 로직으로 자동 결정. 
    return int(val) if val else None

def parse_m_ST_SpacingRule(val):
    # [IGNORE: LaTeX 정렬 시스템 위임] 0~4 정수. "행렬 사이 간격을 어떻게 띄울 것인지"에 대한 Word의 규칙 세트 매핑 번호.
    # 왜 무시하는가? -> bmatrix 등 LaTeX 행렬 환경 내부의 고유 Grid Spacing이 수학적 타이포그래피 표준을 따르므로 변경할 필요가 없어서 생략.
    return int(val) if val else None

def parse_m_ST_UnSignedInteger(val):
    # [IGNORE: Word 수동 비율 제어] 강제 여백의 크기나 비율. (위와 동일한 이유로 무시)
    return int(val) if val else None

def parse_m_ST_Char(val):
    # 수학 괄호나 기호들(∫, (, {, [ 등). LaTeX 변환에 핵심적이므로 살림.
    return str(val) if val else None

def parse_m_ST_OnOff(val):
    # 활성화 유무 플래그.
    return str(val).lower() in ["1", "true", "on"] if val else False

def parse_m_ST_String(val):
    # 문자열 그대로 반환. (글꼴 이름 등)
    return str(val) if val else ""

def parse_m_ST_XAlign(val):
    # [IGNORE: 구조적 차이] 엘리먼트 가로 정렬(left, right, center).
    # 왜 무시하는가? -> Word는 칸마다 우측정렬/좌측정렬을 커스텀하지만, LaTeX 표준 행렬 구조는 본질적으로 열(Column) 전체 단위의 정렬을 베이스로 한다. 각 칸에 직접 \hfill을 넣으면 코드가 무거워지므로 생략합니다.
    return str(val) if val else None

def parse_m_ST_YAlign(val):
    # [IGNORE: 구조적 차이] 엘리먼트 세로 정렬(top, center, bot). 위와 동일.
    return str(val) if val else None

def parse_m_ST_Shp(val):
    # [IGNORE: LaTeX 자동화 영역] 괄호가 본문 높이에 "맞춰서 늘어날지(match)" "가운데 정렬될지(centered)" 여부.
    # 왜 무시하는가? -> LaTeX에서는 \left 와 \right 명령어가 수식의 높이(height)를 수학적으로 분석하여 괄호 모양과 크기를 늘려주므로 Word의 결정 힌트가 무의미.
    return str(val) if val else None

def parse_m_ST_FType(val):
    # [IGNORE: 구조적 단순화] 분수(Fraction)를 렌더링할 때 가로줄 긋기, 빗금 긋기, 줄 없애기 등 디자인.
    # 왜 무시하는가? -> 표준 \frac 으로 통일 및 단순화 처리합니다.
    return str(val) if val else None

def parse_m_ST_LimLoc(val):
    # [IGNORE: LaTeX 자동화 영역] 극한 값을 적분 기호 아래위로 붙일지, 옆에 첨자로 붙일지 결정.
    # 왜 무시하는가? -> LaTeX 또한 인라인 모드냐 디스플레이 모드냐에 따라 위치를 알아서 옮긴다. 강제로 넘겨줄 수 있으나, LaTeX 기본 엔진 로직에 맡기는 것이 낫다고 판단하여 생략.
    return str(val) if val else None

def parse_m_ST_TopBot(val):
    # 윗줄(\overline)인지 아랫줄(\underline), 또는 위 괄호인지 아래 괄호인지 구분하는 핵심 변수.
    return str(val) if val else None

def parse_m_ST_Script(val):
    # 수학 폰트 서식(\mathbb 등). 
    return str(val) if val else None

def parse_m_ST_Style(val):
    # 굵게, 이탤릭. 
    return str(val) if val else None

def parse_m_ST_Jc(val):
    # [IGNORE: LaTeX 환경 위임] 방정식 전체를 좌측/중앙/우측에 둘지 결정.
    # 왜 무시하는가? -> 수식 변환본이 들어갈 \begin{equation*} 등의 환경 전체가 문서의 TeX 클래스 설정(fleqn 등)에 의해 정렬되므로 개별 수식에서의 지정은 구조상 무의미.
    return str(val) if val else None

def parse_m_ST_TwipsMeasure(val):
    # [IGNORE: 완전한 단위 이질성 및 비합리성] 
    # 왜 무시하는가? (중요) Twips는 '1인치의 1/1440'을 뜻하는 Word 기반의 물리적 절대치 단위입니다. 
    # 해상도와 글꼴 환경에 따라 반응형으로 동작해야하는 TeX 시스템(\quad 띄어쓰기 개념 등)에 물리적인 Twips 측정값을 주입하면 수식의 유연성이 완전히 파괴됩니다. 
    # 아예 이식하는 것 자체가 비합리적이라 무조건 버려야 합니다.
    return str(val) if val else None

def parse_m_ST_BreakBin(val):
    # [IGNORE: LaTeX 알고리즘 위임] 
    # 왜 무시하는가? -> Word에서 사용자가 엔터를 치지 않아도 화면이 좁아 수식이 강제 개행될 때 "어느 연산자 기준"으로 줄바꿈할지 결정한다. 
    # LaTeX는 자체적으로 줄바꿈 패널티와 연산자 룰을 보유하고 있으므로, Word의 규칙을 강제할 필요가 없다.
    return str(val) if val else None

def parse_m_ST_BreakBinSub(val):
    # [IGNORE: 위와 동일] 뺄셈 줄바꿈 규칙
    return str(val) if val else None


# ==============================================================================
# Complex Types (CT) - Value Wrappers 
# 단순 어트리뷰트 트리 파싱을 위한 래퍼 계층. 
# ==============================================================================
def parse_m_CT_Integer255(node): return parse_m_ST_Integer255(get_val(node))
def parse_m_CT_Integer2(node): return parse_m_ST_Integer2(get_val(node))
def parse_m_CT_SpacingRule(node): return parse_m_ST_SpacingRule(get_val(node))
def parse_m_CT_UnSignedInteger(node): return parse_m_ST_UnSignedInteger(get_val(node))
def parse_m_CT_Char(node): return parse_m_ST_Char(get_val(node))
def parse_m_CT_OnOff(node): return parse_m_ST_OnOff(get_val(node))
def parse_m_CT_String(node): return parse_m_ST_String(get_val(node))
def parse_m_CT_XAlign(node): return parse_m_ST_XAlign(get_val(node))
def parse_m_CT_YAlign(node): return parse_m_ST_YAlign(get_val(node))
def parse_m_CT_Shp(node): return parse_m_ST_Shp(get_val(node))
def parse_m_CT_FType(node): return parse_m_ST_FType(get_val(node))
def parse_m_CT_LimLoc(node): return parse_m_ST_LimLoc(get_val(node))
def parse_m_CT_TopBot(node): return parse_m_ST_TopBot(get_val(node))
def parse_m_CT_Script(node): return parse_m_ST_Script(get_val(node))
def parse_m_CT_Style(node): return parse_m_ST_Style(get_val(node))
def parse_m_CT_OMathJc(node): return parse_m_ST_Jc(get_val(node))
def parse_m_CT_TwipsMeasure(node): return parse_m_ST_TwipsMeasure(get_val(node))
def parse_m_CT_BreakBin(node): return parse_m_ST_BreakBin(get_val(node))
def parse_m_CT_BreakBinSub(node): return parse_m_ST_BreakBinSub(get_val(node))


# ==============================================================================
# Element Groups (EG) - Routing & Container Macros
# ==============================================================================
def parse_m_EG_ScriptStyle(node):
    if node is None: return {}
    return {
        'scr': parse_m_CT_Script(get_child(node, 'scr')),
        'sty': parse_m_CT_Style(get_child(node, 'sty'))
    }

def parse_m_EG_OMathMathElements(node):
    if node is None: return ""
    tag = get_tag(node)
    
    dispatch_map = {
        'acc': parse_m_CT_Acc, 'bar': parse_m_CT_Bar, 'box': parse_m_CT_Box,
        'borderBox': parse_m_CT_BorderBox, 'd': parse_m_CT_D, 'eqArr': parse_m_CT_EqArr,
        'f': parse_m_CT_F, 'func': parse_m_CT_Func, 'groupChr': parse_m_CT_GroupChr,
        'limLow': parse_m_CT_LimLow, 'limUpp': parse_m_CT_LimUpp, 'm': parse_m_CT_M,
        'nary': parse_m_CT_Nary, 'phant': parse_m_CT_Phant, 'rad': parse_m_CT_Rad,
        'sPre': parse_m_CT_SPre, 'sSub': parse_m_CT_SSub, 'sSubSup': parse_m_CT_SSubSup,
        'sSup': parse_m_CT_SSup, 'r': parse_m_CT_R
    }
    
    if tag in dispatch_map:
        return dispatch_map[tag](node)
    return ""

def parse_m_EG_OMathElements(node):
    return parse_m_EG_OMathMathElements(node)


# ==============================================================================
# Complex Types (CT) - Properties Map
# ==============================================================================
def parse_m_CT_ManualBreak(node):
    if node is None: return {}
    return {'alnAt': parse_m_ST_Integer255(get_val(node, 'alnAt'))}

def parse_m_CT_CtrlPr(node): 
    # [IGNORE: 순수 어플리케이션(Word) 메타데이터]
    # 이유: 이 속성은 특정 수식을 누가 수정했는지나 단축키 지정 등 '파일 편집 환경의 히스토리'를 저장한 한것.
    # 수식 결과물과는 수학적으로 아무 상관도 없으므로 제외.
    if node is None: return {}
    return {}

def parse_m_CT_RPR(node):
    if node is None: return {}
    res = {
        'lit': parse_m_CT_OnOff(get_child(node, 'lit')),
        'nor': parse_m_CT_OnOff(get_child(node, 'nor')),
        'brk': parse_m_CT_ManualBreak(get_child(node, 'brk')),
        'aln': parse_m_CT_OnOff(get_child(node, 'aln'))
    }
    res.update(parse_m_EG_ScriptStyle(node))
    return res

def parse_m_CT_OMathArgPr(node): 
    if node is None: return {}
    return {'argSz': parse_m_CT_Integer2(get_child(node, 'argSz'))}

def parse_m_CT_AccPr(node):
    if node is None: return {}
    return {
        'chr': parse_m_CT_Char(get_child(node, 'chr')),
        'ctrlPr': parse_m_CT_CtrlPr(get_child(node, 'ctrlPr'))
    }

def parse_m_CT_BarPr(node):
    if node is None: return {}
    return {
        'pos': parse_m_CT_TopBot(get_child(node, 'pos')),
        'ctrlPr': parse_m_CT_CtrlPr(get_child(node, 'ctrlPr'))
    }

def parse_m_CT_BoxPr(node):
    if node is None: return {}
    return {
        'opEmu': parse_m_CT_OnOff(get_child(node, 'opEmu')),
        'noBreak': parse_m_CT_OnOff(get_child(node, 'noBreak')),
        'diff': parse_m_CT_OnOff(get_child(node, 'diff')),
        'brk': parse_m_CT_ManualBreak(get_child(node, 'brk')),
        'aln': parse_m_CT_OnOff(get_child(node, 'aln')),
        'ctrlPr': parse_m_CT_CtrlPr(get_child(node, 'ctrlPr'))
    }

def parse_m_CT_BorderBoxPr(node):
    # [IGNORE: 변환 복잡성 및 출력 품질 관리] 테두리 박스의 선을 일부만 지우거나 대각선으로 취소선을 긋는 등 Word의 자유도 높은 그리기 옵션.
    # 이유: hideTop(윗테두리 가리기), strikeTLBR(대각선으로 취소선 긋기) 등 Word의 그리기 도구.
    # LaTeX로 완전한 이식이 매우 어렵고, 수학적 표현의 명확성에도 크게 기여하지 않으므로 일괄적으로 무시.
    if node is None: return {}
    return {
        'hideTop': parse_m_CT_OnOff(get_child(node, 'hideTop')),
        'hideBot': parse_m_CT_OnOff(get_child(node, 'hideBot')),
        'hideLeft': parse_m_CT_OnOff(get_child(node, 'hideLeft')),
        'hideRight': parse_m_CT_OnOff(get_child(node, 'hideRight')),
        'strikeH': parse_m_CT_OnOff(get_child(node, 'strikeH')),
        'strikeV': parse_m_CT_OnOff(get_child(node, 'strikeV')),
        'strikeBLTR': parse_m_CT_OnOff(get_child(node, 'strikeBLTR')),
        'strikeTLBR': parse_m_CT_OnOff(get_child(node, 'strikeTLBR')),
        'ctrlPr': parse_m_CT_CtrlPr(get_child(node, 'ctrlPr'))
    }

def parse_m_CT_DPr(node):
    if node is None: return {}
    return {
        'begChr': parse_m_CT_Char(get_child(node, 'begChr')),
        'sepChr': parse_m_CT_Char(get_child(node, 'sepChr')),
        'endChr': parse_m_CT_Char(get_child(node, 'endChr')),
        'grow': parse_m_CT_OnOff(get_child(node, 'grow')),
        'shp': parse_m_CT_Shp(get_child(node, 'shp')),
        'ctrlPr': parse_m_CT_CtrlPr(get_child(node, 'ctrlPr'))
    }

def parse_m_CT_EqArrPr(node):
    if node is None: return {}
    return {
        'baseJc': parse_m_CT_YAlign(get_child(node, 'baseJc')),
        'maxDist': parse_m_CT_OnOff(get_child(node, 'maxDist')),
        'objDist': parse_m_CT_OnOff(get_child(node, 'objDist')),
        'rSpRule': parse_m_CT_SpacingRule(get_child(node, 'rSpRule')),
        'rSp': parse_m_CT_UnSignedInteger(get_child(node, 'rSp')),
        'ctrlPr': parse_m_CT_CtrlPr(get_child(node, 'ctrlPr'))
    }

def parse_m_CT_FPr(node):
    if node is None: return {}
    return {
        'type': parse_m_CT_FType(get_child(node, 'type')),
        'ctrlPr': parse_m_CT_CtrlPr(get_child(node, 'ctrlPr'))
    }

def parse_m_CT_FuncPr(node):
    if node is None: return {}
    return {'ctrlPr': parse_m_CT_CtrlPr(get_child(node, 'ctrlPr'))}
    
def parse_m_CT_GroupChrPr(node):
    if node is None: return {}
    return {
        'chr': parse_m_CT_Char(get_child(node, 'chr')),
        'pos': parse_m_CT_TopBot(get_child(node, 'pos')),
        'vertJc': parse_m_CT_TopBot(get_child(node, 'vertJc')),
        'ctrlPr': parse_m_CT_CtrlPr(get_child(node, 'ctrlPr'))
    }

def parse_m_CT_LimLowPr(node):
    if node is None: return {}
    return {'ctrlPr': parse_m_CT_CtrlPr(get_child(node, 'ctrlPr'))}

def parse_m_CT_LimUppPr(node):
    if node is None: return {}
    return {'ctrlPr': parse_m_CT_CtrlPr(get_child(node, 'ctrlPr'))}

def parse_m_CT_MCPr(node):
    if node is None: return {}
    return {
        'count': parse_m_CT_Integer255(get_child(node, 'count')),
        'mcJc': parse_m_CT_XAlign(get_child(node, 'mcJc'))
    }

def parse_m_CT_MPr(node):
    if node is None: return {}
    return {
        'baseJc': parse_m_CT_YAlign(get_child(node, 'baseJc')),
        'plcHide': parse_m_CT_OnOff(get_child(node, 'plcHide')),
        'rSpRule': parse_m_CT_SpacingRule(get_child(node, 'rSpRule')),
        'cGpRule': parse_m_CT_SpacingRule(get_child(node, 'cGpRule')),
        'rSp': parse_m_CT_UnSignedInteger(get_child(node, 'rSp')),
        'cSp': parse_m_CT_UnSignedInteger(get_child(node, 'cSp')),
        'cGp': parse_m_CT_UnSignedInteger(get_child(node, 'cGp')),
        'mcs': parse_m_CT_MCS(get_child(node, 'mcs')),
        'ctrlPr': parse_m_CT_CtrlPr(get_child(node, 'ctrlPr'))
    }

def parse_m_CT_NaryPr(node):
    if node is None: return {}
    return {
        'chr': parse_m_CT_Char(get_child(node, 'chr')),
        'limLoc': parse_m_CT_LimLoc(get_child(node, 'limLoc')),
        'grow': parse_m_CT_OnOff(get_child(node, 'grow')),
        'subHide': parse_m_CT_OnOff(get_child(node, 'subHide')),
        'supHide': parse_m_CT_OnOff(get_child(node, 'supHide')),
        'ctrlPr': parse_m_CT_CtrlPr(get_child(node, 'ctrlPr'))
    }

def parse_m_CT_PhantPr(node):
    if node is None: return {}
    return {
        'show': parse_m_CT_OnOff(get_child(node, 'show')),
        'zeroWid': parse_m_CT_OnOff(get_child(node, 'zeroWid')),
        'zeroAsc': parse_m_CT_OnOff(get_child(node, 'zeroAsc')),
        'zeroDesc': parse_m_CT_OnOff(get_child(node, 'zeroDesc')),
        'transp': parse_m_CT_OnOff(get_child(node, 'transp')),
        'ctrlPr': parse_m_CT_CtrlPr(get_child(node, 'ctrlPr'))
    }
    
def parse_m_CT_RadPr(node):
    if node is None: return {}
    return {
        'degHide': parse_m_CT_OnOff(get_child(node, 'degHide')),
        'ctrlPr': parse_m_CT_CtrlPr(get_child(node, 'ctrlPr'))
    }

def parse_m_CT_SPrePr(node):
    if node is None: return {}
    return {'ctrlPr': parse_m_CT_CtrlPr(get_child(node, 'ctrlPr'))}
def parse_m_CT_SSubPr(node):
    if node is None: return {}
    return {'ctrlPr': parse_m_CT_CtrlPr(get_child(node, 'ctrlPr'))}
def parse_m_CT_SSupPr(node):
    if node is None: return {}
    return {'ctrlPr': parse_m_CT_CtrlPr(get_child(node, 'ctrlPr'))}

def parse_m_CT_SSubSupPr(node):
    if node is None: return {}
    return {
        'alnScr': parse_m_CT_OnOff(get_child(node, 'alnScr')), 
        'ctrlPr': parse_m_CT_CtrlPr(get_child(node, 'ctrlPr'))
    }

def parse_m_CT_OMathParaPr(node):
    if node is None: return {}
    return {'jc': parse_m_CT_OMathJc(get_child(node, 'jc'))}

def parse_m_CT_MathPr(node): 
    if node is None: return {}
    
    wrapIndent = get_child(node, 'wrapIndent')
    wrapRight = get_child(node, 'wrapRight')
    
    return {
        'mathFont': parse_m_CT_String(get_child(node, 'mathFont')),
        'brkBin': parse_m_CT_BreakBin(get_child(node, 'brkBin')),
        'brkBinSub': parse_m_CT_BreakBinSub(get_child(node, 'brkBinSub')),
        'smallFrac': parse_m_CT_OnOff(get_child(node, 'smallFrac')),
        'dispDef': parse_m_CT_OnOff(get_child(node, 'dispDef')),
        'lMargin': parse_m_CT_TwipsMeasure(get_child(node, 'lMargin')),
        'rMargin': parse_m_CT_TwipsMeasure(get_child(node, 'rMargin')),
        'defJc': parse_m_CT_OMathJc(get_child(node, 'defJc')),
        'preSp': parse_m_CT_TwipsMeasure(get_child(node, 'preSp')),
        'postSp': parse_m_CT_TwipsMeasure(get_child(node, 'postSp')),
        'interSp': parse_m_CT_TwipsMeasure(get_child(node, 'interSp')),
        'intraSp': parse_m_CT_TwipsMeasure(get_child(node, 'intraSp')),
        'wrapIndent': parse_m_CT_TwipsMeasure(wrapIndent) if wrapIndent is not None else None,
        'wrapRight': parse_m_CT_OnOff(wrapRight) if wrapRight is not None else None,
        'intLim': parse_m_CT_LimLoc(get_child(node, 'intLim')),
        'naryLim': parse_m_CT_LimLoc(get_child(node, 'naryLim'))
    }

# ==============================================================================
# Complex Types (CT) - AST to LaTeX Generation
# ==============================================================================
# [설명: 연산자 속성(~~Pr) 변수들을 변환 구문에서 의도적으로 무시(Ignore)하는 이유]
# AST 생성 단계에서 DPr, FPr, MPr, ArgPr 등 수많은 속성(Pr) 딕셔너리를 파싱해 가져오지만,
# 실제 아래의 함수들에서 LaTeX 문자열을 조립(Return)할 때는 이 속성값 대부분을 안 쓴ㄷ다
# 워드의 'Pr'들은 화면의 픽셀 간격, 요소 정렬(alnScr), 특정 테두리 숨김(hideTop),
# 문자 강제 성장(grow/shp) 등 완벽하게 'GUI 워드프로세서 렌더링용 시각 지시자'이기 때문이다. 
# 
# 반면 LaTeX은 \frac, \begin{bmatrix}, \left( \right), ^, _ 등의 명령어 자체에
# 전문적인 타이포그래피/간격/배치 알고리즘이 내장되어 있다. 이 내장 렌더링을
# 따르는 것이 수학적으로 더 정확하고 코드가 깔끔해진다고 판단하여, Word의 시각적 지시자들은 대부분 무시하고 LaTeX의 기본 렌더링 엔진에 맡기는 전략.
# (단, 괄호기호의 종류(begChr)나, 적분기호 종류(chr) 같이 '수학적 본질'을 결정하는 속성은 사용합니다.)
# ==============================================================================
def parse_m_CT_OMathArg(node):
    if node is None: return ""
    # [IGNORE: argPr] 첨자의 크기를 강제로 키우고 줄이는(argSz) 물리적 비율 속성 무시
    argPr = parse_m_CT_OMathArgPr(get_child(node, 'argPr')) 
    
    content = ""
    for child in node:
        tag = get_tag(child)
        if tag not in ('argPr', 'ctrlPr'): 
            content += parse_m_EG_OMathElements(child)
    return content

def parse_m_CT_Text(node):
    if node is None: return ""
    return node.text if node.text else ""

def parse_m_CT_R(node):
    rPr = parse_m_CT_RPR(get_child(node, 'rPr'))
    
    content = ""
    for t_node in get_children(node, 't'):
        raw_text = parse_m_CT_Text(t_node)
        content += cleanup_text(raw_text)

    scr = rPr.get('scr')
    if scr == 'fraktur': content = f"\\mathfrak{{{content}}}"
    elif scr == 'double-struck': content = f"\\mathbb{{{content}}}"
    elif scr == 'script': content = f"\\mathcal{{{content}}}"
    elif scr == 'sans-serif': content = f"\\textsf{{{content}}}"
    elif scr == 'monospace': content = f"\\mathtt{{{content}}}"

    sty = rPr.get('sty')
    if sty == 'b': content = f"\\mathbf{{{content}}}"
    elif sty == 'i': content = f"\\mathit{{{content}}}"
    elif sty == 'bi': content = f"\\boldsymbol{{{content}}}"
    
    return content

def parse_m_CT_F(node):
    # [IGNORE: fPr] 분수 가로줄의 형태(빗금 모양, 수평선 긋기 여부 등) 시각 설정 무시. 표준 \frac 사용
    fPr = parse_m_CT_FPr(get_child(node, 'fPr')) 
    num = parse_m_CT_OMathArg(get_child(node, 'num'))
    den = parse_m_CT_OMathArg(get_child(node, 'den'))
    return f"\\frac{{{num}}}{{{den}}}"

def parse_m_CT_Rad(node):
    # [IGNORE: radPr] 차수 숨기기(degHide) 등 어플리케이션 옵션 무시. deg(차수) 노드 유무만으로 분기
    radPr = parse_m_CT_RadPr(get_child(node, 'radPr'))
    deg = parse_m_CT_OMathArg(get_child(node, 'deg'))
    e = parse_m_CT_OMathArg(get_child(node, 'e'))
    return f"\\sqrt[{deg}]{{{e}}}" if deg.strip() else f"\\sqrt{{{e}}}"

def parse_m_CT_SSup(node):
    # [IGNORE: sSupPr] 위첨자 메타데이터 무시 (위치는 LaTeX 자동 배치 엔진에 위임)
    sSupPr = parse_m_CT_SSupPr(get_child(node, 'sSupPr'))
    e = parse_m_CT_OMathArg(get_child(node, 'e'))
    sup = parse_m_CT_OMathArg(get_child(node, 'sup'))
    return f"{{{e}}}^{{{sup}}}"

def parse_m_CT_SSub(node):
    # [IGNORE: sSubPr] 아래첨자 메타데이터 무시
    sSubPr = parse_m_CT_SSubPr(get_child(node, 'sSubPr'))
    e = parse_m_CT_OMathArg(get_child(node, 'e'))
    sub = parse_m_CT_OMathArg(get_child(node, 'sub'))
    return f"{{{e}}}_{{{sub}}}"

def parse_m_CT_SSubSup(node):
    # [IGNORE: sSubSupPr] 첨자 간 상대적 위치 정렬(alnScr) 무시. LaTeX 기본 커닝(Kerning) 사용
    sSubSupPr = parse_m_CT_SSubSupPr(get_child(node, 'sSubSupPr'))
    e = parse_m_CT_OMathArg(get_child(node, 'e'))
    sub = parse_m_CT_OMathArg(get_child(node, 'sub'))
    sup = parse_m_CT_OMathArg(get_child(node, 'sup'))
    return f"{{{e}}}_{{{sub}}}^{{{sup}}}"

def parse_m_CT_SPre(node):
    # [IGNORE: sPrePr] 전위첨자 메타데이터 무시
    # 좌측 위, 좌측 아래 첨자.
    sPrePr = parse_m_CT_SPrePr(get_child(node, 'sPrePr'))
    e = parse_m_CT_OMathArg(get_child(node, 'e'))
    sub = parse_m_CT_OMathArg(get_child(node, 'sub'))
    sup = parse_m_CT_OMathArg(get_child(node, 'sup'))
    return f"{{}}_{{{sub}}}^{{{sup}}}{e}"

def parse_m_CT_Nary(node):
    naryPr = parse_m_CT_NaryPr(get_child(node, 'naryPr'))
    chr_val = naryPr.get('chr') or '∫'
    
    char_map = {'∫': '\\int', '∑': '\\sum', '∏': '\\prod', '∐': '\\coprod'}
    latex_char = char_map.get(chr_val, '\\int')
    
    sub = parse_m_CT_OMathArg(get_child(node, 'sub'))
    sup = parse_m_CT_OMathArg(get_child(node, 'sup'))
    e = parse_m_CT_OMathArg(get_child(node, 'e'))
    
    res = latex_char
    if sub: res += f"_{{{sub}}}"
    if sup: res += f"^{{{sup}}}"
    return f"{res} {{{e}}}"

def parse_m_CT_D(node):
    dPr = parse_m_CT_DPr(get_child(node, 'dPr'))
    # [설명: DPr 사용 방식] begChr(여는괄호), endChr(닫는괄호), sepChr(구분자)처럼 핵심 '수학 데이터'는 꺼내 쓰지만,
    # [IGNORE: DPr 나머지] grow(본문 크기에 맞춤), shp(가운데 정렬) 등 모양 치장용 속성은 무시 (\left, \right가 알아서 최적화함)
    beg = dPr.get('begChr') or "("
    end = dPr.get('endChr') or ")"
    sep = dPr.get('sepChr') or "|"
    
    args = [parse_m_CT_OMathArg(e) for e in get_children(node, 'e')]
    content = f" {sep} ".join(args)
    left = f"\\left{beg}" if beg not in ("{", "}") else f"\\left\\{beg}"
    right = f"\\right{end}" if end not in ("{", "}") else f"\\right\\{end}"
    return f"{left} {content} {right}"

def parse_m_CT_Acc(node):
    # AccPr의 핵심인 악센트 기호 문자(chr)는 가져오지만, 에디터 단축키 정보(ctrlPr) 등은 무시
    accPr = parse_m_CT_AccPr(get_child(node, 'accPr'))
    chr_val = accPr.get('chr', '\u0302') 
    e = parse_m_CT_OMathArg(get_child(node, 'e'))
    
    acc_map = {'\u0302': '\\hat', '\u0300': '\\grave', '\u0303': '\\tilde', '\u0304': '\\bar', '\u0307': '\\dot', '\u0308': '\\ddot', '\u20D7': '\\vec'}
    acc_cmd = acc_map.get(chr_val, '\\hat')
    return f"{acc_cmd}{{{e}}}"

def parse_m_CT_Bar(node):
    barPr = parse_m_CT_BarPr(get_child(node, 'barPr'))
    pos = barPr.get('pos') or 'top'
    e = parse_m_CT_OMathArg(get_child(node, 'e'))
    return f"\\underline{{{e}}}" if pos == 'bot' else f"\\overline{{{e}}}"

def parse_m_CT_MC(node):
    # [IGNORE: mcPr] 열의 개수(count)나 가로 지정 정렬(mcJc) 무시. LaTeX 앰퍼샌드(&) 환경이 자동 결정
    mcPr = parse_m_CT_MCPr(get_child(node, 'mcPr'))
    return mcPr

def parse_m_CT_MCS(node):
    return [parse_m_CT_MC(mc) for mc in get_children(node, 'mc')]

def parse_m_CT_M(node):
    # [IGNORE: mPr] 열/행 간격(cSp, cGp, rSp) 및 베이스라인 정렬 등.
    # 워드의 absolute 수치를 강제하는 것보다 bmatrix 환경의 자체 그리드 알고리즘에 맡기기로 결정.
    mPr = parse_m_CT_MPr(get_child(node, 'mPr'))
    rows = [parse_m_CT_MR(mr) for mr in get_children(node, 'mr')]
    return "\\begin{bmatrix}\n" + " \\\\\n".join(rows) + "\n\\end{bmatrix}"

def parse_m_CT_MR(node):
    cols = [parse_m_CT_OMathArg(e) for e in get_children(node, 'e')]
    return " & ".join(cols)

def parse_m_CT_EqArr(node):
    # [IGNORE: eqArrPr] 방정식 배열의 줄간격 룰(rSpRule) 및 여백 옵션 무시. aligned 환경 기본값 사용
    eqArrPr = parse_m_CT_EqArrPr(get_child(node, 'eqArrPr'))
    args = [parse_m_CT_OMathArg(e) for e in get_children(node, 'e')]
    return "\\begin{aligned}\n" + " \\\\\n".join(args) + "\n\\end{aligned}"

def parse_m_CT_Func(node):
    funcPr = parse_m_CT_FuncPr(get_child(node, 'funcPr'))
    fName = parse_m_CT_OMathArg(get_child(node, 'fName')).strip()
    e = parse_m_CT_OMathArg(get_child(node, 'e'))
    
    # fName 내부에 \u2061 같은 보이지 않는 문자가 이미 제거된 상태
    # 기본 제공되는 표준 수학 함수들 매핑
    standard_funcs = ['sin', 'cos', 'tan', 'csc', 'sec', 'cot', 'sinh', 'cosh', 'tanh', 'csch', 'sech', 'coth', 'arcsin', 'arccos', 'arctan', 'lim', 'log', 'ln', 'min', 'max']
    
    # fName 정리 ('sin', 'lim'만 추출하기 위함)
    clean_fName = fName.replace(' ', '').replace('{', '').replace('}', '')
    
    # 특수한 경우: fName 이 첨자를 포함할 때 (예: \lim_{n \to \infty})
    # 이 부분은 fName 자체가 복잡할 수 있으므로, 단순 문자열 변환보다는 기본 매핑
    if clean_fName in standard_funcs:
        return f"\\{clean_fName} {{{e}}}"
    elif fName.startswith('lim') or fName.startswith('max') or fName.startswith('min'):
        # lim 아래에 첨자가 달린 형태가 fName 쪽으로 파싱되어 들어온 경우 대비
        return f"\\mathop{{{fName}}} {{{e}}}"
        
    return f"\\mathop{{{fName}}} {{{e}}}"

def parse_m_CT_GroupChr(node):
    groupChrPr = parse_m_CT_GroupChrPr(get_child(node, 'groupChrPr'))
    pos = groupChrPr.get('pos')
    chr_val = groupChrPr.get('chr')
    e = parse_m_CT_OMathArg(get_child(node, 'e'))
    if pos == 'bot' or chr_val == '\u23DF': return f"\\underbrace{{{e}}}" 
    return f"\\overbrace{{{e}}}"

def parse_m_CT_LimLow(node):
    limLowPr = parse_m_CT_LimLowPr(get_child(node, 'limLowPr'))
    e = parse_m_CT_OMathArg(get_child(node, 'e'))
    lim = parse_m_CT_OMathArg(get_child(node, 'lim'))
    return f"\\mathop{{{e}}}\\limits_{{{lim}}}"

def parse_m_CT_LimUpp(node):
    limUppPr = parse_m_CT_LimUppPr(get_child(node, 'limUppPr'))
    e = parse_m_CT_OMathArg(get_child(node, 'e'))
    lim = parse_m_CT_OMathArg(get_child(node, 'lim'))
    return f"\\mathop{{{e}}}\\limits^{{{lim}}}"

def parse_m_CT_Box(node):
    boxPr = parse_m_CT_BoxPr(get_child(node, 'boxPr'))
    return parse_m_CT_OMathArg(get_child(node, 'e'))

def parse_m_CT_BorderBox(node): 
    # [IGNORE 설명 참조 (BorderBoxPr 부분)]
    borderBoxPr = parse_m_CT_BorderBoxPr(get_child(node, 'borderBoxPr'))
    return f"\\boxed{{{parse_m_CT_OMathArg(get_child(node, 'e'))}}}"

def parse_m_CT_Phant(node): 
    phantPr = parse_m_CT_PhantPr(get_child(node, 'phantPr'))
    return f"\\phantom{{{parse_m_CT_OMathArg(get_child(node, 'e'))}}}"

# ==============================================================================
# Roots Definitions
# ==============================================================================
def parse_m_CT_OMath(node):
    return "".join(parse_m_EG_OMathElements(child) for child in node)

def parse_m_CT_OMathPara(node):
    oMathParaPr = parse_m_CT_OMathParaPr(get_child(node, 'oMathParaPr'))
    content = "".join(parse_m_CT_OMath(omath) for omath in get_children(node, 'oMath'))
    return f"$$\n{content}\n$$"

def parse_m_CT_MathPr_Root(node):
    return parse_m_CT_MathPr(node)

def parse_omml_to_latex(node):
    tag = get_tag(node)
    if tag == 'mathPr': return parse_m_CT_MathPr_Root(node)
    elif tag == 'oMathPara': return parse_m_CT_OMathPara(node)
    elif tag == 'oMath': return parse_m_CT_OMath(node)
    return ""


def convert_omml_to_latex(node):
    """Public alias for converting an OMML element into LaTeX."""
    return parse_omml_to_latex(node)


def parse_omml_xml(xml: str):
    """Parse an OMML XML string and convert its root element into LaTeX."""
    return parse_omml_to_latex(ET.fromstring(xml))

if __name__ == "__main__":
    sample_xml = """<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <m:oMath>
            <m:r>
                <m:rPr><m:scr m:val="double-struck"/><m:sty m:val="b"/></m:rPr>
                <m:t>&#x2061;R&#x2062;α&#x221E;</m:t>
            </m:r>
        </m:oMath>
    </m:oMathPara>"""
    import xml.etree.ElementTree as tree
    print(parse_omml_to_latex(tree.fromstring(sample_xml)))
