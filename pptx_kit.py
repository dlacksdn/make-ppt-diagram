"""
pptx_kit.py — RILS LAB 세미나 다이어그램 키트 (편집 가능한 PowerPoint 네이티브 도형만 사용)

원칙
- 베이스 = 템플릿 6페이지를 통째로 복제 (모서리 장식·로고·바닥글·쪽번호·회색 베이스 포함)
- 상자 = 사각형 + 별도 텍스트상자를 그룹으로 묶음 (여백을 따로 조절하려고)
- 수식 = PowerPoint 수식 개체(OMML), 아이콘 = 원본과 같은 SVG 그래픽
- 모든 수치는 아래 상수에 명시한다. 테마 기본값을 물려받지 않는다.

좌표 단위는 inch. 슬라이드 = 20 x 11.25 in.
"""
import copy, os, re
from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from PIL import ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
EMU = 914400

# ───────────────────────── 확정 수치 ─────────────────────────
BASE_SLIDE_INDEX = 6                      # 템플릿에서 복제할 슬라이드 (1-based)
BASE = (1.083, 1.333, 17.833, 8.542)      # 회색 베이스 (x, y, w, h) — 템플릿 6p 그대로
MARGIN = 0.37                             # 베이스 테두리↔패널, 패널↔패널, 패널↔노드 간격

PT_FIG_TITLE = 24     # 그림 대표 제목 (Building the world model)
PT_SECTION = 20       # 구획 라벨 (빨강 #C00000)
PT_BOX_TITLE = 18     # 상자 제목
PT_BOX_SUB = 14       # 상자 부가 텍스트
PT_EDGE = 12          # 화살표 라벨
PT_VERDICT = 24       # 하단 총평 문장

W_SHAPE = 2.25        # 모든 도형 윤곽선 (pt)
W_LINE = 1.5          # 모든 화살표·선 (pt)
# 줄글 금지 — 핵심만 간결하게 (lint가 저장 시 검사)
MAX_WORDS_PER_LINE = 6     # 상자 제목·부가·주석 한 줄 최대 단어 수 (수식 제외)
MAX_LINES = 2              # 텍스트상자 하나의 최대 줄 수
MAX_VERDICT_WORDS = 12     # 총평 문장
MAX_TEXT_PIECES = 35       # 슬라이드당 텍스트상자 수 (넘으면 슬라이드 분할)
FONT = "Calibri"
MATH_FONT = "Cambria Math"
RED = "C00000"

# 역할 → 테마 색 (schemeClr, lumMod, lumOff). HEX는 Office 2007 테마 기준 참고값
FILL = {
    "white":    ("bg1", None, None),          # FFFFFF  노드 기본
    "panel":    ("bg1", 95000, None),         # F2F2F2  서브패널
    "base":     ("bg1", 85000, None),         # D9D9D9  베이스
    "raw":      ("bg2", None, None),          # EEECE1  원시 관측 (RGB-D video)
    "input":    ("accent4", 40000, 60000),    # CCC1DA  입력 데이터셋 패널
    "input_row":("accent4", 20000, 80000),    # E6E0EC  데이터셋 행 / 큰 연산 묶음
    "loss":     ("accent5", 20000, 80000),    # DBEEF4  손실·목적함수
    "sky_group":("accent5", 40000, 60000),    # B7DEE8  하늘 계열 그룹
    "trainable":("accent2", 20000, 80000),    # F2DCDB  🔥 학습 파라미터 / 최종 실행 결과
    "output":   ("accent3", 20000, 80000),    # EBF1DE  산출물
    "olive_group":("accent3", 40000, 60000),  # D7E4BD  올리브 계열 그룹
    "control":  ("accent1", 20000, 80000),    # DCE6F2  제어 입력 (keyboard/mouse)
}
MUTED = "BFBFBF"   # 의도적 미사용 요소: 지우지 말고 회색 글자로 남긴다 (예: s : proprioception)
# 도형 어휘: 종류 = 의미
SHAPE = {
    "rect": "rect",                  # 연산·모듈·객체 (기본)
    "store": "can",                  # 데이터 저장소 (Replay buffer) — 원통
    "branch": "flowChartDecision",   # 조건 분기 (pass/fail 라벨 필수) — 마름모
    "encoder": "trapezoid",          # 인코더·압축 — 사다리꼴
    "var": "ellipse",                # 확률변수·샘플 — 작은 원
}
ICON = {  # 파일, 크기(in) — 원본 실측
    "fire": ("fire", 0.26), "snow": ("snowflake", 0.33), "eye": ("eye", 0.25),
}

# ───────────────────────── 네임스페이스 ─────────────────────────
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "asvg": "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
}
NSDECL = " ".join(f'xmlns:{k}="{v}"' for k, v in NS.items())

def X(el, expr): return etree.XPath(expr, namespaces=NS)(el)
def E(v): return int(round(v * EMU))
def esc(t): return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
def frag(xml):
    """네임스페이스 선언을 붙여 XML 조각을 파싱"""
    root = etree.fromstring(f"<root {NSDECL}>{xml}</root>")
    return list(root)

STYLE_SHAPE = ('<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef>'
               '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
               '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>')
STYLE_LINE = ('<p:style><a:lnRef idx="1"><a:schemeClr val="dk1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="dk1"/></a:fillRef>'
              '<a:effectRef idx="0"><a:schemeClr val="dk1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>')
EMPTY_TX = '<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p></p:txBody>'

def fill_xml(role):
    if role is None: return "<a:noFill/>"
    if isinstance(role, str) and re.fullmatch(r"[0-9A-Fa-f]{6}", role):
        return f'<a:solidFill><a:srgbClr val="{role}"/></a:solidFill>'
    clr, mod, off = FILL[role]
    inner = (f'<a:lumMod val="{mod}"/>' if mod else "") + (f'<a:lumOff val="{off}"/>' if off else "")
    return f'<a:solidFill><a:schemeClr val="{clr}">{inner}</a:schemeClr></a:solidFill>'

# ───────────────────────── 글자 폭 측정 (Carlito = Calibri 호환 메트릭) ─────────────────────────
_FONTS = {}
def _font(bold, italic, pt, math=False):
    fam = "Caladea" if math else "Carlito"
    style = {(0,0):"Regular",(1,0):"Bold",(0,1):"Italic",(1,1):"BoldItalic"}[(int(bold), int(italic))]
    key = (fam, style, pt)
    if key not in _FONTS:
        for d in (os.path.join(HERE, "assets", "fonts"), "/usr/share/fonts/truetype/crosextra"):
            f = os.path.join(d, f"{fam}-{style}.ttf")
            if os.path.exists(f): _FONTS[key] = ImageFont.truetype(f, int(pt * 10)); break
        else:
            raise FileNotFoundError("Carlito/Caladea 폰트가 없습니다 → assets/fonts/ 확인")
    return _FONTS[key]

def text_width(segs, pt):
    """segs: [(kind, text, bold)] → inch"""
    w = 0
    for kind, t, b in segs:
        if kind == "math":
            w += _font(b, True, pt, True).getlength(_plain(t)) / 10 / 72 * 1.05
        else:
            w += _font(b, False, pt).getlength(t) / 10 / 72
    return w

# ───────────────────────── 미니 수식 (LaTeX 일부 → OMML) ─────────────────────────
GREEK = dict(alpha="α", beta="β", gamma="γ", delta="δ", epsilon="ε", varepsilon="ε", theta="θ", lambda_="λ", mu="μ",
             pi="π", rho="ρ", sigma="σ", tau="τ", phi="ϕ", psi="ψ", omega="ω", eta="η", xi="ξ", zeta="ζ", kappa="κ", nu="ν", chi="χ")
SYM = dict(cdot="·", in_="∈", sim="∼", ell="ℓ", times="×", odot="⊙", to="→", rightarrow="→", leftarrow="←", approx="≈",
           le="≤", ge="≥", neq="≠", infty="∞", ldots="…", dots="…", nabla="∇", partial="∂", sum="∑", prod="∏",
           Delta="Δ", Theta="Θ", Lambda="Λ", Sigma="Σ", Phi="Φ", Psi="Ψ", Omega="Ω", mathcal_M="ℳ")

def _ital(ch, bold):
    o = ord(ch)
    if "A" <= ch <= "Z": return chr((0x1D468 if bold else 0x1D434) + o - 65)
    if "a" <= ch <= "z":
        if ch == "h" and not bold: return "ℎ"
        return chr((0x1D482 if bold else 0x1D44E) + o - 97)
    if 0x3B1 <= o <= 0x3C9: return chr((0x1D736 if bold else 0x1D6FC) + o - 0x3B1)
    return ch

def _tokens(s):
    i, out = 0, []
    while i < len(s):
        c = s[i]
        if c == "\\":
            m = re.match(r"\\([A-Za-z]+)", s[i:]); name = m.group(1); i += len(m.group(0))
            out.append(("cmd", name))
        elif c in "{}_^": out.append((c, c)); i += 1
        elif c == " ": i += 1
        else: out.append(("ch", c)); i += 1
    return out

def _parse(toks, i=0, stop=None):
    """→ list of nodes: ('r', text) | ('sub', base, sub) | ('sup',..) | ('subsup',..) | ('hat', body)"""
    nodes = []
    def atom(i):
        k, v = toks[i]
        if k == "{":
            body, i = _parse(toks, i + 1, "}"); return body, i + 1
        if k == "cmd":
            if v in ("hat", "tilde", "bar", "dot"):
                body, j = atom(i + 1); return [("acc", v, body)], j
            if v in ("mathbf", "boldsymbol"):
                body, j = atom(i + 1); return [("bold", body)], j
            if v == "mathcal":
                body, j = atom(i + 1); ch = body[0][1] if body and body[0][0]=="r" else "M"
                return [("r", {"M":"ℳ","L":"ℒ","D":"𝒟","N":"𝒩"}.get(ch, ch), True)], j
            g = GREEK.get(v) or GREEK.get(v + "_")
            if g: return [("r", g, False)], i + 1
            sy = SYM.get(v) or SYM.get(v + "_")
            if sy: return [("r", sy, True)], i + 1
            return [("r", v, True)], i + 1
        return [("r", v, False)], i + 1
    while i < len(toks):
        k, v = toks[i]
        if stop and k == stop: return nodes, i
        if k in "_^":
            base = [nodes.pop()] if nodes else [("r", " ", True)]
            arg, i = atom(i + 1)
            if base[0][0] in ("sub", "sup") and base[0][0] != {"_":"sub","^":"sup"}[k] and len(base[0]) == 3:
                b = base[0]
                sub, sup = (b[2], arg) if b[0] == "sub" else (arg, b[2])
                nodes.append(("subsup", b[1], sub, sup)); continue
            nodes.append(({"_": "sub", "^": "sup"}[k], base, arg)); continue
        a, i = atom(i); nodes.extend(a)
    return nodes, i

def _plain(tex):
    nodes, _ = _parse(_tokens(tex))
    def f(ns):
        o = ""
        for n in ns:
            if n[0] == "r": o += n[1]
            elif n[0] in ("sub", "sup"): o += f(n[1]) + f(n[2])
            elif n[0] == "subsup": o += f(n[1]) + f(n[2]) + f(n[3])
            elif n[0] == "acc": o += f(n[2])
            elif n[0] == "bold": o += f(n[1])
        return o
    return f(nodes)

def _flat(nodes, bl=0):
    out = []
    for n in nodes:
        if n[0] == "r": out.append((n[1], bl))
        elif n[0] == "sub": out += _flat(n[1], bl) + _flat(n[2], -25000)
        elif n[0] == "sup": out += _flat(n[1], bl) + _flat(n[2], 30000)
        elif n[0] == "subsup": out += _flat(n[1], bl) + _flat(n[2], -25000) + _flat(n[3], 30000)
        elif n[0] == "acc": out += _flat(n[2], bl)
        elif n[0] == "bold": out += _flat(n[1], bl)
    merged = []
    for t, b in out:
        if merged and merged[-1][1] == b: merged[-1] = (merged[-1][0] + t, b)
        else: merged.append((t, b))
    return merged

def _omml(nodes, pt, bold):
    def rpr(b):
        return (f'<a:rPr lang="en-US" sz="{pt*100}" b="{int(b)}" i="1"><a:latin typeface="{MATH_FONT}" '
                f'panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr>')
    def run(t, upright, b):
        txt = t if upright else "".join(_ital(c, b) for c in t)
        sty = '<m:rPr><m:sty m:val="p"/></m:rPr>' if upright else ""
        return f"<m:r>{sty}{rpr(b)}<m:t>{esc(txt)}</m:t></m:r>"
    def ctrl(b): return f"<m:ctrlPr>{rpr(b)}</m:ctrlPr>"
    def g(ns, b):
        o = ""
        for n in ns:
            if n[0] == "r":
                t, up = n[1], n[2]
                if not up and re.fullmatch(r"[0-9.,()\[\]+\-=|:;/!'*<>]+", t): up = True
                o += run(t, up, b)
            elif n[0] == "sub": o += f"<m:sSub><m:sSubPr>{ctrl(b)}</m:sSubPr><m:e>{g(n[1],b)}</m:e><m:sub>{g(n[2],b)}</m:sub></m:sSub>"
            elif n[0] == "sup": o += f"<m:sSup><m:sSupPr>{ctrl(b)}</m:sSupPr><m:e>{g(n[1],b)}</m:e><m:sup>{g(n[2],b)}</m:sup></m:sSup>"
            elif n[0] == "subsup": o += (f"<m:sSubSup><m:sSubSupPr>{ctrl(b)}</m:sSubSupPr><m:e>{g(n[1],b)}</m:e>"
                                         f"<m:sub>{g(n[2],b)}</m:sub><m:sup>{g(n[3],b)}</m:sup></m:sSubSup>")
            elif n[0] == "acc":
                ch = {"hat": "̂", "tilde": "̃", "bar": "̅", "dot": "̇"}[n[1]]
                o += f'<m:acc><m:accPr><m:chr m:val="{ch}"/>{ctrl(b)}</m:accPr><m:e>{g(n[2],b)}</m:e></m:acc>'
            elif n[0] == "bold": o += g(n[1], True)
        return o
    return f"<m:oMath>{g(nodes, bold)}</m:oMath>"

# ───────────────────────── 리치 텍스트 ─────────────────────────
def segs_of(text, bold):
    """'Reward classifier $C_\\psi$' → [('t','Reward classifier ',b), ('math','C_\\psi',b)]
       **굵게** 구간 지원 (bold=False인 문단 안에서 일부만 굵게)."""
    out = []
    for i, part in enumerate(re.split(r"(\$[^$]+\$)", text)):
        if not part: continue
        if part.startswith("$"): out.append(("math", part[1:-1], bold)); continue
        for j, p in enumerate(re.split(r"\*\*", part)):
            if p: out.append(("t", p, bold or j % 2 == 1))
    return out

def para_xml(segs, pt, color=None, align="ctr", for_fallback=False):
    has_math = any(k == "math" for k, _, _ in segs)
    clr = f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>' if color else ""
    o = f'<a:p><a:pPr algn="{align}"/>'
    for k, t, b in segs:
        if k == "t":
            o += (f'<a:r><a:rPr lang="en-US" sz="{pt*100}" b="{int(b)}" dirty="0">{clr}'
                  f'<a:latin typeface="{FONT}"/></a:rPr><a:t>{esc(t)}</a:t></a:r>')
        elif for_fallback:   # 수식 미지원 뷰어용: 아래/위첨자를 baseline으로 흉내
            for txt, bl in _flat(_parse(_tokens(t))[0]):
                blx = f' baseline="{bl}"' if bl else ""
                ital = ' i="1"' if re.search(r"[A-Za-z\u03b1-\u03c9]", txt) else ""
                o += (f'<a:r><a:rPr lang="en-US" sz="{pt*100}" b="{int(b)}"{ital}{blx} dirty="0">{clr}'
                      f'<a:latin typeface="{MATH_FONT}"/></a:rPr><a:t>{esc(txt)}</a:t></a:r>')
        else:
            nodes, _ = _parse(_tokens(t))
            o += f"<a14:m>{_omml(nodes, pt, b)}</a14:m>"
    o += f'<a:endParaRPr lang="en-US" sz="{pt*100}" b="0" dirty="0"><a:latin typeface="{FONT}"/></a:endParaRPr></a:p>'
    return o, has_math

# ───────────────────────── 슬라이드 ─────────────────────────
class Slide:
    def __init__(self, deck, slide):
        self.deck, self.s = deck, slide
        self.tree = slide.shapes._spTree
        self._id = max(int(x) for x in X(self.tree, ".//p:cNvPr/@id")) + 1
        self.boxes = {}
        self.log = []          # (pt, text) — lint용
        self.boxlog = []       # (key, 부가 텍스트 줄 수)

    # ── 내부
    def nid(self):
        self._id += 1; return self._id
    def _add(self, xml, parent=None):
        els = frag(xml); (parent if parent is not None else self.tree).extend(els); return els[0]

    # ── 머리글
    def header(self, text):
        """상단 'Method' 자리 텍스트 교체 (35pt 그대로)"""
        for sp in self.s.shapes:
            if sp.has_text_frame and sp.text_frame.text.strip() == "Method":
                r = sp.text_frame.paragraphs[0].runs[0]; r.text = text
                sp.width = E(text_width([("t", text, True)], 35) + 0.25)
                sp.left = E((20 - sp.width / EMU) / 2)
        return self

    def page(self, n):
        for sp in self.s.shapes:
            if sp.has_text_frame and sp.text_frame.text.strip().isdigit():
                sp.text_frame.paragraphs[0].runs[0].text = str(n)
                sp.text_frame.word_wrap = False
                w = 0.156 + 0.2 * (len(str(n)) - 1) + 0.2   # 템플릿 쪽번호 중심 x=19.694 유지
                sp.left, sp.width = E(19.694 - w / 2), E(w)
        return self

    # ── 기본 도형
    def rect(self, x, y, w, h, fill="white", line=True, dash=None, line_color=None, prst="rect", name=None, parent=None):
        lc = f'<a:solidFill><a:srgbClr val="{line_color}"/></a:solidFill>' if line_color else ""
        dsh = f'<a:prstDash val="{dash}"/>' if dash else ""
        ln = f'<a:ln w="{E(W_SHAPE/72)}">{lc}{dsh}</a:ln>' if line else "<a:ln><a:noFill/></a:ln>"
        i = self.nid(); name = name or f"Rect {i}"
        return self._add(f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="{esc(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
                         f'<p:spPr><a:xfrm><a:off x="{E(x)}" y="{E(y)}"/><a:ext cx="{E(w)}" cy="{E(h)}"/></a:xfrm>'
                         f'<a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>{fill_xml(fill)}{ln}</p:spPr>'
                         f'{STYLE_SHAPE}{EMPTY_TX}</p:sp>', parent)

    def text(self, x, y, s, pt=PT_BOX_SUB, bold=False, color=None, align="l", anchor="tl", parent=None, name=None):
        """텍스트상자 1개. s: 문자열(\\n 줄바꿈, $수식$, **굵게**). anchor: tl|tc|tr|cl|cc|cr …
           반환: (element, w, h)"""
        lines = [segs_of(l, bold) for l in s.split("\n")]
        self.log.append((pt, s))
        w = max(text_width(l, pt) for l in lines) + 0.2
        h = len(lines) * pt * 1.2 / 72 + 0.1
        ax, ay = anchor[1], anchor[0]
        x0 = x - (w / 2 if ax == "c" else w if ax == "r" else 0)
        y0 = y - (h / 2 if ay == "c" else h if ay == "b" else 0)
        alg = {"l": "l", "c": "ctr", "r": "r"}[align] if len(align) == 1 else align
        paras = [para_xml(l, pt, color, alg) for l in lines]
        body = "".join(p for p, _ in paras); has_math = any(m for _, m in paras)
        i = self.nid(); name = name or f"TextBox {i}"
        def sp(body_xml):
            return (f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="{esc(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
                    f'<p:spPr><a:xfrm><a:off x="{E(x0)}" y="{E(y0)}"/><a:ext cx="{E(w)}" cy="{E(h)}"/></a:xfrm>'
                    f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
                    f'<p:txBody><a:bodyPr wrap="none" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>{body_xml}</p:txBody></p:sp>')
        if has_math:
            fb = "".join(para_xml(l, pt, color, alg, for_fallback=True)[0] for l in lines)
            xml = (f'<mc:AlternateContent xmlns:mc="{NS["mc"]}" xmlns:a14="{NS["a14"]}" xmlns:m="{NS["m"]}"><mc:Choice Requires="a14">{sp(body)}</mc:Choice>'
                   f'<mc:Fallback>{sp(fb)}</mc:Fallback></mc:AlternateContent>')
        else:
            xml = sp(body)
        el = self._add(xml, parent)
        return el, w, h

    def group(self, children_fn, name=None):
        """children_fn(parent) 안에서 그린 도형들을 그룹으로 묶는다."""
        i = self.nid()
        g = self._add(f'<p:grpSp><p:nvGrpSpPr><p:cNvPr id="{i}" name="{esc(name or f"Group {i}")}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
                      f'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:grpSp>')
        children_fn(g)
        xs, ys, xe, ye = [], [], [], []
        for off in X(g, ".//a:off")[1:]:
            ext = off.getnext()
            x, y = int(off.get("x")), int(off.get("y")); cx, cy = int(ext.get("cx")), int(ext.get("cy"))
            xs.append(x); ys.append(y); xe.append(x + cx); ye.append(y + cy)
        xf = g.find("p:grpSpPr/a:xfrm", NS)
        bx, by, bw, bh = min(xs), min(ys), max(xe) - min(xs), max(ye) - min(ys)
        for tag, a, b in (("a:off", "x", "y"), ("a:chOff", "x", "y")):
            el = xf.find(tag, NS); el.set(a, str(bx)); el.set(b, str(by))
        for tag in ("a:ext", "a:chExt"):
            el = xf.find(tag, NS); el.set("cx", str(bw)); el.set("cy", str(bh))
        return g

    # ── 아이콘 (SVG + PNG 대체)
    def icon(self, kind, x, y, parent=None):
        fname, size = ICON[kind]
        svg_rid = self.deck._media(self.s, os.path.join(HERE, "assets", fname + ".svg"), "image/svg+xml")
        png_rid = self.deck._media(self.s, os.path.join(HERE, "assets", fname + ".png"), "image/png")
        i = self.nid()
        descr = {"fire": "Trainable", "snow": "Frozen", "eye": "Inference only"}[kind]
        return self._add(
            f'<p:pic><p:nvPicPr><p:cNvPr id="{i}" name="Icon {kind} {i}" descr="{descr}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="{png_rid}"><a:extLst><a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}">'
            f'<asvg:svgBlip r:embed="{svg_rid}"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            f'<p:spPr><a:xfrm><a:off x="{E(x)}" y="{E(y)}"/><a:ext cx="{E(size)}" cy="{E(size)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>', parent)

    # ── 고수준 요소
    def panel(self, x, y, w, h, fill="panel", title=None, section=None, title_align="c"):
        """서브패널. title=24pt 검정, section=20pt 빨강(좌상단). section은 '빨강|검정 보조' 형식 지원"""
        self.rect(x, y, w, h, fill=fill, name="Panel")
        if title:
            if title_align == "c": self.text(x + w / 2, y + 0.08, title, PT_FIG_TITLE, True, anchor="tc", align="c")
            else: self.text(x + 0.12, y + 0.08, title, PT_FIG_TITLE, True)
        if section:
            red, _, rest = section.partition("|")
            el, tw, _ = self.text(x + 0.08, y + 0.08, red, PT_SECTION, True, RED)
            if rest: self.text(x + 0.08 + tw - 0.12, y + 0.08, rest, PT_SECTION, False)
        return (x, y, w, h)

    def fig_title(self, x, y, s):
        return self.text(x, y, s, PT_FIG_TITLE, True)

    def box(self, key, cx, cy, title, sub=None, w=None, h=None, fill="white", icon=None, title_pt=PT_BOX_TITLE,
            sub_pt=PT_BOX_SUB, sub_bold=False, dash=None, line_color=None, shape="rect", muted=False):
        """상자 = 사각형 + 제목 텍스트상자(18pt 굵게) + 부가 텍스트상자(14pt) 그룹. (cx, cy) = 중심.
           icon: 'fire'|'snow'|'eye' → 제목 오른쪽 끝, 위쪽에 붙음"""
        subs = sub.split("\n") if isinstance(sub, str) else (sub or [])
        self.boxlog.append((key, len(subs)))
        tw = text_width(segs_of(title, True), title_pt)
        sw = max([text_width(segs_of(s, sub_bold), sub_pt) for s in subs] + [0])
        ic = ICON[icon][1] if icon else 0
        w = w or max(tw + ic + 0.55, sw + 0.45, 1.0)
        th = title_pt * 1.2 / 72 + 0.1; sh = sub_pt * 1.2 / 72 + 0.1
        h = h or (0.62 if not subs else 0.12 + th + len(subs) * (sh - 0.05) + 0.12)
        x, y = cx - w / 2, cy - h / 2
        tc = MUTED if muted else None
        def draw(g):
            self.rect(x, y, w, h, fill=fill, dash=dash, line_color=line_color, prst=SHAPE.get(shape, shape), parent=g, name=f"Box {key}")
            if subs:
                ty = y + 0.12
                self.text(cx - ic / 2, ty, title, title_pt, True, color=tc, anchor="tc", align="c", parent=g)
                for k, s in enumerate(subs):
                    self.text(cx, ty + th - 0.05 + k * (sh - 0.05), s, sub_pt, sub_bold, color=tc, anchor="tc", align="c", parent=g)
            else:
                self.text(cx - ic / 2, cy, title, title_pt, True, color=tc, anchor="cc", align="c", parent=g)
                ty = cy - th / 2
            if icon:
                self.icon(icon, cx - ic / 2 + tw / 2 + 0.03, ty + 0.02, parent=g)
        self.group(draw, name=f"Node {key}")
        self.boxes[key] = (x, y, w, h)
        return self.boxes[key]

    def stack(self, key, cx, cy, title, n=3, dx=0.24, dy=0.24, count_label=None, **kw):
        """겹친 사각형(체크포인트·다중 인스턴스). 뒤 상자 2장은 오른쪽 위로 오프셋."""
        tw = text_width(segs_of(title, True), PT_BOX_TITLE)
        w = kw.pop("w", None) or max(tw + 0.55, 1.0); h = kw.pop("h", None) or 0.62
        for k in range(n - 1, 0, -1):
            self.rect(cx - w / 2 + k * dx, cy - h / 2 - k * dy, w, h, fill=kw.get("fill", "white"))
        b = self.box(key, cx, cy, title, w=w, h=h, **kw)
        if count_label:
            self.text(cx + w / 2 + (n - 1) * dx + 0.05, cy - h / 2 - (n - 1) * dy, count_label, PT_EDGE + 4, anchor="tl")
        return b

    # ── 연결선
    def _seg(self, x1, y1, x2, y2, head=False, color=None, dash=None, w=W_LINE):
        fh = ' flipH="1"' if x2 < x1 else ""; fv = ' flipV="1"' if y2 < y1 else ""
        clr = f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>' if color else ""
        dsh = f'<a:prstDash val="{dash}"/>' if dash else ""
        tail = '<a:tailEnd type="triangle"/>' if head else ""
        i = self.nid()
        return self._add(
            f'<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="{i}" name="{"Arrow" if head else "Line"} {i}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>'
            f'<p:spPr><a:xfrm{fh}{fv}><a:off x="{E(min(x1,x2))}" y="{E(min(y1,y2))}"/><a:ext cx="{E(abs(x2-x1))}" cy="{E(abs(y2-y1))}"/></a:xfrm>'
            f'<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln w="{E(w/72)}">{clr}{dsh}{tail}</a:ln></p:spPr>{STYLE_LINE}</p:cxnSp>')

    def path(self, pts, head=True, color=None, dash=None, label=None, label_at=None, label_side="above", label_bold=False):
        """직각 꺾은선 화살표. pts=[(x,y),...] — 마지막 선분에만 화살촉. (사용자 방식: 직선 조각들)"""
        for k in range(len(pts) - 1):
            self._seg(*pts[k], *pts[k + 1], head=head and k == len(pts) - 2, color=color, dash=dash)
        if label:
            (ax, ay), (bx, by) = (pts[label_at], pts[label_at + 1]) if label_at is not None else (pts[0], pts[1])
            mx, my = (ax + bx) / 2, (ay + by) / 2
            if ay == by:   # 가로 선분: 위/아래
                self.text(mx, my - 0.02 if label_side == "above" else my + 0.02, label, PT_EDGE, label_bold,
                          anchor="bc" if label_side == "above" else "tc", align="c")
            else:          # 세로 선분: 오른쪽/왼쪽
                self.text(mx + 0.05 if label_side != "left" else mx - 0.05, my, label, PT_EDGE, label_bold,
                          anchor="cl" if label_side != "left" else "cr")

    def arrow(self, a, b, side_a=None, side_b=None, **kw):
        """상자 key a → b 자동 연결. 같은 열이면 세로, 같은 행이면 가로, 아니면 ㄱ자."""
        ax, ay, aw, ah = self.boxes[a]; bx, by, bw, bh = self.boxes[b]
        acx, acy, bcx, bcy = ax + aw / 2, ay + ah / 2, bx + bw / 2, by + bh / 2
        if abs(acx - bcx) < 0.05:
            pts = [(acx, ay + ah), (acx, by)] if bcy > acy else [(acx, ay), (acx, by + bh)]
        elif abs(acy - bcy) < 0.05:
            pts = [(ax + aw, acy), (bx, acy)] if bcx > acx else [(ax, acy), (bx + bw, acy)]
        else:   # 출발: 옆면, 도착: 위/아래면
            sx = ax + aw if bcx > acx else ax
            if side_a == "bottom": pts = [(acx, ay + ah), (acx, bcy), (bx if bcx > acx else bx + bw, bcy)]
            else: pts = [(sx, acy), (bcx, acy), (bcx, by if bcy > acy else by + bh)]
        self.path(pts, **kw)

    def loss_link(self, a, b):
        """손실 연결: 얇은 파란 실선(accent1), 화살촉 없음"""
        ax, ay, aw, ah = self.boxes[a]; bx, by, bw, bh = self.boxes[b]
        x = ax + aw / 2
        self._seg(x, ay + ah, x, by, color="4F81BD")

    def block_arrow(self, x, y, w, h, direction="down"):
        """굵은 파란 블록 화살표 (큰 단계 전환)"""
        prst = {"down": "downArrow", "right": "rightArrow", "up": "upArrow", "left": "leftArrow"}[direction]
        return self.rect(x, y, w, h, fill="4F81BD", prst=prst, name="Block arrow")

    def divider(self, x1, y1, x2, y2):
        self._seg(x1, y1, x2, y2)

    def legend(self, x, y, kinds=("fire", "snow", "eye")):
        """아이콘 범례 (3종이 다 나오면 필수). (x, y) = 좌상단"""
        lab = {"fire": "Trainable", "snow": "Frozen", "eye": "Inference only"}
        rows = len(kinds); h = 0.2 + rows * 0.42; w = 2.1
        def draw(g):
            self.rect(x, y, w, h, fill="white", parent=g, name="Legend")
            for k, kind in enumerate(kinds):
                yy = y + 0.12 + k * 0.42
                self.icon(kind, x + 0.15, yy + 0.04, parent=g)
                self.text(x + 0.55, yy + 0.19, lab[kind], PT_BOX_SUB, anchor="cl", parent=g)
        self.group(draw, name="Legend")

    def verdict(self, s, y=None):
        """하단 총평 문장 — 베이스 하단 중앙, 24pt"""
        bx, by, bw, bh = BASE
        return self.text(bx + bw / 2, y if y is not None else by + bh - MARGIN / 2, s, PT_VERDICT, anchor="bc", align="c")


class Deck:
    def __init__(self, template):
        self.prs = Presentation(template)
        self._mediacache = {}
        self.slides = []
        lst = self.prs.slides._sldIdLst
        ids = list(lst)
        self._base_xml = None
        keep = ids[BASE_SLIDE_INDEX - 1] if len(ids) >= BASE_SLIDE_INDEX else ids[0]
        for sid in ids:
            if sid is not keep:
                self.prs.part.drop_rel(sid.get("{%s}id" % NS["r"])); lst.remove(sid)
        self._base = self.prs.slides[0]
        self._used_base = False

    def _media(self, slide, path, ctype):
        key = (id(slide), path)
        if key in self._mediacache: return self._mediacache[key]
        pkg = self.prs.part.package
        ext = os.path.splitext(path)[1]
        if path not in self._mediacache:
            pn = pkg.next_partname("/ppt/media/kit_image%d" + ext)
            self._mediacache[path] = Part(pn, ctype, pkg, open(path, "rb").read())
        rid = slide.part.relate_to(self._mediacache[path], RT.IMAGE)
        self._mediacache[key] = rid
        return rid

    def new_slide(self):
        """템플릿 6p를 복제한 새 슬라이드"""
        if not self._used_base:
            self._used_base = True; sl = Slide(self, self._base); self.slides.append(sl); return sl
        src = self._base
        new = self.prs.slides.add_slide(src.slide_layout)
        for sp in list(new.shapes._spTree)[2:]: new.shapes._spTree.remove(sp)
        # 원본(복제 시점의 템플릿 6p 원형)을 다시 읽기 위해 보관해 둔 XML 사용
        for el in self._pristine:
            new.shapes._spTree.append(copy.deepcopy(el))
        # 로고 등 이미지 rId 재매핑
        remap = {rid: new.part.relate_to(src.part.rels[rid]._target, RT.IMAGE) for rid in self._pristine_rids}
        for el in new.shapes._spTree.iter():
            for attr in ("{%s}embed" % NS["r"], "{%s}link" % NS["r"]):
                v = el.get(attr)
                if v in remap: el.set(attr, remap[v])
        sl = Slide(self, new); self.slides.append(sl); return sl

    def snapshot_base(self):
        tree = self._base.shapes._spTree
        self._pristine = [copy.deepcopy(e) for e in list(tree)[2:]]
        self._pristine_rids = {r for e in self._pristine for r in X(e, ".//@r:embed|.//@r:link")}

    def lint(self):
        """줄글·과밀 검사. 경고가 있으면 내용을 줄이거나 슬라이드를 나눈다 (글자 크기를 줄이지 않는다)."""
        warns = []
        for n, sl in enumerate(self.slides, 1):
            for key, k in sl.boxlog:
                if k > MAX_LINES: warns.append(f"[p{n}] 상자 {key!r} 부가 텍스트 {k}줄 > {MAX_LINES} → 핵심만 남길 것")
            if len(sl.log) > MAX_TEXT_PIECES:
                warns.append(f"[p{n}] 텍스트 {len(sl.log)}개 > {MAX_TEXT_PIECES} → 슬라이드를 나눌 것")
            for pt, t in sl.log:
                lines = t.split("\n")
                lim = MAX_VERDICT_WORDS if pt == PT_VERDICT else MAX_WORDS_PER_LINE
                if pt >= PT_FIG_TITLE and pt != PT_VERDICT: continue
                if len(lines) > MAX_LINES:
                    warns.append(f"[p{n}] {len(lines)}줄 > {MAX_LINES}: {t[:40]!r} → 핵심만 남길 것")
                for l in lines:
                    words = len([w for w in re.sub(r"\$[^$]*\$", "X", l).replace("**", "").split() if re.search(r"\w", w)])
                    if words > lim:
                        warns.append(f"[p{n}] 줄글 의심 ({words}단어 > {lim}): {l[:50]!r}")
        return warns

    def save(self, path, strict=True):
        warns = self.lint()
        for w in warns: print("LINT", w)
        if warns and strict:
            raise SystemExit("LINT 경고를 해결한 뒤 저장하세요 (strict=False로 강제 저장 가능)")
        self.prs.save(path)


def open_deck(template=None):
    """template 생략 시 kit 폴더의 base.pptx (템플릿 6p만 남긴 사본) 사용"""
    if template is None:
        template = os.path.join(HERE, "base.pptx")
    d = Deck(template); d.snapshot_base(); return d


def list_refs(query=None, folder=None):
    """refs/ 폴더의 레퍼런스 목록을 슬라이드 노트로 출력한다. 반환: [(파일 경로, 슬라이드 번호, 노트)]
       query: 노트 검색어 (문자열 또는 목록, 하나라도 들어 있으면 표시) — 예: list_refs("단계적 빌드"), list_refs(["좌우", "분기"])
       - kit_*.pptx  : kit 규칙으로 만든 발표 → 구성·장치·수치까지 참고 (먼저, 최신 날짜부터)
       - legacy_*.pptx : 옛 발표 → 구성·장치만 참고 (베이스·선 두께·글자 크기가 kit와 다를 수 있음)
       구조가 가장 비슷한 장을 골라 render()로 열어 본다."""
    import glob
    folder = folder or os.path.join(HERE, "refs")
    files = sorted(glob.glob(os.path.join(folder, "kit_*.pptx")), reverse=True) + \
            sorted(glob.glob(os.path.join(folder, "legacy_*.pptx")))
    if not files:
        print("refs/ 폴더에 레퍼런스가 없음"); return []
    out = []
    for f in files:
        tag = "kit (수치까지 참고)" if os.path.basename(f).startswith("kit_") else "legacy (구성만 참고)"
        print(f"\n■ {os.path.basename(f)} — {tag}")
        qs = [query] if isinstance(query, str) else (query or [])
        shown = 0
        for i, sl in enumerate(Presentation(f).slides, 1):
            note = sl.notes_slide.notes_text_frame.text if sl.has_notes_slide else ""
            if qs and not any(q.lower() in note.lower() for q in qs): continue
            out.append((f, i, note)); print(f"  {i:>2}. {note}"); shown += 1
        if qs and not shown: print("  (해당 없음)")
    return out


def add_refs(src_pptx, slides, notes, name, folder=None):
    """발표 파일에서 고른 슬라이드를 refs/<name>.pptx로 저장한다 (발표 1개 = 파일 1개).
       slides: 1부터 세는 슬라이드 번호 목록 / notes: {번호: 노트} — 노트 형식은 SKILL.md '레퍼런스 추가' 참조
       name: 'kit_YYYY-MM-DD_발표' (kit 규칙으로 만든 발표) 또는 'legacy_…' (옛 발표)
       원본 발표 파일은 건드리지 않는다. 같은 이름이 있으면 덮어쓰지 않고 예외."""
    if not re.match(r"^(kit|legacy)_", name):
        raise ValueError("이름은 kit_ 또는 legacy_ 로 시작해야 합니다")
    folder = folder or os.path.join(HERE, "refs"); os.makedirs(folder, exist_ok=True)
    out = os.path.join(folder, name + ".pptx")
    if os.path.exists(out):
        raise FileExistsError(f"{out} 이미 있음 → 같은 발표에서 더 고르는 경우 그 파일에 이어 붙이도록 슬라이드 목록을 합쳐 다시 만드세요")
    missing = [k for k in slides if not notes.get(k)]
    if missing: raise ValueError(f"노트가 없는 슬라이드: {missing} — 노트가 있어야 list_refs()로 찾을 수 있습니다")
    prs = Presentation(src_pptx); lst = prs.slides._sldIdLst; keep = set(slides)
    R = "{%s}id" % NS["r"]
    for k, sid in enumerate(list(lst), 1):
        if k not in keep: prs.part.drop_rel(sid.get(R)); lst.remove(sid)
    for k, sl in zip(sorted(keep), prs.slides):
        sl.notes_slide.notes_text_frame.text = notes[k]
    prs.save(out); print("saved", out, f"({len(prs.slides)}장)")
    return out


# ───────────────────────── 렌더 검수 (OS 무관) ─────────────────────────
def _find_soffice():
    import shutil
    for name in ("soffice", "libreoffice"):
        if shutil.which(name): return shutil.which(name)
    for c in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(c): return c
    return None

def render(pptx_path, out_dir=None, dpi=110):
    """pptx → PDF(LibreOffice) → 슬라이드별 PNG. 반환: PNG 경로 목록.
       PDF→PNG는 PyMuPDF(pip install pymupdf)를 쓰고, 없으면 pdftoppm을 쓴다."""
    import subprocess, shutil, glob, tempfile
    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice(soffice)를 찾을 수 없습니다 → 설치 후 PATH에 추가")
    out_dir = out_dir or os.path.dirname(os.path.abspath(pptx_path))
    os.makedirs(out_dir, exist_ok=True)
    profile = tempfile.mkdtemp(prefix="lo_")   # 다른 LibreOffice 창과 충돌 방지
    subprocess.run([soffice, f"-env:UserInstallation=file:///{profile.replace(os.sep, '/')}", "--headless",
                    "--convert-to", "pdf", "--outdir", out_dir, pptx_path],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
    pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf")
    stem = os.path.splitext(pdf)[0]
    try:
        try: import pymupdf as fitz
        except ImportError: import fitz  # 구버전 PyMuPDF
        pngs = []
        for i, page in enumerate(fitz.open(pdf), 1):
            f = f"{stem}-{i}.png"; page.get_pixmap(dpi=dpi).save(f); pngs.append(f)
        return pngs
    except ImportError:
        if not shutil.which("pdftoppm"):
            raise RuntimeError("PDF→PNG 도구 없음 → pip install pymupdf")
        subprocess.run(["pdftoppm", "-r", str(dpi), "-png", pdf, stem], check=True)
        return sorted(glob.glob(stem + "-*.png"))



# ───────────────────────── 사용자의 기존 발표 파일에 넣기·고치기 ─────────────────────────
def fingerprint(path):
    """파일 지문 (수정 시각 ns, 크기, sha256). 수정 직전·저장 직전 비교용."""
    import hashlib
    st = os.stat(path)
    h = hashlib.sha256(open(path, "rb").read()).hexdigest()
    return (st.st_mtime_ns, st.st_size, h)

def is_open_in_powerpoint(path):
    """PowerPoint가 파일을 열면 같은 폴더에 '~$파일명' 잠금 파일이 생긴다. Windows에서는 쓰기 잠금도 확인."""
    d, b = os.path.split(os.path.abspath(path))
    for lock in ("~$" + b, "~$" + b[2:]):
        if os.path.exists(os.path.join(d, lock)): return True
    if os.name == "nt":
        try:
            with open(path, "r+b"): pass
        except PermissionError:
            return True
    return False

def next_version_path(path):
    """원본 보존용 새 이름: deck.pptx → deck_v2.pptx → deck_v3.pptx …"""
    stem, ext = os.path.splitext(path)
    m = re.match(r"(.*)_v(\d+)$", stem)
    base, n = (m.group(1), int(m.group(2)) + 1) if m else (stem, 2)
    while os.path.exists(f"{base}_v{n}{ext}"): n += 1
    return f"{base}_v{n}{ext}"

def safe_save(prs, src_path, before, out_path=None, overwrite=False):
    """저장 직전 재확인 후 저장. before = 수정 직전에 잰 fingerprint(src_path).
       - 그 사이 원본이 바뀌었으면 저장하지 않고 예외 → 처음부터 다시 읽을 것
       - PowerPoint에서 열려 있으면 저장하지 않고 예외
       - 기본은 새 버전(_v2 …)으로 저장, 덮어쓰기는 overwrite=True일 때만"""
    if fingerprint(src_path) != before:
        raise RuntimeError("원본이 수정 중에 바뀌었습니다 → 저장하지 않음. 최신 파일을 다시 읽고 처음부터 다시 하세요.")
    if is_open_in_powerpoint(src_path):
        raise RuntimeError("PowerPoint에서 파일이 열려 있습니다 → 닫아 달라고 요청한 뒤 다시 저장하세요.")
    out = src_path if overwrite else (out_path or next_version_path(src_path))
    prs.save(out)
    return out

def _base_ids():
    """base.pptx 슬라이드의 템플릿 원형 요소 id. kit 요소는 항상 이보다 큰 id를 받으므로 id로 구분한다
       (header()가 머리글 위치를 바꿔도 구분이 흔들리지 않음)"""
    p = Presentation(os.path.join(HERE, "base.pptx"))
    ids = set()
    for el in list(p.slides[0].shapes._spTree)[2:]:
        v = X(el, ".//p:cNvPr/@id")
        if v: ids.add(v[0])
    return ids

def _find_user_base(slide):
    """대상 슬라이드에서 사용자가 깐 회색 베이스(bg1 85%, 폭 10in 이상)를 찾는다 → (x, y, w, h) in"""
    for sp in slide.shapes:
        xml = etree.tostring(sp._element).decode()
        if 'val="bg1"><a:lumMod val="85000"' in xml and sp.width and sp.width > E(10):
            return (sp.left / EMU, sp.top / EMU, sp.width / EMU, sp.height / EMU)
    return None

def _has_frame(slide):
    """템플릿 틀(바닥글 'Robot Intelligence Learning & System')이 있는가"""
    return any(sp.has_text_frame and "Robot Intelligence Learning" in sp.text_frame.text for sp in slide.shapes)

def _is_base_rect(el):
    xml = etree.tostring(el).decode()
    ext = X(el, "./p:spPr/a:xfrm/a:ext")
    return 'val="bg1"><a:lumMod val="85000"' in xml and ext and int(ext[0].get("cx")) > E(10)

def copy_diagram(src_slide, dst_slide, align_to_base=True, dx=0.0, dy=0.0, ensure_base=True, ensure_frame=True):
    """kit로 만든 슬라이드(src)에서 다이어그램 요소만 대상 슬라이드(dst)로 복사한다.
       - 템플릿 원형 요소(base 6p 복제분: 베이스·머리글·로고·쪽번호)는 건너뛴다
       - 대상에 베이스가 없으면 (ensure_base) kit 베이스를 맨 아래에 깐다
         대상에 템플릿 틀(머리글·로고·바닥글)까지 없으면 (ensure_frame) 틀 전체를 맨 아래에 깐다
       - cNvPr id는 대상 슬라이드 최대값 다음부터 새로 매긴다 (수식 Choice/Fallback 쌍은 같은 id 유지)
       - 아이콘·로고 이미지 관계를 대상 슬라이드에 다시 연결한다 (대상 파일 안에 새 이름으로 복제)
       - align_to_base: 대상 베이스 위치가 kit BASE와 다르면 그만큼 평행 이동
       대상 슬라이드의 기존 요소는 건드리지 않는다.
       반환: {"copied": 다이어그램 요소 수, "added_base": bool, "added_frame": bool} → 깐 것이 있으면 사용자에게 알린다"""
    base_ids = _base_ids()
    items, frame = [], []
    for el in list(src_slide.shapes._spTree)[2:]:
        v = X(el, ".//p:cNvPr/@id")
        (frame if v and v[0] in base_ids else items).append(copy.deepcopy(el))
    ub = _find_user_base(dst_slide)
    added_base = added_frame = False
    under = []                                  # 맨 아래(기존 요소보다 뒤)에 깔 요소
    if ub is None and ensure_frame and not _has_frame(dst_slide):
        under = frame; added_base = added_frame = True
    elif ub is None and ensure_base:
        under = [el for el in frame if _is_base_rect(el)]; added_base = bool(under)
    # 위치 보정 (사용자 베이스가 있을 때만)
    if align_to_base and ub:
        dx += ub[0] - BASE[0]; dy += ub[1] - BASE[1]
        if abs(ub[2] - BASE[2]) > 0.05 or abs(ub[3] - BASE[3]) > 0.05:
            print(f"주의: 대상 베이스 크기 {ub[2]:.3f}×{ub[3]:.3f} in ≠ kit {BASE[2]:.3f}×{BASE[3]:.3f} in → 넘치는지 렌더링으로 확인")
    if dx or dy:
        for el in items:
            for xf in X(el, "./p:spPr/a:xfrm | ./p:grpSpPr/a:xfrm | ./mc:Choice/p:sp/p:spPr/a:xfrm | ./mc:Fallback/p:sp/p:spPr/a:xfrm"):
                o = xf.find("a:off", NS); o.set("x", str(int(o.get("x")) + E(dx))); o.set("y", str(int(o.get("y")) + E(dy)))
                co = xf.find("a:chOff", NS)
                if co is not None: co.set("x", str(int(co.get("x")) + E(dx))); co.set("y", str(int(co.get("y")) + E(dy)))
    # id 재지정
    used = [int(v) for v in X(dst_slide.shapes._spTree, ".//p:cNvPr/@id")]
    nxt = max(used + [1]) + 1; idmap = {}
    for el in under + items:
        for c in X(el, ".//p:cNvPr"):
            old = c.get("id")
            if old not in idmap: idmap[old] = str(nxt); nxt += 1
            c.set("id", idmap[old])
        for c in X(el, ".//a:stCxn | .//a:endCxn"):
            if c.get("id") in idmap: c.set("id", idmap[c.get("id")])
    # 이미지 관계 재연결
    R_E = "{%s}embed" % NS["r"]
    remap = {}
    for el in under + items:
        for node in el.iter():
            rid = node.get(R_E)
            if not rid: continue
            if rid not in remap:   # 이미지 파트를 대상 파일 안에 새 이름으로 복제해 연결 (이름 충돌 방지)
                t = src_slide.part.rels[rid]._target; pkg = dst_slide.part.package
                ext = str(t.partname).rsplit(".", 1)[-1]
                newp = Part(pkg.next_partname("/ppt/media/image%d." + ext), t.content_type, pkg, t.blob)
                remap[rid] = dst_slide.part.relate_to(newp, RT.IMAGE)
            node.set(R_E, remap[rid])
    tree = dst_slide.shapes._spTree
    for k, el in enumerate(under): tree.insert(2 + k, el)   # nvGrpSpPr, grpSpPr 바로 다음 = 맨 아래
    tree.extend(items)
    if added_frame:
        # 쪽 번호를 이 슬라이드의 실제 위치로 맞춘다 (kit 슬라이드의 번호가 따라오므로)
        try:
            prs = dst_slide.part.package.presentation_part.presentation
            Slide(None, dst_slide).page(prs.slides.index(dst_slide) + 1)
        except Exception as e:
            print("쪽 번호 자동 설정 실패 → 직접 확인:", e)
        print("대상 슬라이드에 템플릿 틀이 없어 틀 전체(베이스·머리글·로고·바닥글·쪽번호)를 깔았습니다 → 머리글 텍스트 확인")
    elif added_base: print("대상 슬라이드에 베이스가 없어 kit 베이스를 깔았습니다")
    return {"copied": len(items), "added_base": added_base, "added_frame": added_frame}
