# -*- coding: utf-8 -*-
"""LaTeX 수식 -> 한글(HWP) 수식 스크립트 변환기 (자립형, 자체 구현).

한글 수식편집기 스크립트는 LaTeX가 아니다:
  \frac{a}{b} -> {a} over {b}     \sqrt{x} -> sqrt {x}
  x^{2} -> x ^{2}                 \alpha -> alpha (백슬래시 없음)
  \sin -> sin (함수명은 한글이 자동 정자체, rm 불필요)
LaTeX↔HWP 매핑 개념은 공개 프로젝트 soseon203/exam-paper-korean 을 참고해
자체적으로 재작성하고 보강(축약형 분수/루트·도(°)·정자체 규칙 등)했다.
"""
import re

GREEK = ("alpha beta gamma delta epsilon varepsilon zeta eta theta vartheta "
         "iota kappa lambda mu nu xi pi rho sigma tau upsilon phi varphi chi psi omega "
         "Gamma Delta Theta Lambda Xi Pi Sigma Upsilon Phi Chi Psi Omega").split()
GREEK_MAP = {"\\" + g: g for g in GREEK}

SYMBOL_MAP = {
    r"\times": "TIMES", r"\cdot": "CDOT", r"\div": "DIV", r"\pm": "PLUSMINUS", r"\mp": "MINUSPLUS",
    r"\leq": "LEQ", r"\le": "LEQ", r"\geq": "GEQ", r"\ge": "GEQ", r"\neq": "neq", r"\ne": "neq",
    r"\approx": "APPROX", r"\equiv": "EQUIV", r"\sim": "SIM", r"\simeq": "SIMEQ", r"\cong": "CONG",
    r"\propto": "PROPTO", r"\asymp": "ASYMP", r"\doteq": "DOTEQ", r"\prec": "PREC", r"\succ": "SUCC",
    r"\ll": "<<", r"\gg": ">>", r"\gt": ">", r"\lt": "<",
    r"\infty": "inf", r"\partial": "partial", r"\nabla": "LAPLACE", r"\forall": "forall",
    r"\exists": "EXIST", r"\in": "in", r"\notin": "notin", r"\ni": "OWNS",
    r"\subset": "subset", r"\supset": "supset", r"\subseteq": "subseteq", r"\supseteq": "supseteq",
    r"\cup": "SMALLUNION", r"\cap": "SMALLINTER", r"\emptyset": "emptyset",
    r"\vee": "VEE", r"\lor": "VEE", r"\wedge": "WEDGE", r"\land": "WEDGE",
    r"\neg": "LNOT", r"\lnot": "LNOT", r"\oplus": "OPLUS", r"\otimes": "OTIMES",
    r"\therefore": "therefore", r"\because": "because", r"\angle": "angle",
    r"\perp": "BOT", r"\parallel": "parallel", r"\triangle": "TRIANGLE",
    r"\square": '"□"', r"\circ": "CIRC", r"\bullet": "BULLET", r"\star": "STAR",
    r"\diamond": "DIAMOND", r"\top": "TOP", r"\vdash": "VDASH", r"\models": "MODELS",
    r"\rightarrow": "->", r"\leftarrow": "<-", r"\leftrightarrow": "<->", r"\to": "->", r"\gets": "<-",
    r"\uparrow": "uparrow", r"\downarrow": "downarrow", r"\mapsto": "mapsto",
    r"\Rightarrow": "RARROW", r"\Leftarrow": "LARROW", r"\Leftrightarrow": "LRARROW",
    r"\Longleftrightarrow": "LRARROW", r"\Longrightarrow": "RARROW", r"\Longleftarrow": "LARROW",
    r"\iff": "LRARROW", r"\implies": "RARROW",
    r"\ldots": "LDOTS", r"\cdots": "CDOTS", r"\vdots": "VDOTS", r"\ddots": "DDOTS",
    r"\prime": "prime", r"\ell": "ELL", r"\hbar": "HBAR", r"\Im": "IMAG", r"\Re": "REIMAGE",
}

FUNC = ("sin cos tan sec csc cot arcsin arccos arctan sinh cosh tanh coth "
        "log ln lg exp det max min sup inf lim gcd arg dim ker hom mod lcm").split()
FUNC_MAP = {"\\" + f: f for f in FUNC}

ACCENT_MAP = {
    r"\vec": "VEC", r"\bar": "BAR", r"\hat": "HAT", r"\tilde": "TILDE", r"\dot": "DOT",
    r"\ddot": "DDOT", r"\acute": "acute", r"\grave": "grave", r"\check": "check", r"\breve": "arch",
    r"\overline": "overline", r"\underline": "underline",
    r"\overrightarrow": "VEC", r"\widehat": "HAT", r"\widetilde": "TILDE",
}


def _bg(name):
    L0 = r"[^{}]*"; L1 = r"(?:[^{}]|\{" + L0 + r"\})*"; L2 = r"(?:[^{}]|\{" + L1 + r"\})*"; L3 = r"(?:[^{}]|\{" + L2 + r"\})*"
    return r"\{(?P<" + name + r">" + L3 + r")\}"

def _bgc(name):
    L0 = r"[^{}]*"; L1 = r"(?:[^{}]|\{" + L0 + r"\})*"; L2 = r"(?:[^{}]|\{" + L1 + r"\})*"; L3 = r"(?:[^{}]|\{" + L2 + r"\})*"
    return r"(?:\{(?P<" + name + r">" + L3 + r")\}|(?P<" + name + r"_c>[^\s{}\\]))"


P_FRAC = re.compile(r"\\[dt]?frac\s*" + _bg("num") + r"\s*" + _bg("den"))
P_SQRTN = re.compile(r"\\sqrt\s*\[([^\]]+)\]\s*" + _bg("body"))
P_SQRT = re.compile(r"\\sqrt\s*" + _bg("body"))
P_BIGOP = re.compile(r"\\(sum|prod|coprod|int|iint|iiint|oint|bigcup|bigcap)"
                     r"(?:\s*_\s*" + _bgc("lo") + r")?(?:\s*\^\s*" + _bgc("hi") + r")?")
P_ACCENT = re.compile(r"\\(" + "|".join(re.escape(k[1:]) for k in ACCENT_MAP) + r")\s*" + _bg("body"))
P_LEFTRIGHT = re.compile(r"\\left\s*([(\[{|.])\s*(.*?)\s*\\right\s*([)\]}|.])", re.DOTALL)
P_SUP = re.compile(r"\^\s*" + _bgc("sup"))
P_SUB = re.compile(r"_\s*" + _bgc("sub"))
P_TEXT = re.compile(r"\\text\s*" + _bg("txt"))
P_MATHRM = re.compile(r"\\mathrm\s*" + _bg("txt"))
P_MATHBF = re.compile(r"\\mathbf\s*" + _bg("txt"))
P_BINOM = re.compile(r"\\binom\s*" + _bg("top") + r"\s*" + _bg("bot"))
P_ENV = re.compile(r"\\begin\{(cases|pmatrix|bmatrix|vmatrix|matrix)\}\s*(.*?)\s*\\end\{\1\}", re.DOTALL)
ENV_MAP = {"cases": "CASES", "pmatrix": "PMATRIX", "bmatrix": "BMATRIX", "vmatrix": "DMATRIX", "matrix": "MATRIX"}
BIGOP_MAP = {"SUM": "SUM", "PROD": "PROD", "COPROD": "COPROD", "INT": "INT",
             "IINT": "DINT", "IIINT": "TINT", "OINT": "OINT", "BIGCUP": "UNION", "BIGCAP": "INTER"}


def _gm(m, name):
    v = m.group(name)
    if v is None:
        v = m.group(name + "_c")
    return v or ""


def _convert(s):
    if not s:
        return ""
    # 0. 행렬/조건식
    def env_repl(m):
        content = re.sub(r"\\\\", " # ", m.group(2))
        return ENV_MAP[m.group(1)] + " {" + _convert(content) + "}"
    s = P_ENV.sub(env_repl, s)
    # 1. text/mathrm/mathbf
    s = P_TEXT.sub(lambda m: '"' + m.group("txt") + '"', s)
    s = P_MATHRM.sub(lambda m: "rm {" + m.group("txt") + "}", s)
    s = P_MATHBF.sub(lambda m: "bold " + m.group("txt"), s)
    # 2. binom
    s = P_BINOM.sub(lambda m: "LEFT ( {" + _convert(m.group("top")) + "} atop {" + _convert(m.group("bot")) + "} RIGHT )", s)
    # 3. frac
    s = P_FRAC.sub(lambda m: "{" + _convert(m.group("num")) + "} over {" + _convert(m.group("den")) + "}", s)
    # 4. sqrt[n] / sqrt
    s = P_SQRTN.sub(lambda m: "root {" + _convert(m.group(1)) + "} of {" + _convert(m.group("body")) + "}", s)
    s = P_SQRT.sub(lambda m: "sqrt {" + _convert(m.group("body")) + "}", s)
    # 5. 큰 연산자
    def bigop_repl(m):
        op = BIGOP_MAP.get(m.group(1).upper(), m.group(1).upper())
        lo, hi = _gm(m, "lo"), _gm(m, "hi")
        r = op
        if lo: r += " _{" + _convert(lo) + "}"
        if hi: r += " ^{" + _convert(hi) + "}"
        return r
    s = P_BIGOP.sub(bigop_repl, s)
    # 6. left ... right
    def lr_repl(m):
        dm = {"(": "(", ")": ")", "[": "[", "]": "]", r"\{": "lbrace", r"\}": "rbrace",
              "{": "lbrace", "}": "rbrace", "|": "|", ".": ""}
        l = dm.get(m.group(1), m.group(1)); r = dm.get(m.group(3), m.group(3))
        inner = _convert(m.group(2))
        if l and r: return f"LEFT {l} {inner} RIGHT {r}"
        if l: return f"LEFT {l} {inner}"
        if r: return f"{inner} RIGHT {r}"
        return inner
    s = P_LEFTRIGHT.sub(lr_repl, s)
    # 7. accent
    s = P_ACCENT.sub(lambda m: ACCENT_MAP.get("\\" + m.group(1), m.group(1).upper()) + " {" + _convert(m.group("body")) + "}", s)
    # 8. 그리스 (앞뒤 공백)
    for k, v in sorted(GREEK_MAP.items(), key=lambda x: -len(x[0])):
        s = s.replace(k, " " + v + " ")
    # 9. 기호
    for k, v in sorted(SYMBOL_MAP.items(), key=lambda x: -len(x[0])):
        s = s.replace(k, " " + v + " ")
    # 10. 함수명 (rm 안 붙임: 한글이 자동 정자체, rm은 뒤 변수까지 정자체로 번짐)
    for k, v in sorted(FUNC_MAP.items(), key=lambda x: -len(x[0])):
        s = s.replace(k, " " + v + " ")
    # 11. 위/아래 첨자
    s = P_SUP.sub(lambda m: " ^{" + _convert(_gm(m, "sup")) + "}", s)
    s = P_SUB.sub(lambda m: " _{" + _convert(_gm(m, "sub")) + "}", s)
    # 12. 단순 그룹 재귀
    s = re.sub(r"\{([^{}]+)\}", lambda m: "{" + _convert(m.group(1)) + "}", s)
    # 13. 남은 LaTeX 공백/명령 정리
    s = s.replace("\\,", "`").replace("\\;", "~").replace("\\!", "")
    s = s.replace("\\qquad", "~~~~").replace("\\quad", "~~").replace("\\\\", "")
    s = re.sub(r"\\[a-zA-Z]+", "", s)
    return s


# ---------- 전처리 (축약형 보정) ----------
def _read_group(s, i):
    n = len(s)
    while i < n and s[i] == " ":
        i += 1
    if i >= n:
        return "", i
    c = s[i]
    if c == "{":
        depth = 0
        for j in range(i, n):
            if s[j] == "{": depth += 1
            elif s[j] == "}":
                depth -= 1
                if depth == 0:
                    return _normalize(s[i + 1:j]), j + 1
        return _normalize(s[i + 1:]), n
    if c == "\\":
        m = re.match(r"\\[a-zA-Z]+", s[i:])
        if m: return m.group(0), i + m.end()
        return s[i:i + 2], i + 2
    return c, i + 1


def _normalize(s):
    """\\frac/\\sqrt 의 중괄호 없는 축약형 보정 (\\dfrac\\pi6, \\sqrt3 등)."""
    out = []; i = 0; n = len(s)
    while i < n:
        m = re.match(r"\\[dt]?frac", s[i:])
        if m:
            i += m.end()
            a, i = _read_group(s, i); b, i = _read_group(s, i)
            out.append("\\frac{" + a + "}{" + b + "}"); continue
        if re.match(r"\\sqrt", s[i:]):
            after = i + 5; j = after
            while j < n and s[j] == " ": j += 1
            if j < n and s[j] not in "[{":
                a, k = _read_group(s, after)
                out.append("\\sqrt{" + a + "}"); i = k; continue
        out.append(s[i]); i += 1
    return "".join(out)


def preprocess(latex):
    s = latex
    s = re.sub(r"\\[ :>]", " ", s)                       # \ \: \>  얇은 공백
    s = re.sub(r"\\tag\s*\{[^{}]*\}", "", s)             # \tag{} 제거
    s = re.sub(r"\\boxed\s*", "", s)                     # \boxed 테두리 제거(내용만)
    s = re.sub(r"([\^_])\s*(\\[a-zA-Z]+)", r"\1{\2}", s)  # ^\circ -> ^{\circ}
    s = _normalize(s)
    return s


def to_hwp(latex):
    s = latex.strip().strip("$").strip()
    for e in (r"\[", r"\]", r"\(", r"\)"):
        s = s.replace(e, "")
    for env in ("equation", "align", "gather", "displaymath"):
        s = re.sub(r"\\begin\{" + env + r"\*?\}", "", s)
        s = re.sub(r"\\end\{" + env + r"\*?\}", "", s)
    s = _convert(preprocess(s.strip()))
    return re.sub(r"  +", " ", s).strip()


if __name__ == "__main__":
    import io, sys
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    for t in [r"\dfrac\pi6", r"\tan t", r"\sin\alpha+\cos\beta", r"f(g(x))",
              r"\sqrt{a^2+b^2}", r"\lim_{x\to0}\dfrac{\sin x}{x}=1",
              r"\int_0^1 x^2\,dx", r"\overline{\mathrm{AB}}", r"90^\circ",
              r"\sum_{k=1}^{n}k=\dfrac{n(n+1)}{2}"]:
        print(f"{t:42} -> {to_hwp(t)}")
