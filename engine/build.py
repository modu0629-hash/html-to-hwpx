# -*- coding: utf-8 -*-
"""개념.html -> HWPX 변환 빌더 v3 (한글 COM). 박스=한글 표(셀) → 한글에서 직접 편집 가능."""
import sys, io, os, re, base64, tempfile, argparse, struct, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from conv import to_hwp
from bs4 import BeautifulSoup, NavigableString, Tag
import win32com.client as win32

SKIP_CLASSES = {"corner", "sheet-header", "sheet-footer"}

F_TITLE_B = "경기천년제목 Bold"
F_TITLE_M = "경기천년제목 Medium"
F_BAT_R = "경기천년바탕 Regular"
F_BAT_B = "경기천년바탕 Bold"

TEXT_W_MM = 166.0   # A4 - 좌우여백(22+22)

def bold_face(face):
    return F_TITLE_B if "제목" in face else F_BAT_B

# 표 박스 종류: 테두리 색/두께
BOX = {
    "formula": dict(color=(0, 0, 0),        width="0.4mm"),
    "stmt":    dict(color=(0, 0, 0),        width="0.4mm"),
    "cover":   dict(color=(0, 0, 0),        width="0.4mm"),
    "grading": dict(color=(0x55, 0x55, 0x55), width="0.3mm"),
}

STYLE = {
    "covertitle": dict(face=F_TITLE_B, size=3200, lh=130, before=0,   after=0,   align="center"),
    "small":      dict(face=F_TITLE_M, size=1200, lh=140, before=0,   after=4,   align="center"),
    "h3c":        dict(face=F_BAT_R,   size=1250, lh=150, before=5,   after=0,   align="center"),
    "h1":         dict(face=F_TITLE_B, size=1800, lh=135, before=0,   after=3.5, align="left"),
    "h2":         dict(face=F_BAT_B,   size=1300, lh=140, before=3.5, after=1.5, align="left"),
    "h3":         dict(face=F_BAT_B,   size=1250, lh=140, before=2.5, after=1,   align="left"),
    "subtitle":   dict(face=F_BAT_R,   size=1100, lh=165, before=2.5, after=0.6, align="left"),
    "body":       dict(face=F_BAT_R,   size=1050, lh=165, before=0,   after=0.8, align="justify"),
    "li":         dict(face=F_BAT_R,   size=1000, lh=155, before=0,   after=0.3, align="justify"),
    "note":       dict(face=F_BAT_R,   size=850,  lh=148, before=1.5, after=1,   align="justify"),
    "formula":    dict(face=F_BAT_R,   size=1100, lh=140, before=0.5, after=0.5, align="center"),
    "src":        dict(face=F_BAT_R,   size=850,  lh=135, before=0,   after=0.6, align="left", color=(0x66, 0x66, 0x66)),
    # 표 셀 안 본문
    "cell":       dict(face=F_BAT_R,   size=1050, lh=150, before=0,   after=0,   align="justify"),
}

# ---------------- 파싱 ----------------
def split_math(text, attrs):
    text = text.replace("\\(", "$").replace("\\)", "$")
    runs = []
    for p in re.split(r"(\$[^$]*\$)", text):
        if not p:
            continue
        if len(p) >= 2 and p[0] == "$" and p[-1] == "$":
            runs.append(("math", p[1:-1], set(attrs)))
        else:
            runs.append(("text", p, set(attrs)))
    return runs

def inline_runs(el, attrs=frozenset()):
    runs = []
    for ch in el.children:
        if isinstance(ch, NavigableString):
            runs += split_math(str(ch), attrs)
        elif isinstance(ch, Tag):
            cls = ch.get("class") or []
            a = set(attrs)
            if ch.name in ("b", "strong"):
                a.add("bold")
            if "cite" in cls or "score-badge" in cls or "src" in cls:
                a.add("cite")
            if ch.name == "br":
                runs.append(("br", "", set(attrs)))
            else:
                runs += inline_runs(ch, a)
    return runs

def extract_display(el):
    txt = el.get_text()
    found = re.findall(r"\$\$(.+?)\$\$", txt, re.S)
    if not found:
        found = re.findall(r"\$(.+?)\$", txt, re.S)
    return [f.strip() for f in found if f.strip()]

def has_cls(ch, name):
    return name in (ch.get("class") or [])

class Ctr:
    def __init__(self): self.v = 0
    def nxt(self): self.v += 1; return self.v

def add(blocks, blk, box, grp):
    if box:
        blk["box"] = box
        blk["grp"] = grp
    blocks.append(blk)

def walk_blocks(el, blocks, gctr, box=None, grp=None):
    for ch in el.children:
        if isinstance(ch, NavigableString):
            t = str(ch).strip()
            if t:
                add(blocks, {"style": "body", "runs": split_math(t, frozenset())}, box, grp)
            continue
        if not isinstance(ch, Tag):
            continue
        cls = ch.get("class") or []
        if SKIP_CLASSES.intersection(cls):
            continue
        name = ch.name
        # --- 채점기준: 표(내용|점수) 블록으로 ---
        if has_cls(ch, "grading-box"):
            te = ch.select_one(".grading-title")
            title = te.get_text(" ", strip=True) if te else ""
            rows = []
            for li in ch.select("li"):
                step = li.select_one(".step")
                pts = li.select_one(".pts")
                srun = inline_runs(step) if step else inline_runs(li)
                ptxt = pts.get_text(" ", strip=True) if pts else ""
                rows.append((srun, ptxt))
            blocks.append({"style": "gradingtable", "title": title, "rows": rows})
            continue
        # --- 박스 컨테이너: 1칸 표로 묶음 ---
        if has_cls(ch, "stmt-box"):
            walk_blocks(ch, blocks, gctr, "stmt", gctr.nxt()); continue
        if has_cls(ch, "cover-toc"):
            walk_blocks(ch, blocks, gctr, "cover", gctr.nxt()); continue
        if has_cls(ch, "formula"):
            g = gctr.nxt()
            for lx in extract_display(ch):
                add(blocks, {"style": "formula", "latex": lx}, "formula", g)
            continue
        if name == "img":
            add(blocks, {"style": "image", "src": ch.get("src", "")}, box, grp)
        elif name == "h1" or has_cls(ch, "chapter-title"):
            add(blocks, {"style": "h1", "runs": inline_runs(ch)}, box, grp)
        elif name == "h2" or has_cls(ch, "section-title"):
            add(blocks, {"style": "h2", "runs": inline_runs(ch)}, box, grp)
        elif name in ("h3", "h4"):
            add(blocks, {"style": "h3", "runs": inline_runs(ch)}, box, grp)
        elif has_cls(ch, "cover-title"):
            add(blocks, {"style": "covertitle", "runs": inline_runs(ch)}, box, grp)
        elif has_cls(ch, "cover-small") or has_cls(ch, "toc-foot"):
            add(blocks, {"style": "small", "runs": inline_runs(ch)}, box, grp)
        elif has_cls(ch, "cover-sub"):
            add(blocks, {"style": "h3c", "runs": inline_runs(ch)}, box, grp)
        elif has_cls(ch, "sub-title") or has_cls(ch, "sub-label") or has_cls(ch, "grading-title") or has_cls(ch, "sol-num"):
            add(blocks, {"style": "subtitle", "runs": inline_runs(ch)}, box, grp)
        elif has_cls(ch, "note"):
            add(blocks, {"style": "note", "runs": inline_runs(ch)}, box, grp)
        elif has_cls(ch, "src"):
            add(blocks, {"style": "src", "runs": inline_runs(ch)}, box, grp)
        elif name == "p":
            add(blocks, {"style": "body", "runs": inline_runs(ch)}, box, grp)
        elif name == "li":
            add(blocks, {"style": "li", "runs": inline_runs(ch)}, box, grp)
        else:
            walk_blocks(ch, blocks, gctr, box, grp)

def sheet_blocks(sheet, gctr):
    body = sheet.select_one(".sheet-body") or sheet.select_one(".cover-wrap") or sheet
    blocks = []
    walk_blocks(body, blocks, gctr)
    return blocks

# ---------------- 단위 ----------------
def mm2hwp(mm):
    return int(round(mm * 7200.0 / 25.4))

def png_size(data):
    if data[:8] == b"\x89PNG\r\n\x1a\n" and data[12:16] == b"IHDR":
        return struct.unpack(">II", data[16:24])
    return None

# ---------------- COM 빌드 ----------------
class Builder:
    def __init__(self):
        self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        try: self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception: pass
        try: self.hwp.SetMessageBoxMode(0x00020000)
        except Exception: pass
        self.hwp.XHwpWindows.Item(0).Visible = False
        self.act = self.hwp.HAction
        self.tmpfiles = []
        self.page_setup()

    def page_setup(self):
        p = self.hwp.HParameterSet.HSecDef
        self.act.GetDefault("PageSetup", p.HSet)
        pd = p.PageDef
        pd.PaperWidth = mm2hwp(210); pd.PaperHeight = mm2hwp(297)
        pd.LeftMargin = mm2hwp(22); pd.RightMargin = mm2hwp(22)
        pd.TopMargin = mm2hwp(20); pd.BottomMargin = mm2hwp(18)
        pd.HeaderLen = 0; pd.FooterLen = 0; pd.GutterLen = 0
        try: self.act.Execute("PageSetup", p.HSet)
        except Exception as e: print("PageSetup 실패:", e)

    def rgb(self, c):
        try: return self.hwp.RGBColor(*c)
        except Exception: return c[0] | (c[1] << 8) | (c[2] << 16)

    def lt_solid(self):
        try: return self.hwp.HwpLineType("Solid")
        except Exception: return 1

    def lw(self, width):
        try: return self.hwp.HwpLineWidth(width)
        except Exception: return 9

    def set_char(self, face, size, bold=False, color=(0, 0, 0)):
        p = self.hwp.HParameterSet.HCharShape
        self.act.GetDefault("CharShape", p.HSet)
        for attr in ("FaceNameHangul", "FaceNameLatin", "FaceNameHanja",
                     "FaceNameJapanese", "FaceNameOther", "FaceNameSymbol", "FaceNameUser"):
            try: setattr(p, attr, face)
            except Exception: pass
        p.Height = size; p.Bold = 1 if bold else 0; p.Italic = 0
        try: p.TextColor = self.rgb(color)
        except Exception: pass
        self.act.Execute("CharShape", p.HSet)

    def set_para(self, st):
        p = self.hwp.HParameterSet.HParaShape
        self.act.GetDefault("ParagraphShape", p.HSet)
        try:
            p.LineSpacingType = 0; p.LineSpacing = st.get("lh", 160)
        except Exception: pass
        try:
            p.PrevSpacing = mm2hwp(st.get("before", 0))
            p.NextSpacing = mm2hwp(st.get("after", 0))
        except Exception: pass
        # 테두리/배경 해제(표를 쓰므로 문단 테두리 불필요)
        bf = p.BorderFill
        for side in ("Top", "Bottom", "Left", "Right"):
            try: setattr(bf, "BorderType" + side, 0)
            except Exception: pass
        try: bf.FillAttr.type = 0
        except Exception: pass
        try: p.BorderConnect = 0
        except Exception: pass
        self.act.Execute("ParagraphShape", p.HSet)
        amap = {"left": "ParagraphShapeAlignLeft", "center": "ParagraphShapeAlignCenter",
                "right": "ParagraphShapeAlignRight", "justify": "ParagraphShapeAlignJustify"}
        self.act.Run(amap.get(st.get("align", "justify"), "ParagraphShapeAlignJustify"))

    def itext(self, t):
        t = re.sub(r"\s+", " ", t)
        if t == "": return
        p = self.hwp.HParameterSet.HInsertText
        self.act.GetDefault("InsertText", p.HSet); p.Text = t
        self.act.Execute("InsertText", p.HSet)

    def ieq(self, script):
        p = self.hwp.HParameterSet.HEqEdit
        self.act.GetDefault("EquationCreate", p.HSet); p.string = script
        self.act.Execute("EquationCreate", p.HSet)
        self.act.Run("Cancel")   # 수식 빠져나옴 (MoveRight는 표 셀에서 칸 밀림 유발 → 제거)

    def newpara(self): self.act.Run("BreakPara")
    def pagebreak(self): self.act.Run("BreakPage")

    def render_runs(self, runs, face, size):
        bface = bold_face(face)
        for kind, val, attrs in runs:
            cite = "cite" in attrs; bold = "bold" in attrs
            rsize = 850 if cite else size
            rcolor = (0x33, 0x33, 0x33) if cite else (0, 0, 0)
            rface = bface if bold else face
            if kind == "br":
                self.newpara()
            elif kind == "text":
                self.set_char(rface, rsize, color=rcolor); self.itext(val)
            elif kind == "math":
                self.set_char(face, size, color=(0, 0, 0))
                try: sc = to_hwp(val)
                except Exception: sc = ""
                if sc.strip(): self.ieq(sc)

    # --- 표 ---
    def create_table(self, rows, cols, colw_mm, border_w="0.4mm", border_c=(0, 0, 0)):
        p = self.hwp.HParameterSet.HTableCreation
        self.act.GetDefault("TableCreate", p.HSet)
        p.Rows = rows; p.Cols = cols
        p.WidthType = 2; p.HeightType = 1
        colw = [mm2hwp(w) for w in colw_mm]
        try: p.WidthValue = sum(colw)
        except Exception: pass
        p.CreateItemArray("ColWidth", cols)
        for i in range(cols): p.ColWidth.SetItem(i, colw[i])
        p.CreateItemArray("RowHeight", rows)
        for i in range(rows): p.RowHeight.SetItem(i, mm2hwp(5))
        self.act.Execute("TableCreate", p.HSet)
        # 기본 표 테두리 사용(한글에서 두께/색 직접 편집 가능). 커서는 (0,0) 셀.

    def next_cell(self): self.act.Run("TableRightCell")

    def escape_table(self):
        # 표(셀) 밖 본문으로 확실히 나갈 때까지 아래로 이동
        self.act.Run("TableColEnd")           # 마지막 셀
        for _ in range(300):
            try:
                pos = self.hwp.GetPos()        # (list, para, pos): 본문 list==0
            except Exception:
                pos = None
            if pos is not None and pos[0] == 0:
                break
            self.act.Run("MoveDown")
        self.act.Run("MoveLineEnd")

    def cell_content(self, blocks):
        """현재 셀에 blocks 렌더 (블록 사이 BreakPara)."""
        for k, blk in enumerate(blocks):
            if k > 0:
                self.newpara()
            self.render_content(blk)

    def render_content(self, blk):
        style = blk["style"]
        if style == "image":
            self.insert_image(blk["src"]); return
        base = STYLE.get(style, STYLE["body"])
        st = dict(base); st["before"] = 0; st["after"] = 0   # 셀 안은 간격 0
        self.set_para(st)
        face, size = st["face"], st["size"]
        if style == "formula":
            self.set_char(face, size);
            try: sc = to_hwp(blk["latex"])
            except Exception: sc = ""
            if sc.strip(): self.ieq(sc)
            return
        if style == "note":
            self.set_char(face, size); self.itext("※ ")
        self.render_runs(blk["runs"], face, size)

    # --- 일반 블록 ---
    def insert_image(self, src):
        m = re.match(r"data:image/([\w+]+);base64,(.*)", src, re.S)
        if not m: return
        ext = m.group(1).split("+")[0]
        if ext == "jpeg": ext = "jpg"
        data = base64.b64decode(m.group(2))
        fd, path = tempfile.mkstemp(suffix="." + ext); os.close(fd)
        open(path, "wb").write(data); self.tmpfiles.append(path)
        self.set_para({"align": "center", "lh": 130, "before": 1.5, "after": 1.5})
        sz = png_size(data); w_mm, h_mm = 72, 0
        if sz:
            w, h = sz; h_mm = 72.0 * h / w
        try:
            self.hwp.InsertPicture(path, True, 0, False, 0, 0, mm2hwp(w_mm), mm2hwp(h_mm) if h_mm else 0)
        except Exception as e:
            print("  [이미지 삽입 실패]", e)
        self.newpara()

    def render_plain(self, blk):
        style = blk["style"]
        if style == "image":
            self.insert_image(blk["src"]); return
        st = STYLE.get(style, STYLE["body"])
        self.set_para(st)
        face, size = st["face"], st["size"]
        if style == "note":
            self.set_char(face, size); self.itext("※ ")
        self.render_runs(blk["runs"], face, size)
        self.newpara()

    def render_box_table(self, box, grp):
        cfg = BOX.get(box, BOX["formula"])
        self.set_para({"align": "justify", "before": 1.5, "after": 0})  # 표 위 간격
        self.create_table(1, 1, [TEXT_W_MM], cfg["width"], cfg["color"])
        self.cell_content(grp)
        self.escape_table()
        self.set_para({"align": "justify", "before": 1.5, "after": 0})  # 표 아래 간격

    def render_grading_table(self, blk):
        cfg = BOX["grading"]
        rows = blk["rows"]; nrows = 1 + len(rows)
        self.set_para({"align": "justify", "before": 1.5, "after": 0})
        self.create_table(nrows, 2, [TEXT_W_MM - 22, 22], cfg["width"], cfg["color"])
        # 행0 두 칸 병합 (제목이 가로 전체 차지)
        merged = False
        try:
            self.act.Run("TableCellBlock")
            self.act.Run("TableCellBlockExtend")
            self.act.Run("TableRightCell")
            self.act.Run("TableMergeCell")
            merged = True
        except Exception:
            pass
        # 제목
        self.render_content({"style": "subtitle", "runs": [("text", blk["title"], set())]})
        if not merged:
            self.next_cell()  # 병합 안 됐으면 (0,1) 건너뜀
        for srun, ptxt in rows:
            self.next_cell()
            st = dict(STYLE["li"]); st["before"] = 0; st["after"] = 0
            self.set_para(st)
            self.render_runs(srun, F_BAT_R, STYLE["li"]["size"])
            self.next_cell()
            self.set_char(F_BAT_R, STYLE["li"]["size"]); self.itext(ptxt)
        self.escape_table()
        self.set_para({"align": "justify", "before": 1.5, "after": 0})

    def build(self, sheets_blocks):
        for i, blocks in enumerate(sheets_blocks):
            j = 0; n = len(blocks)
            while j < n:
                blk = blocks[j]
                if blk["style"] == "gradingtable":
                    self.render_grading_table(blk); j += 1; continue
                box = blk.get("box")
                if box:
                    grp = blk.get("grp")
                    run = []
                    while j < n and blocks[j].get("box") == box and blocks[j].get("grp") == grp:
                        run.append(blocks[j]); j += 1
                    self.render_box_table(box, run)
                else:
                    self.render_plain(blk); j += 1
            if i != len(sheets_blocks) - 1:
                self.pagebreak()

    def save(self, hwpx, pdf=None):
        for path in [p for p in (hwpx, pdf) if p]:
            if os.path.exists(path):
                os.remove(path)
        self.hwp.SaveAs(hwpx, "HWPX", "")
        ok_pdf = False
        if pdf:
            try:
                self.hwp.SaveAs(pdf, "PDF", ""); ok_pdf = os.path.exists(pdf)
            except Exception as e:
                print("PDF 실패:", e)
        self.hwp.Quit()
        for t in self.tmpfiles:
            try: os.remove(t)
            except Exception: pass
        return os.path.exists(hwpx), ok_pdf


def main():
    ap = argparse.ArgumentParser(description="HTML(수학교재) -> 한글 HWPX 변환기")
    ap.add_argument("html", help="변환할 HTML 파일 경로")
    ap.add_argument("--out", default=None, help="출력 hwpx 경로(미지정 시 자동)")
    ap.add_argument("--no-pdf", action="store_true", help="PDF 미리보기 생성 안 함")
    ap.add_argument("--start", type=int, default=0, help="시작 시트(0부터)")
    ap.add_argument("--end", type=int, default=None, help="끝 시트(미포함)")
    args = ap.parse_args()

    src = os.path.abspath(args.html)
    if not os.path.exists(src):
        print("[오류] HTML 파일을 찾을 수 없음:", src); sys.exit(1)

    if args.out:
        out_hwpx = os.path.abspath(args.out)
    else:
        base = os.path.splitext(os.path.basename(src))[0]
        ts = datetime.datetime.now().strftime("%Y-%m-%d_%H%M")
        out_hwpx = os.path.join(os.path.dirname(src), f"{base}_한글변환_{ts}.hwpx")
    out_pdf = None if args.no_pdf else os.path.splitext(out_hwpx)[0] + ".pdf"

    print("입력:", src)
    html = open(src, encoding="utf-8").read()
    soup = BeautifulSoup(html, "lxml")
    for t in soup(["script", "style"]):
        t.decompose()
    sheets = soup.select(".sheet")
    if not sheets:
        print("[경고] .sheet 요소가 없음 — mathbook 양식 HTML이 아닐 수 있음. 전체를 1장으로 처리.")
        sheets = [soup.body or soup]
    end = args.end if args.end is not None else len(sheets)
    sel = sheets[args.start:end]
    print(f"시트 {args.start}~{end-1} ({len(sel)}개) 변환 중...")
    gctr = Ctr()
    sb = [sheet_blocks(sh, gctr) for sh in sel]
    print(f"  총 블록 {sum(len(b) for b in sb)}개")
    b = Builder()
    b.build(sb)
    ok, okpdf = b.save(out_hwpx, out_pdf)
    print("─" * 50)
    print("HWPX 생성:", "성공" if ok else "실패", "→", out_hwpx)
    if out_pdf:
        print("PDF  생성:", "성공" if okpdf else "실패", "→", out_pdf)
    print("완료.")


if __name__ == "__main__":
    main()
