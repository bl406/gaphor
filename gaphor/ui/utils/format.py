# -*- coding: utf-8 -*-
"""
v3 fulltable • fix list again & stronger CenteredTitle
- 列表：优先使用 Word 的 ListString 作为“真实显示标签”（绝大多数模板最稳）；
        若 ListString 为空，再用 NumberStyle + ListValue 推导；
        先 RemoveNumbers()，再检查段首是否已有该标签，若没有再插入，避免“a. a.”。
- 封面/居中：在文首将连续的标题行（多段 ^p）合并为一个段（^p→^l），并套用 CenteredTitle；
             关键词匹配“应用指南/航天器用”，或段落长度较短且为文首若干段。
- 表格：逐段处理（不跳过），表内“表1/表A.1…”→表注，其余→表中文。
- 其它保持：图片+^l+图/表注、普通 ^l→^p、A.1/A.1.1… 层级。
"""
import argparse, re, sys, os
from pathlib import Path
import win32com.client as win32
from win32com.client import constants as c

RE_HAS_CJK      = re.compile(r"[\u4e00-\u9fff]")
RE_HEADING_NUM  = re.compile(r"^\s*(\d+(?:\.\d+){0,4})(?:\s|[．\.\-、：:])?\s*")
RE_APPENDIX_HDR = re.compile(r"^\s*附录\s*([A-ZＡ-Ｚ])\s*$")
RE_APPX_SEC     = re.compile(r"^\s*[A-ZＡ-Ｚ]\s*[\.．]\s*\d+(?:[\.．]\d+)*")

RE_FIGCAP_NUM   = re.compile(r"^\s*(图|Fig\.?|Figure)\s*[：:\.]?\s*\d+(?:[-．\.]\d+)*", re.I)
RE_TABCAP_NUM   = re.compile(r"^\s*(表|Tab\.?|Table)\s*[：:\.]?\s*\d+(?:[-．\.]\d+)*", re.I)
RE_FIGCAP_APPX  = re.compile(r"^\s*(图|Fig\.?|Figure)\s*[A-ZＡ-Ｚ]\s*[\.．]\s*\d+(?:[-．\.]\d+)*", re.I)
RE_TABCAP_APPX  = re.compile(r"^\s*(表|Tab\.?|Table)\s*[A-ZＡ-Ｚ]\s*[\.．]\s*\d+(?:[-．\.]\d+)*", re.I)
RE_MANUAL_MARK  = re.compile(r"^\s*([A-Za-z]+|\d+)\s*([\)\.、．])\s+")

SOFT_LF = "\x0b"

def ptext(par): 
    return par.Range.Text.replace("\r","").replace("\x07","").strip()

def within_table(par):
    try: return bool(par.Range.Information(c.wdWithInTable))
    except Exception:
        try: return par.Range.Cells.Count > 0 or par.Range.Tables.Count > 0
        except Exception: return False

def has_caption_field(par):
    try:
        for f in par.Range.Fields:
            code = f.Code.Text if hasattr(f, "Code") else ""
            if "SEQ" in code and (" 图" in code or "表" in code or " Figure" in code or " Table" in code):
                return True
    except Exception: pass
    return False

def load_style_map(path):
    if not path: return None
    import json; p = Path(path)
    with p.open("r", encoding="utf-8") as f: return json.load(f)

def attach_template(doc, template_path):
    if not template_path: return None
    p = Path(template_path)
    if not p.exists(): raise FileNotFoundError(f"DOTM 模板不存在: {p}")
    doc.AttachedTemplate = str(p); doc.UpdateStylesOnOpen = True
    return str(p)

def ensure_style_from_template(app, doc, template_path, style_name):
    try: return doc.Styles(style_name)
    except Exception: pass
    if not template_path: raise FileNotFoundError(f"需要模板以获取样式：{style_name}")
    dest = doc.FullName
    if not dest or not os.path.exists(dest):
        tmp = Path(doc.Path or os.getcwd()) / (Path(doc.Name).stem + "_tmp_for_stylecopy.docx")
        doc.SaveAs2(str(tmp)); dest = str(tmp)
    app.OrganizerCopy(Source=template_path, Destination=dest, Name=style_name, Object=c.wdOrganizerObjectStyles)
    return doc.Styles(style_name)

def resolve_style(app, doc, tmpl, style_map, key, level=None):
    name = None
    if style_map:
        if key == "H":
            hm = style_map.get("Heading", {}); name = (hm.get(str(level)) or hm.get(level))
        else:
            name = style_map.get(key)  # Body / FigCaption / TabCaption / TableBody / CenteredTitle
    if name:
        try: return doc.Styles(name)
        except Exception:
            try: return ensure_style_from_template(app, doc, tmpl, name)
            except Exception as e: print(f"[!] 获取样式失败 {name}: {e}")
    if key == "H":
        try: return doc.Styles(getattr(c, f"wdStyleHeading{level or 1}"))
        except Exception:
            for nm in (f"标题 {level or 1}", f"Heading {level or 1}"):
                try: return doc.Styles(nm)
                except Exception: continue
            return doc.Styles.Add(f"Heading {level or 1}", c.wdStyleTypeParagraph)
    try: return doc.Styles(c.wdStyleNormal)
    except Exception: return doc.Styles.Add("正文", c.wdStyleTypeParagraph)

def expected_level_from_text(txt):
    mA = RE_APPX_SEC.match(txt or "")
    if mA:
        rest = re.sub(r"^\s*[A-ZＡ-Ｚ]\s*[\.．]\s*", "", mA.group(0))
        depth = 1 + len([p for p in re.split(r"[\.．]", rest) if p.strip()])
        return min(max(depth, 2), 5)
    m = RE_HEADING_NUM.match(txt or "")
    if m: return min(1 + m.group(1).count("."), 5)
    return None

def split_soft_in_paragraph(par):
    rng = par.Range.Duplicate; f = rng.Find
    f.ClearFormatting(); f.Text = "^l"; f.Replacement.ClearFormatting(); f.Replacement.Text = "^p"
    f.Forward = True; f.Wrap = c.wdFindStop; f.Execute(Replace=c.wdReplaceAll)

def merge_para_to_soft_range(rng):
    f = rng.Find; f.ClearFormatting(); f.Text = "^p"
    f.Replacement.ClearFormatting(); f.Replacement.Text = "^l"
    f.Forward = True; f.Wrap = c.wdFindStop; f.Execute(Replace=c.wdReplaceAll)

def try_merge_centered_block(doc, i):
    """把“附录X / （资料性附录） / 信息”三行合并为一个段（^p→^l）"""
    try:
        p1 = doc.Paragraphs(i);  t1 = ptext(p1)
        if not RE_APPENDIX_HDR.match(t1): return False
        p2 = doc.Paragraphs(i+1); t2 = ptext(p2)
        p3 = doc.Paragraphs(i+2); t3 = ptext(p3)
    except Exception: return False
    if ("资料性" in t2 or "附录" in t2) and ("信息" in t3 or t3.endswith("信息")):
        rng = doc.Range(Start=p1.Range.Start, End=p3.Range.End)
        merge_para_to_soft_range(rng)
        return True
    return False

def make_cover_centered(doc, app, tmpl, style_map):
    """文首‘封面标题’多段合并为单段 ^l，并应用 CenteredTitle"""
    max_scan = min(12, doc.Paragraphs.Count)
    # 寻找包含关键词的第一段
    hit_i = None
    for i in range(1, max_scan+1):
        t = ptext(doc.Paragraphs(i))
        if "应用指南" in t or "航天器用" in t:
            hit_i = i; break
    if hit_i is None:
        return
    # 合并直到遇到空段或出现正文较长句子的段
    start = hit_i; end = hit_i
    for j in range(hit_i+1, min(hit_i+6, doc.Paragraphs.Count)+1):
        tj = ptext(doc.Paragraphs(j))
        if not tj: break
        # 标题行通常较短（< 30 字）
        if len(tj) > 30: break
        end = j
    rng = doc.Range(Start=doc.Paragraphs(start).Range.Start, End=doc.Paragraphs(end).Range.End)
    merge_para_to_soft_range(rng)
    st = resolve_style(app, doc, tmpl, style_map, "CenteredTitle", level=1)
    doc.Paragraphs(start).Range.Style = st
    try: doc.Paragraphs(start).OutlineLevel = 1
    except Exception: pass

def list_label_from_word(lf):
    """优先使用 Word 渲染的 ListString；若为空再回退推导"""
    try:
        label = lf.ListString
        if label: return label
    except Exception:
        label = ""
    # 回退：依据 NumberStyle + ListValue + NumberFormat
    try:
        lvl = lf.ListLevelNumber if lf.ListLevelNumber else 1
        fmt = lf.ListTemplate.ListLevels(lvl).NumberFormat or "%1."
        style = lf.ListTemplate.ListLevels(lvl).NumberStyle
        val = lf.ListValue or 0
    except Exception:
        return ""
    # 标点
    punct = "."
    for ch in (")","）","、","．","."):
        if ch in fmt: punct = ch; break
    # 生成核心编号
    def int_to_roman(n, upper=False):
        nums = [(1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'),
                (100, 'C'), (90, 'XC'), (50, 'L'), (40, 'XL'),
                (10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')]
        s = ""
        for v, sym in nums:
            while n >= v:
                s += sym; n -= v
        return s if upper else s.lower()
    try:
        if style in (c.wdListNumberStyleArabic, c.wdListNumberStyleArabicFullWidth):
            core = str(val)
        elif style in (c.wdListNumberStyleLowercaseLetter,):
            core = chr(ord('a') + val - 1) if 1 <= val <= 26 else str(val)
        elif style in (c.wdListNumberStyleUppercaseLetter,):
            core = chr(ord('A') + val - 1) if 1 <= val <= 26 else str(val)
        elif style in (c.wdListNumberStyleLowercaseRoman,):
            core = int_to_roman(val, False)
        elif style in (c.wdListNumberStyleUppercaseRoman,):
            core = int_to_roman(val, True)
        else:
            core = str(val)
    except Exception:
        core = str(val)
    return f"{core}{punct}"

def is_auto_list(par):
    try:
        lf = par.Range.ListFormat
        return bool(lf and lf.ListType != c.wdListNoNumbering)
    except Exception:
        return False

def is_manual_prefix_text(txt):
    return RE_MANUAL_MARK.match(txt or "") is not None


def ensure_manual_not_list(par):
    try:
        lf = par.Range.ListFormat
        # c.wdListNoNumbering 有时不可用，用 0 兜底（wdListNoNumbering == 0）
        list_type = getattr(lf, "ListType", 0) if lf else 0
        if lf and list_type != 0:
            lf.RemoveNumbers()
    except Exception:
        pass
    

def collect_series(doc, start_idx):
    """从 start_idx 开始收集连续的（手动编号 or 自动列表）段，返回 end_idx。手动优先。"""
    n = doc.Paragraphs.Count
    i = start_idx
    while i <= n:
        p = doc.Paragraphs(i)
        t = ptext(p)

        # ① 手动编号优先（并确保不是列表状态）
        if is_manual_prefix_text(t):
            ensure_manual_not_list(p)
            i += 1
            continue

        # ② 自动列表（哪怕文本为空也纳入）
        try:
            lf = p.Range.ListFormat
            if lf and lf.ListType != c.wdListNoNumbering:
                i += 1
                continue
        except Exception:
            pass

        break
    return i - 1


# 仅匹配“只有编号”的段：  a.  /  a)  /  1.  /  1)  /  …（后面没有其它文字）
_RE_LABEL_ONLY = re.compile(r"^\s*([A-Za-z]+|\d+)\s*([\)\.、．])\s*$")

def convert_auto_run_to_text(doc, start_idx, end_idx):
    # 逆序固化，防止连锁变化
    for i in range(end_idx, start_idx - 1, -1):
        p = doc.Paragraphs(i)
        try:
            p.Range.ListFormat.ConvertNumbersToText()
        except Exception:
            pass
        # —— 新增：清除“只有编号”的段 —— #
        txt = ptext(p)
        if _RE_LABEL_ONLY.match(txt):
            # 只清内容，不删段落标记
            rng = p.Range.Duplicate
            try:
                rng.End = rng.End - 1  # 去掉最后的段落标记再设回
                rng.Text = ""
            except Exception:
                # 兜底（极少数情况下上面会失败）
                pass


def apply_series_block(doc, app, tmpl, style_map, start_idx, end_idx):
    """把编号 series 拆成 auto-run / manual-run 分段处理；文本样式统一改 Body，OutlineLevel=10。"""
    i = start_idx
    while i <= end_idx:
        p = doc.Paragraphs(i)
        t = ptext(p)
        if is_auto_list(p):
            j = i
            while j <= end_idx and is_auto_list(doc.Paragraphs(j)):
                j += 1
            convert_auto_run_to_text(doc, i, j - 1)
            for k in range(i, j):
                st = resolve_style(app, doc, tmpl, style_map, "Body")
                doc.Paragraphs(k).Range.Style = st
                try: doc.Paragraphs(k).OutlineLevel = 10
                except Exception: pass
            i = j
        else:
            # manual-run：不改文本，只套样式
            j = i
            while j <= end_idx and is_manual_prefix_text(ptext(doc.Paragraphs(j))):
                j += 1
            for k in range(i, j):
                st = resolve_style(app, doc, tmpl, style_map, "Body")
                doc.Paragraphs(k).Range.Style = st
                try: doc.Paragraphs(k).OutlineLevel = 10
                except Exception: pass
            i = j

def has_picture(par):
    return par.Range.InlineShapes.Count


def docx_format(in_path=None, out_path=None, template="./ui/utils/template.docx", style_map="./ui/utils/style_map.json"):
    print(Path(style_map).absolute())
    print(Path(template).absolute())
    in_path = Path(in_path).resolve()
    out_path = Path(out_path).resolve() if out_path else in_path.with_name(in_path.stem + "_formatted.docx")
    style_map = load_style_map(style_map)
    
    # app = win32.DispatchEx("Word.Application"); app.Visible = True
    app = win32.gencache.EnsureDispatch("Word.Application"); app.Visible = True
    try:
        doc = app.Documents.Open(str(in_path)); doc.Activate()
        tmpl = attach_template(doc, template)

        # 文首封面标题合并并居中
        make_cover_centered(doc, app, tmpl, style_map)

        n = doc.Paragraphs.Count; i = 1
        while i <= n:
            par = doc.Paragraphs(i); txt = ptext(par)

            # 附录复合块合并
            if RE_APPENDIX_HDR.match(txt):
                if try_merge_centered_block(doc, i):
                    n = doc.Paragraphs.Count
                    par = doc.Paragraphs(i); txt = ptext(par)

            # 居中类（含封面/附录组合）
            if SOFT_LF in txt and (RE_APPENDIX_HDR.match(txt.split(SOFT_LF)[0]) or "应用指南" in txt or "航天器用" in txt):
                st = resolve_style(app, doc, tmpl, style_map, "CenteredTitle", level=1)
                par.Range.Style = st
                try: par.OutlineLevel = 1
                except Exception: pass
                i += 1; continue

            # 图片 + ^l + 图/表注
            try:
                if par.Range.InlineShapes.Count > 0 and SOFT_LF in txt:
                    tail = txt.rsplit(SOFT_LF, 1)[1].strip()
                    if RE_FIGCAP_NUM.match(tail) or RE_TABCAP_NUM.match(tail) or RE_FIGCAP_APPX.match(tail) or RE_TABCAP_APPX.match(tail):
                        split_soft_in_paragraph(par); n = doc.Paragraphs.Count; par = doc.Paragraphs(i); txt = ptext(par)
            except Exception: pass

            # 普通软回车拆段
            if SOFT_LF in txt:
                split_soft_in_paragraph(par); n = doc.Paragraphs.Count; par = doc.Paragraphs(i); txt = ptext(par)

            # —— 新：列表 series（自动 or 手动编号）块处理 —— 
            if is_auto_list(par) or is_manual_prefix_text(txt):
                end_idx = collect_series(doc, i)
                apply_series_block(doc, app, tmpl, style_map, i, end_idx)
                i = end_idx + 1
                continue
        

            # 图/表注
            tnow = ptext(par)
            if tnow:
                if RE_FIGCAP_NUM.match(tnow) or RE_FIGCAP_APPX.match(tnow) or has_caption_field(par):
                    st = resolve_style(app, doc, tmpl, style_map, "FigCaption")
                    par.Range.Style = st; 
                    try: par.OutlineLevel = 10
                    except Exception: pass
                    i += 1; continue
                if RE_TABCAP_NUM.match(tnow) or RE_TABCAP_APPX.match(tnow):
                    st = resolve_style(app, doc, tmpl, style_map, "TabCaption")
                    par.Range.Style = st; 
                    try: par.OutlineLevel = 10
                    except Exception: pass
                    i += 1; continue

            # 表内正文（初版策略）
            if within_table(par):
                st = resolve_style(app, doc, tmpl, style_map, "TableBody")
                par.Range.Style = st
                try: par.OutlineLevel = 10
                except Exception: pass
                i += 1; continue

            # 图片段落
            if has_picture(par):
                st = resolve_style(app, doc, tmpl, style_map, "ImagePara")
                par.Range.Style = st
                try: par.OutlineLevel = 10
                except Exception: pass
                i += 1; continue

            # 标题/正文（含 A.1/A.1.1）
            exp = expected_level_from_text(tnow)
            if exp and RE_HAS_CJK.search(tnow):
                st = resolve_style(app, doc, tmpl, style_map, "H", level=exp)
                par.Range.Style = st; 
                try: par.OutlineLevel = exp
                except Exception: pass
            else:
                try: ol = int(par.OutlineLevel)
                except Exception: ol = 10
                if 1 <= ol <= 5:
                    st = resolve_style(app, doc, tmpl, style_map, "H", level=ol)
                    par.Range.Style = st; 
                    try: par.OutlineLevel = ol
                    except Exception: pass
                else:
                    st = resolve_style(app, doc, tmpl, style_map, "Body")
                    par.Range.Style = st; 
                    try: par.OutlineLevel = 10
                    except Exception: pass

            i += 1

        doc.SaveAs2(str(out_path)); print(f"[OK] 已保存：{out_path}")
    finally:
        try: doc.Close(SaveChanges=False)
        except Exception: pass
        app.Quit()

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("-i","--input", default=r"D:\Study\本子\文档识别生成\para_new\test\test.docx", required=False)
    ap.add_argument("-o","--output", default=r"./ntest", required=False)
    ap.add_argument("-t","--template", default=r"E:\Gaphor\gaphor\gaphor\ui\utils\template.docx", required=False)
    ap.add_argument("-m","--style-map", default=r"E:\Gaphor\gaphor\gaphor\ui\utils\style_map.json", required=False)
    args = ap.parse_args()
    docx_format(args.input, args.output, args.template, args.style_map)
