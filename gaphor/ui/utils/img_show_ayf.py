from docx import Document
from docx.oxml.ns import qn
import re
import unicodedata


# == 图注匹配规则 ==
RE_FIGCAP_APPX  = re.compile(r"^\s*(图|Fig\.?|Figure)\s*[A-ZＡ-Ｚ]\s*[\.\．]\s*\d+(?:[-．\.]\d+)*", re.I)
RE_FIGCAP_NUM   = re.compile(r"^\s*(图|Fig\.?|Figure)\s*[：:\.]?\s*\d+(?:[-．\.]\d+)*", re.I)

def find_style_node_no_ns(doc, style_id: str):
    styles_root = doc.styles.element  # <w:styles>
    xp = f".//*[local-name()='style' and @*[local-name()='styleId']='{style_id}']"
    nodes = styles_root.xpath(xp)
    return nodes[0] if nodes else None

def get_style_outline_no_ns(doc, style_id: str):
    st = find_style_node_no_ns(doc, style_id)
    if st is None:
        return None
    val = st.xpath(".//*[local-name()='pPr']/*[local-name()='outlineLvl']/@*[local-name()='val']")
    return int(val[0]) if val else None

def _extract_inline_images_from_paragraph(paragraph):
    """
    直接读取段内 inline 图片二进制 (part.rels -> blob)
    返回 list[bytes]
    """
    images = []
    pxml = paragraph._p
    for el in pxml.iter():
        if el.tag.endswith('blip'):
            rId = el.get(qn('r:embed')) or el.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if not rId:
                continue
            rel = paragraph.part.rels.get(rId)
            if not rel or not rel.target_part:
                continue
            blob = rel.target_part.blob
            try:
                # 若需 PIL 可自己打开；此处保留二进制
                images.append(blob)
            except Exception:
                pass
    return images

def find_node_place(node, parent):
    """把 node 按层级挂到 parent 子树"""
    if parent['outline_level'] == node['outline_level'] - 1:
        parent['children'].append(node)
        node['parent'] = parent
    elif parent['outline_level'] > node['outline_level'] - 1 and parent.get('parent') is not None:
        return find_node_place(node, parent['parent'])
    else:
        # 允许“退回根部”的情况
        raise ValueError("文档结构错误，请检查标题层级")
    return

def extract_images_with_captions(doc_path):
    """
    扫描文档，构建树 root，标题为节点，图片作为 ('image', blob, caption, para_idx) 挂在节点 data
    """
    from docx.text.paragraph import Paragraph  # 延迟导入以防环境问题
    doc = Document(doc_path)
    paras = doc.paragraphs

    root = {
        'text': '',
        'para_id': 0,
        'outline_level': 0,
        'parent': None,
        'children': [],
        'data': []
    }
    parent = root
    i = 0

    while i < len(paras):
        p = paras[i]
        # 修复弃用；用样式名而非 style_id
        try:
            outlinelevel = get_outline_level_of_style(p.style)
        except Exception:
            outlinelevel = 10

        if outlinelevel != 10:
            node = {
                'text': p.text,
                'para_id': i,
                'outline_level': outlinelevel,
                'parent': None,
                'children': [],
                'data': []
            }
            find_node_place(node, parent)
            parent = node

        imgs = _extract_inline_images_from_paragraph(p)
        if imgs:
            # 下方第一段非空 & 匹配“图/FIGURE”样式的文本作为图注
            caption = ""
            j = i + 1
            while j < len(paras):
                txt = (extract_para_text_all(paras[j]._p.xml) or "").strip()
                if not txt:
                    j += 1
                    continue
                if RE_FIGCAP_APPX.match(txt) or RE_FIGCAP_NUM.match(txt):
                    caption = txt
                break

            for img in imgs:
                image_para_index = i
                parent['data'].append(('image', img, caption, image_para_index))
        i += 1

    return root

def _extract_images_with_captions(doc_path):
    """
    旧工具函数：返回 (tree, flat_results)，保留以防你调试使用
    """
    doc = Document(doc_path)
    paras = doc.paragraphs
    root = {
        'text': '',
        'para_id': 0,
        'outline_level': 0,
        'parent': None,
        'children': [],
        'data': []
    }
    parent = root
    results = []
    i = 0

    while i < len(paras):
        p = paras[i]
        try:
            outlinelevel = get_outline_level_of_style(doc.styles[p.style.name])
        except Exception:
            outlinelevel = 10

        if outlinelevel != 10:
            node = {
                'text': p.text,
                'para_id': i,
                'outline_level': outlinelevel,
                'parent': None,
                'children': [],
                'data': []
            }
            find_node_place(node, parent)
            parent = node

        imgs = _extract_inline_images_from_paragraph(p)
        if imgs:
            caption = ""
            j = i + 1
            while j < len(paras):
                txt = (extract_para_text_all(paras[j]._p.xml) or "").strip()
                if not txt:
                    j += 1
                    continue
                if RE_FIGCAP_APPX.match(txt) or RE_FIGCAP_NUM.match(txt):
                    caption = txt
                break
            for img in imgs:
                image_para_index = i
                parent['data'].append(('image', img, caption, image_para_index))
                results.append(("image", caption, image_para_index))
        i += 1

    return root, results

def Create_image_tree(doc_path):
    return extract_images_with_captions(doc_path)

def para_has_omml(para) -> bool:
    """段落是否包含原生 Word 公式（OMML）"""
    pxml = para._p
    for el in pxml.iter():
        tag = el.tag.split('}')[-1]
        if tag in ('oMath', 'oMathPara'):
            return True
    return False

def extract_para_text_all(xml_p: str) -> str:
    # 取普通 w:t
    def _wt(s: str) -> str:
        return ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', s, flags=re.DOTALL))
    # 取公式 m:t
    def _mt(s: str) -> str:
        return ''.join(re.findall(r'<m:t[^>]*>(.*?)</m:t>', s, flags=re.DOTALL))

    def _parse_math(math_xml: str) -> str:
        def _repl_subsup(m):
            blk = m.group(0)
            base = _mt(''.join(re.findall(r'<m:e\b[^>]*>(.*?)</m:e>', blk, flags=re.DOTALL)))
            sub  = _mt(''.join(re.findall(r'<m:sub\b[^>]*>(.*?)</m:sub>', blk, flags=re.DOTALL)))
            sup  = _mt(''.join(re.findall(r'<m:sup\b[^>]*>(.*?)</m:sup>', blk, flags=re.DOTALL)))
            return f'<m:t>{base}_{sub}^{sup}</m:t>'
        def _repl_sub(m):
            blk = m.group(0)
            base = _mt(''.join(re.findall(r'<m:e\b[^>]*>(.*?)</m:e>', blk, flags=re.DOTALL)))
            sub  = _mt(''.join(re.findall(r'<m:sub\b[^>]*>(.*?)</m:sub>', blk, flags=re.DOTALL)))
            return f'<m:t>{base}_{sub}</m:t>'
        def _repl_sup(m):
            blk = m.group(0)
            base = _mt(''.join(re.findall(r'<m:e\b[^>]*>(.*?)</m:e>', blk, flags=re.DOTALL)))
            sup  = _mt(''.join(re.findall(r'<m:sup\b[^>]*>(.*?)</m:sup>', blk, flags=re.DOTALL)))
            return f'<m:t>{base}^{sup}</m:t>'

        math_xml = re.sub(r'<m:sSubSup\b[^>]*>.*?</m:sSubSup>', _repl_subsup, math_xml, flags=re.DOTALL)
        math_xml = re.sub(r'<m:sSub\b[^>]*>.*?</m:sSub>', _repl_sub, math_xml, flags=re.DOTALL)
        math_xml = re.sub(r'<m:sSup\b[^>]*>.*?</m:sSup>', _repl_sup, math_xml, flags=re.DOTALL)
        return _mt(math_xml)

    out, last = [], 0
    math_pat = re.compile(r'<m:(oMath|oMathPara)\b[^>]*>.*?</m:\1>', re.DOTALL)
    for m in math_pat.finditer(xml_p):
        out.append(_wt(xml_p[last:m.start()]))
        out.append(_parse_math(m.group(0)))
        last = m.end()
    out.append(_wt(xml_p[last:]))
    return ''.join(out)

def split_figure_caption(s: str):
    s = s.strip()
    if ' ' not in s:
        return s, ''
    idx, cap = s.split(' ', 1)
    return idx.strip(), cap.strip()

# =========================
#  树遍历 & 匹配（多节点 + 题注）
# =========================

def collect_data_from_node(tree_node):
    data_list = list(tree_node.get('data', []))
    for child in tree_node.get('children', []):
        data_list.extend(collect_data_from_node(child))
    return data_list

def _norm_text(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKC", s).lower()
    s = re.sub(r"[，。；：！？”“、（）【】《》]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def find_nodes_by_text_all(node, txt):
    """返回所有标题文本包含 txt 的节点"""
    wanted = _norm_text(txt)
    out = []
    def _dfs(n):
        for ch in n.get('children', []):
            if wanted and _norm_text(ch.get('text') or "").find(wanted) != -1:
                out.append(ch)
            _dfs(ch)
    _dfs(node)
    return out

def collect_data_from_nodes(nodes):
    """合并多个节点 data，并去重"""
    merged, seen = [], set()
    for nd in nodes:
        for d in collect_data_from_node(nd):
            if isinstance(d, (list, tuple)) and len(d) >= 1:
                tag = d[0]
                caption = d[2] if len(d) >= 3 else ""
                idx = d[3] if len(d) >= 4 else None
                key = (tag, idx, _norm_text(caption))
            else:
                key = ("obj", id(d))
            if key not in seen:
                seen.add(key)
                merged.append(d)
    return merged

def collect_all_data(tree_root):
    return collect_data_from_node(tree_root)

def filter_by_caption(data_list, txt):
    """按题注全树过滤"""
    wanted = _norm_text(txt)
    out = []
    seen = set()
    for d in data_list:
        if isinstance(d, (list, tuple)) and len(d) >= 3:
            tag = d[0]
            caption = d[2] or ""
            idx = d[3] if len(d) >= 4 else None
            if wanted and _norm_text(caption).find(wanted) != -1:
                key = (tag, idx, _norm_text(caption))
                if key not in seen:
                    seen.add(key)
                    out.append(d)
    return out

def get_datalist_by_text(image_tree, txt):
    """
    txt 为空：返回整棵树所有 data
    txt 非空：返回
      a) 所有标题包含 txt 的节点子树 data
      b) 全树题注 caption 包含 txt 的 data
    的并集（去重）
    """
    if not txt:
        return collect_data_from_node(image_tree)

    nodes = find_nodes_by_text_all(image_tree, txt)
    by_title = collect_data_from_nodes(nodes) if nodes else []

    all_data = collect_all_data(image_tree)
    by_caption = filter_by_caption(all_data, txt)

    merged, seen = [], set()
    for d in by_title + by_caption:
        if isinstance(d, (list, tuple)) and len(d) >= 1:
            tag = d[0]
            caption = d[2] if len(d) >= 3 else ""
            idx = d[3] if len(d) >= 4 else None
            key = (tag, idx, _norm_text(caption))
        else:
            key = ("obj", id(d))
        if key not in seen:
            seen.add(key)
            merged.append(d)
    return merged

# ============== 其他小工具 ==============

def get_outline_level_of_style(style, para):
    try:
        if para._element.pPr.outlineLvl.val:
            return para._element.pPr.outlineLvl.val + 1
    except:
        try:
            if style.element.pPr.outlineLvl:
                return style.element.pPr.outlineLvl.val + 1
            else:
                return get_outline_level_of_style(style.base_style, None)
        except:
            return 10

def build_outline_tree(items):
    if not items:
        return []
    root = []
    stack = []
    for it in items:
        node = {"text": it["text"], "para_id": it["para_id"], "level": it["level"], "children": []}
        if not stack or it["level"] == 0:
            root.append(node)
            stack = [node]
            continue
        while stack and stack[-1]["level"] >= it["level"]:
            stack.pop()
        if stack:
            stack[-1]["children"].append(node)
        else:
            root.append(node)
        stack.append(node)
    return root



def print_tree(node, indent=0):
    prefix = " " * (indent*4)
    text = (node.get('text') or "").strip() or "<ROOT>"
    para = node.get('para_id', -1)
    img_count = len([d for d in node.get('data', []) if isinstance(d, (list, tuple)) and d and d[0]=='image'])
    img_info = f" 📷x{img_count}" if img_count > 0 else ""
    print(f"{prefix}- [{para}] {text}{img_info}")
    for child in node.get("children", []):
        print_tree(child, indent+1)

if __name__ == "__main__":
    tree = Create_image_tree("/hdd2/ayf/ayf_code/acdm/utils/载人空间站用半导体分立器件CYSR3015C型硅肖特基二极管应用指南_zyh最终修订版.docx")
    datalist = get_datalist_by_text(tree, "器件概况")
    print_tree(tree)
    print("ok")
