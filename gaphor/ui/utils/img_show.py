from docx import Document
from docx.oxml.ns import qn
import re

RE_FIGCAP_APPX  = re.compile(r"^\s*(图|Fig\.?|Figure)\s*[A-ZＡ-Ｚ]\s*[\.\．]\s*\d+(?:[-．\.]\d+)*", re.I)
RE_FIGCAP_NUM   = re.compile(r"^\s*(图|Fig\.?|Figure)\s*[：:\.]?\s*\d+(?:[-．\.]\d+)*", re.I)

def find_style_node_no_ns(doc, style_id: str):
    styles_root = doc.styles.element  # 就是 styles.xml 的 <w:styles>
    xp = f".//*[local-name()='style' and @*[local-name()='styleId']='{style_id}']"
    nodes = styles_root.xpath(xp)
    return nodes[0] if nodes else None

def get_style_outline_no_ns(doc, style_id: str):
    st = find_style_node_no_ns(doc, style_id)
    if st is None:
        return None
    # 找 <pPr>/<outlineLvl> 的 @val
    val = st.xpath(".//*[local-name()='pPr']/*[local-name()='outlineLvl']/@*[local-name()='val']")
    return int(val[0]) if val else None

def _extract_inline_images_from_paragraph(paragraph):
    """
    不用 zip、不用 xpath，直接读图片二进制:
        rel.target_part.blob -> BytesIO -> PIL.Image
    返回 list[PIL.Image]
    """
    images = []
    pxml = paragraph._p

    for el in pxml.iter():
        if el.tag.endswith('blip'):
            # 取关系 rId
            rId = el.get(qn('r:embed')) or el.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if not rId:
                continue

            rel = paragraph.part.rels.get(rId)
            if not rel or not rel.target_part:
                continue

            blob = rel.target_part.blob
            try:
                # img = Image.open(BytesIO(blob))
                images.append(blob)
            except Exception:
                pass

    return images

def find_node_place(node, parent):
    """在 parent 树中找到 node 的位置，返回对应的 parent 节点"""
    if parent['outline_level'] == node['outline_level'] - 1:
        #为当前父的子
        parent['children'].append(node)
        node['parent'] = parent
    elif parent['outline_level'] > node['outline_level'] - 1:
        find_node_place(node, parent['parent'])
    else:
        raise ValueError("文档结构错误，请检查标题层级")
    return 

def extract_images_with_captions(doc_path):
    """
    找所有 inline 图片，并取下一段作为图注（若匹配）
    返回 list[(PIL.Image, caption_text)]
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
    i = 0
    
    while i < len(paras):
        p = paras[i]
        outlinelevel = get_outline_level_of_style(doc.styles[p.style.style_id])

        if outlinelevel != 10:
            #为标题，构建节点
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
                # 如果要分index和caption文本：就启用下面的
                # index, caption_txt = split_figure_caption(caption)
                # results.append((img, index, caption_txt, caption))
                image_para_index = i
                node['data'].append((img, caption, image_para_index))
        i += 1

    return root

def _extract_images_with_captions(doc_path):
    """
    找所有 inline 图片，并取下一段作为图注（若匹配）
    返回 list[(PIL.Image, caption_text)]
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
        outlinelevel = get_outline_level_of_style(doc.styles[p.style.style_id])

        if outlinelevel != 10:
            #为标题，构建节点
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
                # 如果要分index和caption文本：就启用下面的
                # index, caption_txt = split_figure_caption(caption)
                # results.append((img, index, caption_txt, caption))
                image_para_index = i
                node['data'].append((img, caption, image_para_index))
                results.append(("img", caption, image_para_index))
        i += 1

    return root, results

def Create_image_tree(doc_path):
    image_tree = extract_images_with_captions(doc_path)
    return image_tree

def para_has_omml(para) -> bool:
    """段落是否包含原生 Word 公式（OMML）"""
    pxml = para._p
    for el in pxml.iter():
        tag = el.tag.split('}')[-1]  # 去掉命名空间
        if tag in ('oMath', 'oMathPara'):
            return True
    return False

def extract_para_text_all(xml_p: str) -> str:
    # 提取普通 w:t
    def _wt(s: str) -> str:
        return ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', s, flags=re.DOTALL))

    # 提取公式内 m:t
    def _mt(s: str) -> str:
        return ''.join(re.findall(r'<m:t[^>]*>(.*?)</m:t>', s, flags=re.DOTALL))

    # 将一个公式块转换为行内文本：先把子/上/上下标替换成 <m:t>...</m:t>，再统一取 m:t
    def _parse_math(math_xml: str) -> str:
        # 上下标：base_sub^sup
        def _repl_subsup(m):
            blk = m.group(0)
            base = _mt(''.join(re.findall(r'<m:e\b[^>]*>(.*?)</m:e>', blk, flags=re.DOTALL)))
            sub  = _mt(''.join(re.findall(r'<m:sub\b[^>]*>(.*?)</m:sub>', blk, flags=re.DOTALL)))
            sup  = _mt(''.join(re.findall(r'<m:sup\b[^>]*>(.*?)</m:sup>', blk, flags=re.DOTALL)))
            return f'<m:t>{base}_{sub}^{sup}</m:t>'

        # 下标：base_sub
        def _repl_sub(m):
            blk = m.group(0)
            base = _mt(''.join(re.findall(r'<m:e\b[^>]*>(.*?)</m:e>', blk, flags=re.DOTALL)))
            sub  = _mt(''.join(re.findall(r'<m:sub\b[^>]*>(.*?)</m:sub>', blk, flags=re.DOTALL)))
            return f'<m:t>{base}_{sub}</m:t>'

        # 上标：base^sup
        def _repl_sup(m):
            blk = m.group(0)
            base = _mt(''.join(re.findall(r'<m:e\b[^>]*>(.*?)</m:e>', blk, flags=re.DOTALL)))
            sup  = _mt(''.join(re.findall(r'<m:sup\b[^>]*>(.*?)</m:sup>', blk, flags=re.DOTALL)))
            return f'<m:t>{base}^{sup}</m:t>'

        # 先替换复合结构，再替换单一结构；用 sub(..., count=0) 以处理一个块内多次出现
        math_xml = re.sub(r'<m:sSubSup\b[^>]*>.*?</m:sSubSup>', _repl_subsup, math_xml, flags=re.DOTALL)
        math_xml = re.sub(r'<m:sSub\b[^>]*>.*?</m:sSub>', _repl_sub, math_xml, flags=re.DOTALL)
        math_xml = re.sub(r'<m:sSup\b[^>]*>.*?</m:sSup>', _repl_sup, math_xml, flags=re.DOTALL)

        # 统一收集该公式块的所有 m:t（包括我们刚替换注入的）
        return _mt(math_xml)

    out, last = [], 0
    math_pat = re.compile(r'<m:(oMath|oMathPara)\b[^>]*>.*?</m:\1>', re.DOTALL)

    for m in math_pat.finditer(xml_p):
        out.append(_wt(xml_p[last:m.start()]))   # 公式前正文
        out.append(_parse_math(m.group(0)))      # 公式块文本（保序、含后续内容）
        last = m.end()

    out.append(_wt(xml_p[last:]))                # 尾部正文
    return ''.join(out)

def split_figure_caption(s: str):
    s = s.strip()
    # 只按第一个空格分割
    if ' ' not in s:
        return s, ''   # 没空格则 index = 整串，caption 空
    idx, cap = s.split(' ', 1)
    return idx.strip(), cap.strip()

def collect_data_from_node(tree_node):
    data_list = []
    # 收集当前节点的数据
    for data in tree_node['data']:
        data_list.append(data)
    # 递归收集子节点的数据
    for child in tree_node['children']:
        data_list.extend(collect_data_from_node(child))
    return data_list

def find_node_by_text(node, txt):
    for child in node.get('children', []):
        # 先检查当前 child
        if txt in child.get('text', ''):
            return child
        # 再递归检查它的子节点
        result = find_node_by_text(child, txt)
        if result:
            return result
    return None

def get_datalist_by_text(image_tree, txt):
    # 收集txt下的所有data
    node = find_node_by_text(image_tree, txt)
    data_list = collect_data_from_node(node)
    return data_list
    
def get_outline_level_of_style(style):
    try:
        if style.element.pPr.outlineLvl is not None:
            return style.element.pPr.outlineLvl.val + 1
        else:
            return get_outline_level_of_style(style.base_style)
    except:
        return 10

def build_outline_tree(items):
    """
    items: list of dict
        [
            {"text": "...", "level": 0, "para_id": 5},
            {"text": "...", "level": 1, "para_id": 6},
            {"text": "...", "level": 1, "para_id": 9},
            ...
        ]

    返回：层级结构 list
    """

    if not items:
        return []

    root = []
    stack = []   # 每个元素为 node，对应其 level

    for it in items:
        node = {
            "text": it["text"],
            "para_id": it["para_id"],
            "level": it["level"],
            "children": []
        }

        # 如果是最高级（level=0），直接进入根
        if not stack or it["level"] == 0:
            root.append(node)
            stack = [node]      # 重置栈
            continue

        # 找到最近的上级标题
        while stack and stack[-1]["level"] >= it["level"]:
            stack.pop()

        if stack:
            # 挂到上级
            stack[-1]["children"].append(node)
        else:
            # 找不到更高层，直接当根节点
            root.append(node)

        stack.append(node)

    return root

def blob_to_cvimg(blob_bytes):
    arr = np.frombuffer(blob_bytes, dtype=np.uint8)
    img = cv2.imdecode(arr, cv2.IMREAD_COLOR)  # 或 IMREAD_UNCHANGED
    return img

def print_tree(node, indent=0):
    """
    以树结构打印你的章节树
    node: {'text','para_id','outline_level','children','data'}
    """
    prefix = " " * (indent*4)  # 每层缩进4空格
    text = (node.get('text') or "").strip() or "<ROOT>"
    para = node.get('para_id', -1)

    # 图片数量，用占位符表示
    img_count = len(node.get('data', []))
    img_info = f" 📷x{img_count}" if img_count > 0 else ""

    print(f"{prefix}- [{para}] {text}{img_info}")

    for child in node.get("children", []):
        print_tree(child, indent+1)


if __name__ == "__main__":
    tree = Create_image_tree(r"D:\Study\本子\文档识别生成\para_new\载人空间站用半导体分立器件CYSR3015C型硅肖特基二极管应用指南_zyh最终修订版.docx")
    datalist = get_datalist_by_text(tree, "器件概况")
    print_tree(tree)
    print("ok")
    
    
