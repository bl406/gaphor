# table_tree.py
from docx import Document
from docx.document import Document as _Document
from docx.text.paragraph import Paragraph
from docx.table import Table as _Table
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.oxml.ns import qn
import re

#=====Table类型定义=====#
class wcTable:
    def __init__(self, tbl, table_index, caption=""):
        self._tbl = tbl
        self._rows = len(tbl.rows)
        self._cols = len(tbl.columns)
        self.table_index = table_index
        self.table_content = self._build_table()
        self.table_caption = caption
        
    def _build_table(self):
        table = []
        for r_index in range(len(self._tbl.rows)):
            row = []
            for c_index in range(len(self._tbl.row_cells(r_index))):
                cell = self._tbl.row_cells(r_index)[c_index]
                if c_index and cell._tc is self._tbl.row_cells(r_index)[c_index - 1]._tc:
                    continue
                row.append(wcCell(cell, grid_span=cell.grid_span, cell_index=(self.table_index, r_index, c_index)))
            table.append(row)
        return table
    
    def to_csv(self):
        csv = ""
        for row in self.table_content:
            for cell in row:
                csv += ("\"" + cell.to_csv_cell() + "\",") * cell.grid_span
            csv = csv[:-1] + "\n"
        return csv
    
    def set_cell_text(self, r, c, text):
        self.table_content[r][c].text = text
        
    def __getitem__(self, index):
        return self.table_content[index]
    
class wcCell:
    def __init__(self, cell, grid_span=1, cell_index=None):
        self.cell = cell
        self.grid_span = grid_span
        self._index = cell_index
        self.title = f"t{self._index[0]}_r{self._index[1]}_c{self._index[2]}"
        self.text = self._get_cell_text()
        
    def _get_cell_text(self):
        return self.cell.text
    
    def to_csv_cell(self):
        return self.text.replace("\"", "\"\"").replace("\n", "")
    
   
# === 与 img_show.py 对齐的工具函数（同名/同义） ===
def get_outline_level_of_style(style):
    """有效大纲等级：自身->继承->无则10（正文）"""
    try:
        if style.element.pPr.outlineLvl is not None:
            return style.element.pPr.outlineLvl.val + 1
        else:
            return get_outline_level_of_style(style.base_style)
    except Exception:
        return 10   # 正文
# （逻辑与 img_show.py 中一致）:contentReference[oaicite:3]{index=3}

def find_node_place(node, parent):
    """把 node 按层级挂到 parent 子树（与 img_show.py 一致）"""
    if parent['outline_level'] == node['outline_level'] - 1:
        parent['children'].append(node)
        node['parent'] = parent
    elif parent['outline_level'] > node['outline_level'] - 1:
        # 回溯到更高层
        return find_node_place(node, parent['parent'])
    else:
        raise ValueError("文档结构错误，请检查标题层级")
    return
# :contentReference[oaicite:4]{index=4}

def collect_data_from_node(tree_node):
    data_list = list(tree_node.get('data', []))
    for ch in tree_node.get('children', []):
        data_list.extend(collect_data_from_node(ch))
    return data_list
# :contentReference[oaicite:5]{index=5}

def find_node_by_text(node, txt_substr):
    """按子串匹配标题文本，深度优先"""
    for ch in node.get('children', []):
        if txt_substr in (ch.get('text') or ''):
            return ch
        found = find_node_by_text(ch, txt_substr)
        if found:
            return found
    return None
# :contentReference[oaicite:6]{index=6}


# === 表格 caption 规则（与图类似，只是把 Fig/图 换成 Table/表） ===
RE_TABCAP_APPX = re.compile(r"^\s*(表|Tab\.?|Table)\s*[A-ZＡ-Ｚ]\s*[\.\．]\s*\d+(?:[-．\.]\d+)*", re.I)
RE_TABCAP_NUM  = re.compile(r"^\s*(表|Tab\.?|Table)\s*[：:\.]?\s*\d+(?:[-．\.]\d+)*", re.I)

def _para_text_with_omml_inline(p: Paragraph) -> str:
    """
    与 img_show.py 中 extract_para_text_all 同意图：把段内的普通文本 + 公式文本并回收为一行字符串。
    这里做简化：仅取 w:t（若你需要完整 OMML 提取，可直接搬 img_show.py 的 extract_para_text_all）。:contentReference[oaicite:7]{index=7}
    """
    xml = p._p.xml
    return ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', xml, flags=re.DOTALL))

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

def _iter_block_items(doc: _Document):
    """
    逐块遍历文档 body：段落或表格，保持文档顺序。
    """
    body = doc.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield ('p', Paragraph(child, doc))
        elif isinstance(child, CT_Tbl):
            yield ('tbl', _Table(child, doc))

def _next_nonempty_para_text(doc: _Document, start_after: Paragraph):
    """
    从 start_after 之后向下找第一段非空文本（作为候选 caption）。
    """
    body = doc.element.body
    passed = False
    for child in body.iterchildren():
        if not passed:
            passed = (child is start_after._p or child is start_after._p.getparent())
            continue
        if isinstance(child, CT_P):
            p = Paragraph(child, doc)
            txt = (_para_text_with_omml_inline(p) or "").strip()
            if txt:
                return txt
    return ""

def _prev_para_text(doc: _Document, anchor_child):
    """
    anchor_child 是 CT_Tbl（表格）或 CT_P（段落），
    从其前面开始向上找第一段非空文本（用作表注）。
    """
    body = doc.element.body
    prev = None
    children = list(body.iterchildren())
    for idx, child in enumerate(children):
        if child is anchor_child:
            # 向上回溯直到遇到非空段落
            for j in range(idx - 1, -1, -1):
                c2 = children[j]
                if isinstance(c2, CT_P):
                    p = Paragraph(c2, doc)
                    txt = (extract_para_text_all(p._p.xml) or "").strip()
                    if txt:
                        return txt
            return ""
    return ""


# === 主函数：构建“标题树 + 表格数据” ===
def Create_table_tree(doc_path: str):
    doc = Document(doc_path)

    root = {
        'text': '',
        'para_id': 0,
        'outline_level': 0,
        'parent': None,
        'children': [],
        'data': []
    }
    parent = root

    # 我们需要追踪“到目前为止见过的第几个表”，以便 wcTable(table_index=…)
    tbl_count = 0
    para_idx = -1  # 仅用于在遇到段落时自增索引（和 img_show 的行为对齐）:contentReference[oaicite:8]{index=8}

    for kind, obj in _iter_block_items(doc):
        if kind == 'p':
            para_idx += 1
            p: Paragraph = obj
            lvl = get_outline_level_of_style(p.style)
            if lvl != 10:
                node = {
                    'text': p.text,
                    'para_id': para_idx,
                    'outline_level': lvl,
                    'parent': None,
                    'children': [],
                    'data': []
                }
                find_node_place(node, parent)
                parent = node
            # 非标题段落：忽略（表格由 'tbl' 分支处理）
        elif kind == 'tbl':
            tbl = obj  # python-docx Table
            wct = wcTable(tbl, tbl_count)

            # Caption 策略：优先取“上一段非空文本”；若匹配“表/Tab/Table …”更好
            caption = _prev_para_text(doc, tbl._tbl) or ""
            if caption:
                if RE_TABCAP_APPX.match(caption) or RE_TABCAP_NUM.match(caption):
                    pass  # 合格的表注
                else:
                    caption = "未知表格"
                    pass

            # 把这个表挂到当前 parent 节点
            parent['data'].append(("table", wct, caption, tbl_count))
            tbl_count += 1

    return root


# === API：按标题文本收集该子树所有 data（与你的 img_show 同名接口） ===
def get_datalist_by_text(table_tree, txt):
    # 收集txt下的所有data
    if not txt:
        return collect_data_from_node(table_tree)
    
    node = find_node_by_text(table_tree, txt)
    if not node:
        return []
    return collect_data_from_node(node)
# （find_node_by_text / collect_data_from_node 与 img_show.py 逻辑一致）:contentReference[oaicite:10]{index=10}


# === 简单测试 ===
if __name__ == "__main__":
    path = r"E:\Gaphor\gaphor\gaphor\ui\utils\others\载人空间站用半导体分立器件CYSR3015C型硅肖特基二极管应用指南_zyh最终修订版.docx"
    tree = Create_table_tree(path)
    dl = get_datalist_by_text(tree, "")
    # dl 形如：[("table", wcTable实例, caption, table_index), ...]
    print("器件概况 下表格数：", len([d for d in dl if d and d[0] == "table"]))
