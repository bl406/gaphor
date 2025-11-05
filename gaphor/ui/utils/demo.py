import json
import re
import sys
from typing import List, Tuple, Any, Dict, Iterable, Optional
from gaphor.ui.utils import state
from gaphor.ui.utils import img_show_ayf as img_show
from gaphor.ui.utils import table_tree_ayf as table_tree
import unicodedata
import requests

# =========================
# 配置区
# =========================
MODEL_PATH = "/hdd1/xz/models/DeepSeek-R1-Distill-Qwen-14B/"
DOCX_PATH = "/hdd2/ayf/ayf_code/acdm/utils/载人空间站用半导体分立器件CYSR3015C型硅肖特基二极管应用指南_zyh最终修订版.docx"
DEVICE_ID = 1  # 使用 CUDA:1
MAX_NEW_TOKENS = 1024
TOP_P = 0.9
TEMPERATURE = 0.4

API_BASE = "http://172.17.140.239:8000" 

def call_api_extract_keywords(user_query: str):
    url = f"{API_BASE}/extract_keywords"
    try:
        resp = requests.post(url, json={"query": user_query}, timeout=60)
        resp.raise_for_status()
        obj = resp.json()
        # 返回格式：{"keywords":[...], "sections":[...], "synonyms":[...]}
        return obj
    except Exception as e:
        print(f"⚠️ 调用远程关键词服务失败：{e}")
        # 回退：至少把原句塞到 keywords
        return {"keywords": [user_query.strip()], "sections": [], "synonyms": []}

def init_trees(docx_path: str):
    """
    只在 state 上没有时构建一次 image_tree / table_tree。
    同时做结构校验，确保是 dict 且含 'children'/'data' 键。
    """
    def _is_valid_tree(t):
        return isinstance(t, dict) and ('children' in t) and ('data' in t)
    
    if getattr(state, "image_tree", None) is None:
        print(f"--- 构建 image_tree: {docx_path} ---")
        try:
            state.image_tree = img_show.Create_image_tree(docx_path)
            if not _is_valid_tree(state.image_tree):
                raise TypeError("image_tree 结构异常（应为 dict，含 'children'/'data'）")
            print("--- image_tree 构建完成 ---")
        except Exception as e:
            print(f"⚠️ image_tree 构建失败：{e}")
            state.image_tree = {"text":"", "children":[], "data":[]}  # 兜底为空树
    
    if getattr(state, "table_tree", None) is None:
        print(f"--- 构建 table_tree: {docx_path} ---")
        try:
            state.table_tree = table_tree.Create_table_tree(docx_path)
            if not _is_valid_tree(state.table_tree):
                raise TypeError("table_tree 结构异常（应为 dict，含 'children'/'data'）")
            print("--- table_tree 构建完成 ---")
        except Exception as e:
            print(f"⚠️ table_tree 构建失败：{e}")
            state.table_tree = {"text":"", "children":[], "data":[]}  # 兜底为空树

def unique_keep_order(items: Iterable[Any], key=lambda x: x) -> List[Any]:
    seen = set()
    out = []
    for it in items:
        k = key(it)
        if k in seen:
            continue
        seen.add(k)
        out.append(it)
    return out

def fetch_from_trees(q: str):
    """
    用单个子串 q 分别在 image_tree / table_tree 下检索（安全版）。
    """
    # 保证已初始化（如果你在 run_once 里已经 init 过，这里就不会重复构建）
    if getattr(state, "image_tree", None) is None or getattr(state, "table_tree", None) is None:
        init_trees(DOCX_PATH)
    img = []
    tbl = []
    try:
        img = img_show.get_datalist_by_text(state.image_tree, q) or []
    except Exception as e:
        print(f"⚠️ 图片检索失败：{e}")
        img = []
    try:
        tbl = table_tree.get_datalist_by_text(state.table_tree, q) or []
    except Exception as e:
        print(f"⚠️ 表格检索失败：{e}")
        tbl = []
    return img, tbl

def as_string_result(item: Any, kind_hint: Optional[str] = None) -> str:
    """
    将任意一种结果结构转为字符串。
    兼容：
    1) ("table"/"image", data, caption, idx?)
    2) ("table"/"image", data, caption)
    3) 纯字符串
    """
    if isinstance(item, str):
        return item
    if isinstance(item, (list, tuple)) and item:
        tag = str(item[0]) if len(item) >= 1 else (kind_hint or "unknown")
        # caption 优先
        caption = None
        idx = None
        if len(item) >= 3:
            caption = item[2]
        if len(item) >= 4:
            idx = item[3]
        # 针对表格：若有 idx / caption，组合成“表{idx}: {caption}”
        if tag == "table":
            if caption and idx is not None:
                return f"表{idx}: {caption}"
            elif caption:
                return f"表: {caption}"
            else:
                return "表: <无题注>"
        # 针对图片
        if tag == "image":
            if caption and idx is not None:
                return f"图{idx}: {caption}"
            elif caption:
                return f"图: {caption}"
            else:
                return "图: <无题注>"
        # 兜底
        return str(item)
    # 再兜底：未知类型
    return str(item)

##AYF ADD
def _norm_text(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKC", s).lower()
    s = re.sub(r"[，。；：！？”“、（）【】《》]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _raw_key(item):
    # 支持 ('image'/'table', data, caption, idx?) 或 ('image'/'table', data, caption)
    if isinstance(item, (list, tuple)) and item:
        tag = item[0]
        cap = item[2] if len(item) >= 3 else ""
        idx = item[3] if len(item) >= 4 else None
        return (tag, idx, _norm_text(cap))
    return ("obj", id(item))

def _dedup_merge_raw(items):
    best = {}
    for it in items:
        k = _raw_key(it)
        # 更长 caption 的优先
        def _caplen(x):
            if isinstance(x, (list, tuple)) and len(x) >= 3:
                return len(str(x[2]) or "")
            return 0
        if (k not in best) or (_caplen(it) > _caplen(best[k])):
            best[k] = it
    return list(best.values())


def search_by_queries_raw(queries):
    """
    返回原始结果：(raw_img_list, raw_tbl_list)
    元素形如：
      图片: ('image', blob, caption, idx)
      表格: ('table', wcTable, caption, idx)
    """
    all_img_raw, all_tbl_raw = [], []
    for q in queries:
        img_res, tbl_res = fetch_from_trees(q)
        all_img_raw.extend(img_res)
        all_tbl_raw.extend(tbl_res)
    # 去重合并（不丢 idx）
    all_img_raw = _dedup_merge_raw(all_img_raw)
    all_tbl_raw = _dedup_merge_raw(all_tbl_raw)
    return all_img_raw, all_tbl_raw
#AYF ADD    

def search_by_queries(queries: List[str]) -> Tuple[List[str], List[str]]:
    """
    针对多个 query 子串进行检索并合并去重。
    返回 (image_strings, table_strings)
    """
    all_img, all_tbl = [], []
    for q in queries:
        img_res, tbl_res = fetch_from_trees(q)
        # 规范成字符串
        all_img.extend(as_string_result(x, "image") for x in img_res)
        all_tbl.extend(as_string_result(x, "table") for x in tbl_res)
    # 去重（按字符串本身去重）
    img_str = unique_keep_order(all_img, key=lambda s: s)
    tbl_str = unique_keep_order(all_tbl, key=lambda s: s)
    return img_str, tbl_str

# =========================
# 主流程
# =========================
def run_once(user_query: str):
    init_trees(DOCX_PATH)

    print("\n=== 生成检索关键词(JSON) - via API ===")
    plan = call_api_extract_keywords(user_query)
    print(json.dumps(plan, ensure_ascii=False, indent=2))
    
    # 聚合候选子串
    queries: List[str] = []
    # 优先 keywords，其次 sections，最后 synonyms
    queries.extend(plan.get("keywords", []))
    queries.extend(plan.get("sections", []))
    queries.extend(plan.get("synonyms", []))
    # 去重并过滤太短的词（避免大量误匹配）
    queries = [q for q in unique_keep_order(queries, key=lambda s: s) if len(q.strip()) >= 2]
    
    print("\n=== 开始检索 ===")
    print(queries)
    img_strings, tbl_strings = search_by_queries(queries)
    
    # =========================
    # 最终输出（仅字符串）
    # =========================
    print("\n========== 检索结果 ==========")
    if img_strings:
        print("\n【图片匹配】")
        for i, s in enumerate(img_strings, 1):
            print(f"{i}. {s}")
    else:
        print("\n【图片匹配】无")
    
    if tbl_strings:
        print("\n【表格匹配】")
        for i, s in enumerate(tbl_strings, 1):
            print(f"{i}. {s}")
    else:
        print("\n【表格匹配】无")

# AYF ADD
def _build_queries_from_api_or_fallback(text: str) -> list[str]:
    plan = call_api_extract_keywords(text) 
    
    queries: List[str] = []
    queries.extend(plan.get("keywords", []))
    queries.extend(plan.get("sections", []))
    queries.extend(plan.get("synonyms", []))
    # 去重 + 过滤太短
    queries = [q for q in unique_keep_order(queries, key=lambda s: s) if len(q.strip()) >= 2]
    return queries

def get_research(text: str):
    """
    返回 (image_list, table_list)
    其中：
      image_list: List[('image', blob, caption)]
      table_list: List[('table', wcTable, caption)]
    满足你下游的拆分/取字段逻辑。
    """

    queries = _build_queries_from_api_or_fallback(text)
    raw_img, raw_tbl = search_by_queries_raw(queries)

    # 转成三元组（去掉 idx）
    def to_triplets(items):
        trip = []
        for it in items:
            if isinstance(it, (list, tuple)) and len(it) >= 3:
                tag, data, caption, id = it[0], it[1], it[2], 0
                trip.append((tag, data, caption, 0))
        return trip

    img_trip = to_triplets(raw_img)
    tbl_trip = to_triplets(raw_tbl)
    return (img_trip, tbl_trip)

def main():
    if len(sys.argv) > 1:
        query = " ".join(sys.argv[1:])
    else:
        # 默认示例，可自行修改
        query = "请给我电特性相关的图表"
    print(f"用户输入：{query}")
    run_once(query)

    imgs, tbls = get_research(query)
    print(f"\n[debug] get_research => images: {len(imgs)}, tables: {len(tbls)}")

if __name__ == "__main__":
    main()