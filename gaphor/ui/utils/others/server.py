# server.py
import re
import json
from typing import Dict, List, Optional

import torch
from fastapi import FastAPI
from pydantic import BaseModel
from transformers import AutoTokenizer, AutoModelForCausalLM

# ====== 配置 ======
MODEL_PATH = "/hdd1/xz/models/DeepSeek-R1-Distill-Qwen-14B/"
DEVICE_ID = 1                    # 使用 CUDA:1
MAX_NEW_TOKENS = 1024
TOP_P = 0.9
TEMPERATURE = 0.4

# ====== Prompt（单行版）======
PROMPT_TMPL = (
    "你是一个检索关键词抽取器。给定用户的自然语言查询，你需要提取出其中的**关键短语**、**可能的文档小节名**、以及**常见的中英文同义词或缩写**。"
    "只输出 JSON，格式如下：{{\"keywords\": [\"k1\",\"k2\", ...], \"sections\": [\"s1\",\"s2\", ...], \"synonyms\": [\"a\",\"b\", ...]}}。"
    "要求：1. 仅输出一个合法 JSON，不要任何解释或多余文字。2. 所有 keywords 必须直接包含在用户原句中，不能改写或省略原有关键短语。"
    "3. 若用户输入中有多个并列短语（例如“电特性测试、反向漏电流”），每个都要单独作为一个 keyword。4. 可补充常见 sections 与 synonyms（可含中英混合）。"
    "5. 不要添加注释或额外描述。用户输入：{query}"
)

PROMPT_ZYH = (
    "你是一个检索关键词抽取器，用户给定自然语言查询，你需要提取出其中的**关键词汇**。"
    "只输出 JSON，格式如下：{{\"keywords\": [\"k1\",\"k2\", ...]}}。"
    "<要求>1. 仅输出一个合法 JSON，不要任何解释或多余文字。2. 所有 keywords 必须直接包含在用户输入中，不能改写或省略原有关键短语。"
    "3. 若用户输入中有多个并列短语，每个都要单独作为一个 keyword。"
    "5. 不要添加注释或额外描述。"
    "<用户输入>{query}</用户输入>"
)

PROMPT_GPT = """你是“术语分解器”。任务：把中文用户输入分解为最基础的词语/概念，去掉虚词与礼貌用语，保留名词术语；对常见复合词要保留整体词，同时拆出构成的基础词。
规则：
1) 去掉：请、帮我、一下、相关、的、一下、一下子、给我、能否、可以等功能词。
2) 保留领域名词与概念（如：元器件、电特性、反向漏电流、图表、参数、曲线、数据、测试）。
3) 复合词既保留整体，也拆基础成分：
   - 图表 → 图表，图，表
   - 曲线图 → 曲线图，曲线，图
   - 参数表 → 参数表，参数，表
4) 去重，按重要性从大到小排列（领域核心词在前）。
5) 仅以 JSON 数组输出，不要其他文字例如:{{\"keywords\": [\"k1\",\"k2\", ...]}}
<用户输入>{query}</用户输入>
"""

# ====== FastAPI App ======
app = FastAPI(title="Acsemai API")

class ExtractReq(BaseModel):
    query: str

class TableChatReq(BaseModel):
    prompt: str

class ExtractResp(BaseModel):
    keywords: List[str]
    sections: List[str]
    synonyms: List[str]

tokenizer: Optional[AutoTokenizer] = None
model: Optional[AutoModelForCausalLM] = None

@app.on_event("startup")
def load_model_once():
    global tokenizer, model
    print(f"--- Loading model {MODEL_PATH} to CUDA:{DEVICE_ID} ---")
    tokenizer = AutoTokenizer.from_pretrained(MODEL_PATH, trust_remote_code=True)
    model = AutoModelForCausalLM.from_pretrained(
        MODEL_PATH,
        torch_dtype=torch.float16,
        device_map={"": DEVICE_ID},
        trust_remote_code=True,
    )
    model.eval()
    print("--- Model ready ---")
    
@app.post("/table_chat")
def table_chat_sever(req: TableChatReq):
    assert model is not None and tokenizer is not None, "model not ready"
    prompt = req.prompt
    inputs = tokenizer(prompt, return_tensors="pt").to(f"cuda:{DEVICE_ID}")
    with torch.no_grad():
        out = model.generate(
            **inputs,
            max_new_tokens=MAX_NEW_TOKENS,
            do_sample=True,
            top_p=TOP_P,
            temperature=TEMPERATURE,
        )
    text = tokenizer.decode(out[0], skip_special_tokens=True).strip()
    return text

def _extract_json(text: str) -> Dict[str, List[str]]:
    # 截取第一个 {...}
    m = "{" + re.findall(r"\{(.*?)\}", text, flags=re.DOTALL)[1].strip().replace("\n", "") + "}"
    if not m:
        raise ValueError("no JSON found")
    obj = json.loads(m)
    # 规整字段
    for k in ("keywords", "sections", "synonyms"):
        v = obj.get(k, [])
        if not isinstance(v, list): v = []
        obj[k] = [str(x).strip() for x in v if str(x).strip()]
    return obj

@app.post("/extract_keywords", response_model=ExtractResp)
def extract_keywords(req: ExtractReq):
    assert model is not None and tokenizer is not None, "model not ready"
    prompt = PROMPT_TMPL.format(query=req.query.strip())
    inputs = tokenizer(prompt, return_tensors="pt").to(f"cuda:{DEVICE_ID}")
    with torch.no_grad():
        out = model.generate(
            **inputs,
            max_new_tokens=MAX_NEW_TOKENS,
            do_sample=True,
            top_p=TOP_P,
            temperature=TEMPERATURE,
        )
    text = tokenizer.decode(out[0], skip_special_tokens=True).strip()
    try:
        print(text)
        obj = _extract_json(text)
    except Exception:
        # 回退：至少把原句塞到 keywords
        obj = {"keywords": [req.query.strip()], "sections": [], "synonyms": []}
    # 强制去重 & 过滤极短词
    def uniq(xs): 
        seen=set(); r=[]
        for x in xs:
            if len(x) >= 2 and x not in seen:
                seen.add(x); r.append(x)
        return r
    return ExtractResp(
        keywords=uniq(obj.get("keywords", [])),
        sections=uniq(obj.get("sections", [])),
        synonyms=uniq(obj.get("synonyms", [])),
    )
