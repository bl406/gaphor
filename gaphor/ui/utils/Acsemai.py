from typing import Any
from gaphor.ui import utils
from gaphor.ui.utils import state, img_show, table_tree
import gaphor.ui.utils.demo as demo
from importlib.resources import files
import json, re
import requests

### 测试
from gaphor.ui.utils.table_tree import wcTable
from docx import Document
def get_api(key):
    """
    key传入api_url_config的参数
    search : 搜索
    table_chat : 表格问答
    img_chat : 图片问答
    """
    api_url_file_config = files("gaphor.ui.utils").joinpath("api_url_config.json")
    with open(api_url_file_config,'r', encoding='UTF-8') as f:
        load_dict = json.load(f)
    base = load_dict["url"]
    return f"{base}/{load_dict[key]}"

# 对输入的表格进行问答
def table_chat(user_input:str, table):
    # make table chat prompt
    caption = table.table_caption
    csv = table.to_csv()
    prompt = f"""这是一段CSV格式的表格，你需要仔细阅读整个表格，完成任务，并将你的最终答案通过<answer></answer>包裹。注意，当题注为"未知表格"时，代表这个表格未成功获取题注。
    <Table>
    题注：{caption}
    表格内容：{csv}
    </Table>
    <Task>
    {user_input}
    </Task>
    """
    # connect to LLM
    url = get_api("table_chat")
    try:
        resp = requests.post(url, json={"prompt": prompt}, timeout=60)
        print(resp.json())
        answer = re.findall(r"<answer>(.*?)</answer>", resp.json(), flags=re.DOTALL)[1].strip()
        return answer
    except Exception as e:
        print(f"⚠️ 调用远程关键词服务失败：{e}")
        return f"⚠️ 调用远程关键词服务失败：{e}"
    
# 如果只返回一个合并列表则需要此装饰器@result_spliter
def get_research(text:str) -> list:
    """"
    通过用户的输入返回搜索结果的列表, research = list[answer], answer = (tag, data, caption)
    """
    return demo.get_research(text)

if __name__ == "__main__":
    doc_path = r"E:\Gaphor\gaphor\gaphor\ui\utils\others\载人空间站用半导体分立器件CYSR3015C型硅肖特基二极管应用指南_zyh最终修订版.docx"
    doc = Document(doc_path)
    tables = doc.tables
    print("111")
    tbl = wcTable(tables[0], 0, "表1 最大额定值")
    table_chat("请总结这个表格的内容", tbl)

