from typing import Any
from gaphor.ui import utils
from gaphor.ui.utils import state, img_show, table_tree
import gaphor.ui.utils.demo as demo
from importlib.resources import files
import json, re
import requests
from gaphor.diagram.styleeditor import AcsemAITableChatWindow

### 测试
from gaphor.ui.utils.table_tree import wcTable
from docx import Document

ImageChatPrompt_TEMPL = """
你是一个可靠的多模态助理。请严格遵循以下规则完成任务：

【已知输入】
- 用户输入（User Input）：{0}
- 图片题注（Image Caption）：{1}
- 提示：你将同时收到图像内容（已在本次对话的同一条消息中一并提供）。

【任务要求】
1) 先理解图片与用户输入之间的关系，题注仅作为辅助；当题注为“未知图片”时，忽略题注。    
2) 最终答案必须使用<answer></answer>标签包裹

【现在开始】请基于图像与输入完成任务，并按格式输出。
"""

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

def image_chat(user_input:str, image_pack):
    image_blob, caption = image_pack
    prompt = ImageChatPrompt_TEMPL.format(user_input, caption)
    url = get_api("img_chat")
    try:
        resp = requests.post(url, json={"prompt": prompt, "image_blob": image_blob.decode("latin1")}, timeout=60)
        print(resp.json())
        answer = re.findall(r"<answer>(.*?)</answer>", resp.json(), flags=re.DOTALL)[0].strip()
        return answer
    except Exception as e:
        print(f"⚠️ 调用远程关键词服务失败：{e}")
        return f"⚠️ 调用远程关键词服务失败：{e}"
    
def all_chat(text, image_result, table_result):
    table_content = ""
    for idx, table_tup in enumerate(table_result):
        table_content += f"""
        题注:{table_tup[1].table_caption}
        表格内容:{table_tup[1].to_csv()}
        """
        
    image_content = ""
    image_data = []
    for idx, image_tup in enumerate(image_result):
        image_content += f"图片ID:{idx}, 图注:{image_tup[2]}\n"
        image_data.append(image_tup[1].decode("latin1"))
    
    query = text
        
    # ALL_PROMPT = f"""
    # <Table>{table_content}</Table>
    # <Image>{image_content}</Image>
    # <Task>上面是一些CSV的表格，和一些图片及其按顺序对应的图注，你需要阅读所有材料，完成任务，并将你的最终答案通过<answer></answer>包裹。任务是：{query}</Task>
    # """
    ALL_PROMPT = f"""
        你是一个严谨的视觉-表格联合理解助理。请严格遵守以下规则完成任务：

        【已提供材料】
        1）表格内容：
        <Table>
        {table_content}
        </Table>

        2）图片及对应图注：
        <Image>
        {image_content}
        </Image>

        【特别说明】
        - 表格与图注按顺序一一对应。
        - 若某条图注为“未知图片”，代表该图片未获取到题注，请不要杜撰或使用不存在的信息。
        - 所有推断必须基于表格、图片和题注本身，不允许编造内容。

        【你的任务】
        - 阅读所有表格与图片信息，理解其中的数据、关系和语义。
        - 完成以下任务：{query}

        【输出要求】
        1）只输出最终答案，不输出推理过程、不输出引言、不输出解释。
        2）最终答案必须使用以下格式包裹，不允许在标签外包含其他文字。
        3）如果图片中有数据，你必须结合其中具体数据来对图片内容进行分析，确保每一个结论都有具体数据支撑。
        <answer>你的最终答案</answer>

        现在开始，请阅读材料并按照要求回答。
        """

    url = get_api("all_chat")
    
    print(f"[AcsemAIChatALL RunWith]: \n{ALL_PROMPT}")
    
    try:
        resp = requests.post(url, json={"prompt": ALL_PROMPT, "image_blobs": image_data}, timeout=60)
        print(resp.json())
        answer = re.findall(r"<answer>(.*?)</answer>", resp.json(), flags=re.DOTALL)[0].strip()
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
    trt = table_tree.Create_table_tree(r"E:\Gaphor\gaphor\gaphor\ui\utils\others\载人空间站用半导体分立器件CYSR3015C型硅肖特基二极管应用指南_zyh最终修订版.docx")
    tree = img_show.Create_image_tree(r"E:\Gaphor\gaphor\gaphor\ui\utils\others\载人空间站用半导体分立器件CYSR3015C型硅肖特基二极管应用指南_zyh最终修订版.docx")
    datalist = img_show.get_datalist_by_text(tree, "温度", find_all=True)
    tablelist = table_tree.get_datalist_by_text(trt, "温度", find_all=True)
    all_chat("总结图表内容", datalist, tablelist)
    image_chat("请总结这个图片中的内容", (datalist[0][1], datalist[0][2]))
    tables = doc.images
    print("111")
    tbl = wcTable(tables[0], 0, "表1 最大额定值")
    table_chat("请总结这个表格的内容", tbl)

