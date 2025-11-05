from typing import Any
from gaphor.ui import utils
from gaphor.ui.utils import state
import gaphor.ui.utils.demo as demo

#图表结果分离
def result_spliter(f):
    def result_spliter_helper(n):
        results = f(n)
        table_list = []
        image_list = []
        for result in results:
            tag = get_tag(result)
            if tag == 'table':
                table_list.append(result)
            else:
                image_list.append(result)
        return (image_list, table_list)
    return result_spliter_helper
            

def get_tag(answer) -> str:
    """
    返回搜索结果的类型,应为'table'或'image'
    """
    try: # 这里做try包裹防止出现数据类型错误，后续不再包裹
        if answer[0] == 'table' or answer[0] == 'image':
            return answer[0]
        else:
            raise TypeError("未知的搜索类型")
    except TypeError as e:
        print(e)
    

def get_data(answer):
    """
    返回搜索结果的数据内容, 若为表格则是table_tree.wcTable，若为图片则是image的blob数据
    """
    return answer[1]


def get_caption(answer):
    """
    返回搜索结果的题注
    """
    return answer[2]

# 如果只返回一个合并列表则需要此装饰器@result_spliter
def get_research(text:str) -> list:
    """"
    通过用户的输入返回搜索结果的列表, research = list[answer], answer = (tag, data, caption)
    """
    return demo.get_research(text)
