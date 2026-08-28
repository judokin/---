# 使用提醒:
# 1. xbot包提供软件自动化、数据表格、Excel、日志、AI等功能
# 2. package包提供访问当前应用数据的功能，如获取元素、访问全局变量、获取资源文件等功能
# 3. 当此模块作为流程独立运行时执行main函数
# 4. 可视化流程中可以通过"调用模块"的指令使用此模块

'''
20260518 修改
根据刚刚和刀哥的会议，本周主要工作内容如下：
华越：
完善RPA表格中三张表的字段：补充老FBA父体、固定字段，更新标题和五点文案等，并将字段标题与花型表、上架模版表格的标题对齐。


在花型表格中增加“父体”列，与RPA表格对齐（包括上架选择店铺、卖家SKU、选择运营等）。
AI广告相关事情。
朱哥：

更新父体映射关系（新POD尾缀父体 + 老FBA父体）。
将标题、五点中的 ##### 字段替换为实际内容。
在标题末尾添加颜色。
待图片新命名方式输出后，修改图片抓取顺序。
修改自动上架频率为每天两次：上午10:00、下午15:00各一次。
上架模版改为按父体维度输出。
测试：优先用自然仿生风格父体跑测试。
以上任务请各位本周内推进，有问题随时沟通[致敬]


'''
'''
20260629 修改
取SKC时要把 材质方向 也取下来

根据材质方向，分别对应不同的表格，取不同的固定
材质方向为：印花地毯 取 固定-地毯 链接为：https://wit0jhu6kvu.feishu.cn/base/Jolyb8QBoaPzj6swf0cc6bqenlf?table=tblTer6BHOZRAQkB&view=vewEPcSVia
材质方向为：仿羊绒厨房垫 取 固定-厨房垫 链接为：https://wit0jhu6kvu.feishu.cn/base/Jolyb8QBoaPzj6swf0cc6bqenlf?table=tbl18dyMkn1KS8KH&view=vew294Yp4d

仿羊绒厨房垫-Item Name 20 x32" +20 x48"
仿羊绒厨房垫-Bullet Point1
仿羊绒厨房垫-Bullet Point2
仿羊绒厨房垫-Bullet Point3
仿羊绒厨房垫-Bullet Point4
仿羊绒厨房垫-Bullet Point5
仿羊绒厨房垫-Generic Keyword
仿羊绒厨房垫-Style

材质方向为：三明治户外垫 取 固定-三明治 链接为：https://wit0jhu6kvu.feishu.cn/base/Jolyb8QBoaPzj6swf0cc6bqenlf?table=tblCffxAeXGnocAY&view=vewzrq6EcO
三明治户外垫-商品名称2X3
三明治户外垫-商品名称2.5X8
三明治户外垫-商品名称3X5
三明治户外垫-商品名称5X7
三明治户外垫-商品名称5X8
三明治户外垫-商品名称6X9
三明治户外垫-商品名称8X10
三明治户外垫-商品名称9X12
三明治户外垫-关于此艺术品
三明治户外垫-商品特性1
三明治户外垫-商品特性2
三明治户外垫-商品特性3
三明治户外垫-商品特性4
三明治户外垫-商品特性5
三明治户外垫-搜索关键词
三明治户外垫-设计


@朱哥（朱真金） 
刚刚沟通的在RPA表里已经改好啦
厨房垫sku命名规则：SKC+尺寸+POD。尺寸部分，乘号大写，末尾单位是IN，例如ARA019-RED BEIGE-20X32-20X48IN-POD。
户外地垫sku命名规则：SKC-尺寸-POD1。尺寸部分，规则和地毯一样，乘号大写，尺寸末尾加单位FT，例如AR6006-RED-2.5X8FT-POD1。


三明治材料这边，有一个上架问题。他们想要区分新老父体，新父体老父体内上传不同尺寸的测款。老父体只
上传5X7尺寸的sku，
新父体内上传2.5×8，
3x5，
5×7，
6×9，
8x10，
9x10，
9x12
这些尺寸的sku。
可以从新父体sku的命名上做区分吗？例如三明治的新父体sku命名统一开头是SMZ，结尾是-NEW。
老父体结尾没有-NEW

父体和子体已经加上了 材质方向 后续同步过去子体那边去

三明治新父体尾缀改一下吧，以SMZ开头，-POD结尾。和地毯的新父体统一（目前还没有）
'''


try:
    import xbot
    from xbot import print, sleep
    from .import package
    from .package import variables as glv
except:
    glv = {}
    glv['gvar_module1_done'] = False
    glv['gvar_shop_name'] = "亚马逊--冬豚--北美（子账号）"
    glv['gvar_shop_name'] = '亚马逊--岚风（子账号）'
    glv['gvar_shop_name'] = "亚马逊--灿东（子账号）"
    glv['gvar_shop_name'] = '亚马逊--冬豚--北美（子账号）'
    glv['gvar_shop_name'] = '亚马逊--岚风（子账号）'
    glv['gvar_shop_name'] = "亚马逊-北蓉-北美（子账号）"
    glv['gvar_shop_name_list'] = ['亚马逊--冬豚--北美（子账号）', '亚马逊--岚风（子账号）', '亚马逊--灿东（子账号）', '亚马逊--冬豚--北美（子账号）', '亚马逊-北蓉-北美（子账号）']
    glv['gvar_对应运营id'] = ""
    pass
import json
#import config
import requests
import os
import sys
import time
import copy
import re
import importlib.util
import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils import get_column_letter
from pathlib import Path

# Windows 本地控制台通常使用 GBK；文案中可能包含 U+2011 等无法编码字符。
# 日志显示时替换不可编码字符，避免 print 中断数据处理流程。
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(errors='replace')

OUT_FILE_RES = False
# True：测试模式，只生成本地 Excel，不新增或更新飞书子表。
# False：正式模式，生成 Excel，同时新增并更新飞书子表。
TEST_MODE = True
TEST_MODE = False
# 指定文件路径
config_file_path = Path(__file__).parent / "config.py"

# 加载模块
spec = importlib.util.spec_from_file_location("config", config_file_path)
config = importlib.util.module_from_spec(spec)
spec.loader.exec_module(config)

url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal/"
payload = json.dumps({
"app_id": config.app_id,
"app_secret": config.app_secret
})
headers = {
'Content-Type': 'application/json'
}
response = requests.request("POST", url, headers=headers, data=payload)
print(response.status_code)
print(response.text)
tenant_access_token = response.json()["tenant_access_token"]

def send(text):
    import requests
    import json
    webhook_url = config.webhook_url_v2

    # 发送的消息内容
    message = {
        "msg_type": "text",
        "content": {
            "text": text
        }
    }

    # 发送请求
    response = requests.post(webhook_url, headers={"Content-Type": "application/json"}, data=json.dumps(message))

    # 打印结果
    if response.status_code == 200:
        print("消息发送成功")
    else:
        print(f"消息发送失败: {response.text}")

def download_file_excel(file_name, temp_url):
    headers = {
        'Authorization': 'Bearer ' + tenant_access_token
    }
    response = requests.get(temp_url, headers=headers, stream=True)

    if response.status_code == 200:
        # 确保 Content-Type 是文件类型，而不是 JSON 错误信息
        # Excel 文件的 Content-Type 通常是 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        # 但对于 tmp_url 接口，可能只是 'application/octet-stream'
        
        # 成功的下载流程：
        with open(file_name, "wb") as file:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk: 
                    file.write(chunk)
        print(f"文件已成功保存为: {file_name}")
    else:
        # 如果不是 200，明确打印出 API 返回的错误内容，有助于排查
        try:
            error_msg = response.json()
        except requests.exceptions.JSONDecodeError:
            error_msg = response.text
            
        print(f"下载失败，状态码: {response.status_code}，响应: {error_msg}")
        
def download_file(file_name, url):
    headers = {
        'Authorization': 'Bearer ' + tenant_access_token
    }
    #url = "https://open.feishu.cn/open-apis/drive/v1/medias/batch_get_tmp_download_url?file_tokens=FM5ebnCXJo36FLx9ii4cMrsKn9c&extra=%7B%22bitablePerm%22%3A%7B%22tableId%22%3A%22tblux7wXHLPNgroJ%22%2C%22rev%22%3A3%7D%7D"
    #response = requests.request("GET", url, headers=headers, data=payload)
    response = requests.get(url, headers=headers, stream=True)

    if response.status_code == 200:
        # 将文件保存到本地 
        # file_name = "test"
        with open(file_name, "wb") as file:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:  # 过滤掉保持活动的空块
                    file.write(chunk)
        print(f"文件已成功保存为: {file_name}")
    else:
        print(f"下载失败，状态码: {response.status_code}，响应: {response.text}")

def reset_fields(data, record_id, str1, str2):
    # https://wit0jhu6kvu.feishu.cn/base/Jolyb8QBoaPzj6swf0cc6bqenlf?table=tbl8zJByKrMGUJvM&view=vew4LsxOAs
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/Jolyb8QBoaPzj6swf0cc6bqenlf/tables/tbl8zJByKrMGUJvM/records/{record_id}"
    
    # 定义请求头
    headers = {
        "Authorization": "Bearer " + tenant_access_token,
        "Content-Type": "application/json"
    }
    response = requests.put(url, headers=headers, json=data)

    print(response.status_code)
    print(response.json())
    import requests
import json

def reset_fields_batch_update_v2(data, str1, str2):
    # 飞书多维表格批量更新接口地址
    url = "https://open.feishu.cn/open-apis/bitable/v1/apps/Jolyb8QBoaPzj6swf0cc6bqenlf/tables/tbl8zJByKrMGUJvM/records/batch_update"
    
    # 定义请求头，需替换为实际有效的tenant_access_token
    headers = {
        "Authorization": "Bearer " + tenant_access_token,
        "Content-Type": "application/json"
    }
    
    # 发送POST请求，使用json参数自动序列化数据为JSON格式，无需手动调用json.dumps
    try:
        response = requests.post(url, headers=headers, json=data, timeout=10)
        # 如果响应状态码不在200-299区间，主动抛出HTTP异常
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"请求发生异常: {e}")
        return None
    
    # 打印请求结果
    print(f"响应状态码: {response.status_code}")
    try:
        result = response.json()
        print("响应JSON结果:")
        # 格式化输出JSON结果，便于阅读
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return result
    except json.JSONDecodeError:
        print("响应内容非JSON格式，原始内容:")
        print(response.text)
        return response.text
    

def reset_fields_batch_update(data, str1, str2):
    # https://open.feishu.cn/open-apis/bitable/v1/apps/:app_token/tables/:table_id/records/batch_update
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/Jolyb8QBoaPzj6swf0cc6bqenlf/tables/tbl8zJByKrMGUJvM/records/batch_update"
    
    # 定义请求头
    headers = {
        "Authorization": "Bearer " + tenant_access_token,
        "Content-Type": "application/json"
    }
    response = requests.post(url, headers=headers, json=data)

    print(response.status_code)
    print(response.json())
def get_all_table_data(
    str1='Jolyb8QBoaPzj6swf0cc6bqenlf',
    str2='tbl8zJByKrMGUJvM',
    material_direction=None,
):
    import datetime
    # https://wit0jhu6kvu.feishu.cn/base/Jolyb8QBoaPzj6swf0cc6bqenlf?table=tbl8zJByKrMGUJvM&view=vew4LsxOAs
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{str1}/tables/{str2}/records/search?page_size=500"

    # 定义请求头
    headers = {
        "Authorization": "Bearer " + tenant_access_token,
        "Content-Type": "application/json"
    }
    page_token = ""
    items_list = []
    for i in range(100):
        data = {}
        if material_direction:
            data = {
                "filter": {
                    "conjunction": "and",
                    "conditions": [{
                        "field_name": "材质方向",
                        "operator": "is",
                        "value": [material_direction],
                    }],
                }
            }
        request_url = url + "&page_token=" + page_token if page_token else url
        data_json = {}
        last_failure = "未知错误"
        for attempt in range(5):
            try:
                response = requests.post(
                    request_url, headers=headers, json=data, timeout=30
                )
            except requests.exceptions.RequestException as exc:
                last_failure = f"网络异常：{type(exc).__name__}: {exc}"
                if attempt < 4:
                    wait_seconds = 2 ** attempt
                    print(
                        f"飞书查询网络异常，第 {attempt + 1} 次失败，"
                        f"{wait_seconds} 秒后重试同一页：{exc}"
                    )
                    time.sleep(wait_seconds)
                continue

            print("Status Code1:", response.status_code)
            try:
                data_json = response.json()
            except ValueError:
                data_json = {}
            if (
                response.ok
                and isinstance(data_json.get('data'), dict)
                and isinstance(data_json['data'].get('items'), list)
            ):
                break
            last_failure = (
                f"HTTP {response.status_code}，"
                f"code={data_json.get('code')}，msg={data_json.get('msg')}"
            )
            if attempt < 4:
                wait_seconds = 2 ** attempt
                print(
                    f"飞书查询失败，第 {attempt + 1} 次重试："
                    f"{last_failure}，{wait_seconds} 秒后重试同一页"
                )
                time.sleep(wait_seconds)
        else:
            raise RuntimeError(
                f"飞书表格查询失败：app={str1}，table={str2}，"
                f"page_token={page_token or '<第一页>'}，{last_failure}"
            )
        #datas = []
        for i, items in enumerate(data_json['data']['items']):
            if len(items['fields']) == 0:
                continue
            # print(items)
            # print("----------------------------------- i =", i)
            items_list.append(items)
            #break
        if 'has_more' in data_json['data'] and data_json['data']['has_more']:
            page_token = data_json['data']['page_token']
        else:
            break
        # import pdb;pdb.set_trace()
    return items_list
FIVE_POINTS_APP_TOKEN = 'Jolyb8QBoaPzj6swf0cc6bqenlf'
FIVE_POINTS_TABLE_TOKEN = 'tblYg6tDUuMAKsfT'
FIVE_POINTS_TABLE_URL = (
    'https://wit0jhu6kvu.feishu.cn/base/'
    f'{FIVE_POINTS_APP_TOKEN}?table={FIVE_POINTS_TABLE_TOKEN}'
)


def post_feishu_search_with_retry(
    url,
    headers,
    data,
    context,
    status_label,
    max_retries=5,
):
    """请求飞书多维表格；网络抖动或临时服务异常时重试同一页。"""
    last_failure = "未知错误"
    for attempt in range(max_retries):
        try:
            response = requests.post(
                url,
                headers=headers,
                json=data,
                timeout=30,
            )
        except requests.exceptions.RequestException as exc:
            last_failure = f"网络异常：{type(exc).__name__}: {exc}"
        else:
            print(status_label, response.status_code)
            try:
                data_json = response.json()
            except ValueError:
                data_json = {}
                last_failure = (
                    f"HTTP {response.status_code}，返回非 JSON："
                    f"{response.text[:500]}"
                )
            else:
                if (
                    response.ok
                    and isinstance(data_json.get('data'), dict)
                    and isinstance(data_json['data'].get('items'), list)
                ):
                    return data_json
                last_failure = (
                    f"HTTP {response.status_code}，"
                    f"code={data_json.get('code')}，msg={data_json.get('msg')}"
                )

        if attempt < max_retries - 1:
            wait_seconds = 2 ** attempt
            print(
                f"{context}第 {attempt + 1} 次失败：{last_failure}；"
                f"{wait_seconds} 秒后重试同一请求"
            )
            time.sleep(wait_seconds)

    raise RuntimeError(
        f"{context}连续 {max_retries} 次失败：{last_failure}"
    )


def get_five_points_table_data():
    import datetime
    # https://wit0jhu6kvu.feishu.cn/base/CUgObJq4aas1busUW5HcU4scnSd?table=tblcscrczF5v3pCm&view=vewA4pwieU
    # https://wit0jhu6kvu.feishu.cn/base/U7dFbJst5aur9xs2pWac0k4unuc?table=tblNiu92eNYsktTe&view=vewmDyOW2t
    # 查询多维表格数据
    # https://wit0jhu6kvu.feishu.cn/base/U7dFbJst5aur9xs2pWac0k4unuc?table=tblcLQ2Gk5Jh5T0y&view=vewGFhSEMr
    # https://wit0jhu6kvu.feishu.cn/base/Jolyb8QBoaPzj6swf0cc6bqenlf?table=tblYg6tDUuMAKsfT&view=vewoqs4zTj
    url = (
        'https://open.feishu.cn/open-apis/bitable/v1/apps/'
        f'{FIVE_POINTS_APP_TOKEN}/tables/{FIVE_POINTS_TABLE_TOKEN}/'
        'records/search?page_size=500'
    )

    # 定义请求头
    headers = {
        "Authorization": "Bearer " + tenant_access_token,
        "Content-Type": "application/json"
    }
    page_token = ""
    items_list = []
    feishu_row_number = 0
    for i in range(100):
        if page_token != "":
            url += "&page_token=" + page_token
            pass
        else:
            pass
        data = {}
        data_json = post_feishu_search_with_retry(
            url,
            headers,
            data,
            context=(
                f"飞书五点文案表查询（page_token="
                f"{page_token or '<第一页>'}）"
            ),
            status_label="Status Code2:",
        )
        #datas = []
        for i, items in enumerate(data_json['data']['items']):
            feishu_row_number += 1
            if len(items['fields']) == 0:
                continue
            # 多维表接口不返回界面固定行号，保存当前接口返回顺序，
            # 并同时保留 record_id，便于报错时准确定位记录。
            items['_feishu_row_number'] = feishu_row_number
            # print(items)
            # print("----------------------------------- i =", i)
            items_list.append(items)
            #break
        if 'has_more' in data_json['data'] and data_json['data']['has_more']:
            page_token = data_json['data']['page_token']
        else:
            break
        # import pdb;pdb.set_trace()
    return items_list
import datetime
from typing import Optional

def timestamp_ms_to_datetime(
    timestamp_ms: int,
    tz: Optional[str] = "UTC",
    fmt: str = "%Y-%m-%d %H:%M:%S"
) -> str:
    """
    将毫秒时间戳转换为日期时间字符串
    
    :param timestamp_ms: 毫秒时间戳，例如 1766543156000
    :param tz: 时区，'UTC' 或 'local'，或 IANA 时区名如 'Asia/Shanghai'
    :param fmt: 输出格式，默认为 '%Y-%m-%d %H:%M:%S'
    :return: 格式化后的日期时间字符串
    """
    # 转换为秒
    timestamp_s = timestamp_ms / 1000.0
    
    if tz == "UTC":
        dt = datetime.datetime.utcfromtimestamp(timestamp_s)
    elif tz == "local":
        dt = datetime.datetime.fromtimestamp(timestamp_s)
    else:
        # 使用 zoneinfo（Python 3.9+）处理任意时区
        try:
            from zoneinfo import ZoneInfo
            dt = datetime.datetime.fromtimestamp(timestamp_s, tz=ZoneInfo(tz))
        except ImportError:
            raise RuntimeError("zoneinfo 需要 Python 3.9+，请使用 UTC 或 local")
    
    return dt.strftime(fmt)


def get_table_data(material_direction=None, skcs=None):
    '''
    父类子类基础数据
    先匹配子类，再匹配父类
    '''
    material_direction = str(material_direction or '').strip()
    skcs = list(dict.fromkeys(
        str(skc).strip() for skc in (skcs or []) if str(skc).strip()
    ))
    items_list = get_five_points_table_data()
    parent_items_dict = {}
    for item in items_list:
        if '卖家 SKU' not in item['fields']:
            continue
        parent_items_dict[item['fields']['卖家 SKU']] = item

    # import  pdb;pdb.set_trace()
    import datetime
    # https://wit0jhu6kvu.feishu.cn/base/CUgObJq4aas1busUW5HcU4scnSd?table=tblcscrczF5v3pCm&view=vewA4pwieU
    # https://wit0jhu6kvu.feishu.cn/base/U7dFbJst5aur9xs2pWac0k4unuc?table=tblNiu92eNYsktTe&view=vewmDyOW2t
    # 查询多维表格数据
    # page_size=500&page_token=cGFnZVRva2VuOjUwMA%3D%3D
    url = "https://open.feishu.cn/open-apis/bitable/v1/apps/U7dFbJst5aur9xs2pWac0k4unuc/tables/tblNiu92eNYsktTe/records/search?page_size=500"

    # 定义请求头
    headers = {
        "Authorization": "Bearer " + tenant_access_token,
        "Content-Type": "application/json"
    }
    skc_datas = {}
    # 飞书文本字段的 is 操作符一次只能精确匹配一个值；
    # SKCS.txt 中的每个 SKC 分别发起查询，避免多值条件返回 400。
    query_skcs = skcs or [None]
    for query_skc in query_skcs:
        conditions = [
            {
                "field_name": "父体SKU",
                "operator": "isNotEmpty",
                "value": []
            },
            {
                "field_name": "图所在NAS盘地址",
                "operator": "isNotEmpty",
                "value": []
            },
            {
                "field_name": "FBM上架店铺",
                "operator": "is",
                "value": [glv['gvar_shop_name']]
            }
        ]
        if query_skc:
            conditions.append({
                "field_name": "SKC",
                "operator": "is",
                "value": [query_skc],
            })
        if material_direction:
            conditions.append({
                "field_name": "材质方向",
                "operator": "is",
                "value": [material_direction]
            })
        data ={
            "filter": {
                "conjunction": "and",
                "conditions": conditions
            }
        }
        print(data)
        data_json = post_feishu_search_with_retry(
            url,
            headers,
            data,
            context=f"飞书花型表查询（SKC={query_skc or '<全部>'}）",
            status_label="Status Code3:",
        )
        #datas = []
        for i, items in enumerate(data_json['data']['items']):
            if len(items['fields']) == 0:
                continue
            parent_item_text_info = {}
            pskc = ''
            if '父体SKU' in items['fields'] and items['fields']['父体SKU'] in parent_items_dict:
                parent_item_text_info = parent_items_dict[items['fields']['父体SKU']]
                pskc = items['fields']['父体SKU']
            else:
                parent_item_text_info = list(parent_items_dict.values())[0]
                continue
            # import random
            # my_list = random.choice(items_list)
            five_point_texts = []
            # parent_items_dict
            rp = items['fields']['元素'][0]['text'] if '元素' in items['fields'] else ''
            try:
                glv['材质方向'] = items['fields']['材质方向'][0]
            except:
                print('continue002材质方向为空，跳过')
                continue
            if material_direction and glv['材质方向'] != material_direction:
                print(
                    f"材质方向不匹配，跳过：实际={glv['材质方向']}，"
                    f"筛选={material_direction}"
                )
                continue
            # if '印花地毯' != glv['材质方向']:
            #     import pdb;pdb.set_trace()
            #     pass
            pattern = ''
            product_description = ''
            if glv['材质方向'] in ['印花地毯']:
                for point_key in  ['商品特性1', '商品特性2', '商品特性3', '商品特性4', '商品特性5']:
                    point_key = glv['材质方向'] + '-' + point_key
                    #print(point_key, '文案:', my_list['fields'][point_key][0]['text'])
                    five_point_texts.append(parent_item_text_info['fields'][point_key][0]['text'].replace("#####", rp))
                product_description = get_feishu_field_text(
                    parent_item_text_info['fields'],
                    glv['材质方向'] + '-Product Description',
                    glv['材质方向'] + '-关于此艺术品'
                )
                搜索关键词 = get_feishu_field_text(
                    parent_item_text_info['fields'],
                    glv['材质方向'] + '-搜索关键词',
                    glv['材质方向'] + '-Generic Keyword'
                )
                设计 = get_feishu_field_text(
                    parent_item_text_info['fields'],
                    glv['材质方向'] + '-设计',
                    glv['材质方向'] + '-Style'
                )
                pattern = get_feishu_field_text(
                    parent_item_text_info['fields'],
                    glv['材质方向'] + '-Pattern',
                    glv['材质方向'] + '-Design'
                )
            if glv['材质方向'] in ['三明治户外垫']:
                for point_key in  ['商品特性1', '商品特性2', '商品特性3', '商品特性4', '商品特性5']:
                    point_key = glv['材质方向'] + '-' + point_key
                    #print(point_key, '文案:', my_list['fields'][point_key][0]['text'])
                    point_text = get_feishu_field_text(
                        parent_item_text_info['fields'],
                        point_key,
                    )
                    if not point_text:
                        feishu_row = parent_item_text_info.get(
                            '_feishu_row_number',
                            '<未知>',
                        )
                        record_id = parent_item_text_info.get(
                            'record_id',
                            '<未知>',
                        )
                        error_skc = get_feishu_field_text(
                            items.get('fields', {}),
                            'SKC',
                            '花型命名',
                        ) or '<空>'
                        error_message = (
                            f"飞书五点文案表第 {feishu_row} 行缺少字段：{point_key}；"
                            f"SKC：{error_skc}；父体 SKU：{pskc or '<空>'}；"
                            f"记录 ID：{record_id}；"
                            f"表 ID：{FIVE_POINTS_TABLE_TOKEN}；表链接："
                            f"{FIVE_POINTS_TABLE_URL}"
                        )
                        print(error_message)
                        raise KeyError(error_message)
                    five_point_texts.append(point_text.replace("#####", rp))
                product_description = get_feishu_field_text(
                    parent_item_text_info['fields'],
                    glv['材质方向'] + '-Product Description',
                    glv['材质方向'] + '-关于此艺术品'
                )
                搜索关键词 = get_feishu_field_text(
                    parent_item_text_info['fields'],
                    glv['材质方向'] + '-搜索关键词',
                    glv['材质方向'] + '-Generic Keyword'
                )
                设计 = get_feishu_field_text(
                    parent_item_text_info['fields'],
                    glv['材质方向'] + '-设计',
                    glv['材质方向'] + '-Style'
                )
                pattern = get_feishu_field_text(
                    parent_item_text_info['fields'],
                    glv['材质方向'] + '-Pattern',
                    glv['材质方向'] + '-Design'
                )
            if '仿羊绒厨房垫' == glv['材质方向']:
                '''
                仿羊绒厨房垫-Item Name 20 x32" +20 x48"
                仿羊绒厨房垫-Bullet Point1
                仿羊绒厨房垫-Bullet Point2
                仿羊绒厨房垫-Bullet Point3
                仿羊绒厨房垫-Bullet Point4
                仿羊绒厨房垫-Bullet Point5
                仿羊绒厨房垫-Generic Keyword
                仿羊绒厨房垫-Style
                '''
                for point_key in  ['Bullet Point1', 'Bullet Point2', 'Bullet Point3', 'Bullet Point4', 'Bullet Point5']:
                    point_key = glv['材质方向'] + '-' + point_key
                    #print(point_key, '文案:', my_list['fields'][point_key][0]['text'])
                    five_point_texts.append(parent_item_text_info['fields'][point_key][0]['text'].replace("#####", rp))
                product_description = get_feishu_field_text(
                    parent_item_text_info['fields'],
                    '仿羊绒厨房垫-Product Description',
                    '仿羊绒厨房垫-关于此艺术品'
                )
                搜索关键词 = get_feishu_field_text(
                    parent_item_text_info['fields'],
                    '仿羊绒厨房垫-Generic Keyword',
                    '仿羊绒厨房垫-搜索关键词'
                )
                设计 = get_feishu_field_text(
                    parent_item_text_info['fields'],
                    '仿羊绒厨房垫-Style',
                    '仿羊绒厨房垫-设计'
                )
                pass
            skc = items['fields']['SKC'][0]['text'] if 'SKC' in items['fields'] else ''
            if '' == skc:
                skc = items['fields']['花型命名'][0]['text'] if '花型命名' in items['fields'] else ''
            skc = skc.strip()
            print("skc =", skc)
            # if 'ARC5832-BEIGE ECHO' == skc:
            #     import pdb;pdb.set_trace()
            #     pass
            if '上架状态（RPA回传）' in items['fields']:
                print(items['fields']['上架状态（RPA回传）'], 'continue001')
                continue
            if 'FBM上架店铺' not in items['fields']:
                print('continue002')
                continue
            else: 
                if glv['gvar_shop_name'] != items['fields']['FBM上架店铺']:
                    print('continue003')
                    continue
                    pass
            try:
                if '花型风格' not in items['fields']:
                    items['fields']['花型风格'] = ''
                print("花型风格:", items['fields']['花型风格'])
                print("SKC:", skc)
                print("是否FBM测款:", items['fields']['是否FBM测款'])
                print("图所在NAS盘地址:", items['fields']['图所在NAS盘地址'][0]['text'])
                print("运营团队分配:", items['fields']['运营团队分配'])
                if '印花地毯' == glv['材质方向']:
                    title_size_list = [
                        '2X3',
                        '2X5',
                        '5X7',
                        '8X10',
                    ]
                    for s_txt in title_size_list:
                        title_field = glv['材质方向'] + f'-商品名称{s_txt}'
                        highlight_field = glv['材质方向'] + f'-Item Highlight {s_txt}'
                        if title_field not in parent_item_text_info['fields']:
                            items['fields'][f'{s_txt}的标题'] = ''
                            items['fields'][f'{s_txt}的副标题'] = ''
                        else:
                            items['fields'][f'{s_txt}的标题'] = parent_item_text_info['fields'][title_field]
                            items['fields'][f'{s_txt}的副标题'] = parent_item_text_info['fields'].get(highlight_field, [{}])[0].get('text', '') if highlight_field in parent_item_text_info['fields'] else ''
                        print(f"商品名称{s_txt}:", items['fields'][f'{s_txt}的标题'])
                if '仿羊绒厨房垫' == glv['材质方向']:
                    if '仿羊绒厨房垫-Item Name 20 x32" +20 x48"' not in parent_item_text_info['fields']:
                        items['fields'][f'20X32-20X48的标题'] = ''
                    else:
                        items['fields'][f'20X32-20X48的标题']= parent_item_text_info['fields']['仿羊绒厨房垫-Item Name 20 x32" +20 x48"']
                    items['fields'][f'20X32-20X48的副标题'] = get_feishu_field_text(
                        parent_item_text_info['fields'],
                        '仿羊绒厨房垫-Item Highlight 20 x32" +20 x48"'
                    )
                    print(
                        "商品名称20X32-20X48:",
                        items['fields']['20X32-20X48的标题']
                    )
                if '三明治户外垫' == glv['材质方向']:
                    title_size_list = [
                        '2X3',
                        '2.5X8',
                        '3X5',
                        '5X7',
                        '5X8',
                        '6X9',
                        '8X10',
                        '9X12',
                    ]
                    for s_txt in title_size_list:
                        title_field = glv['材质方向'] + f'-商品名称{s_txt}'
                        highlight_field = glv['材质方向'] + f'-Item Highlight {s_txt}'
                        if title_field not in parent_item_text_info['fields']:
                            items['fields'][f'{s_txt}的标题'] = ''
                        else:
                            items['fields'][f'{s_txt}的标题'] = parent_item_text_info['fields'][title_field]
                        items['fields'][f'{s_txt}的副标题'] = get_feishu_field_text(
                            parent_item_text_info['fields'],
                            highlight_field,
                        )
                        print(f"商品名称{s_txt}:", items['fields'][f'{s_txt}的标题'])
                        print(f"Item Highlight {s_txt}:", items['fields'][f'{s_txt}的副标题'])
                
                if '北美FBM' not in items['fields']['是否FBM测款'] and 'Amazon-US-FBM' not in items['fields']['是否FBM测款']:
                    print('continue004')
                    continue
                if '上架状态（RPA回传）' in items['fields']:
                    print(items['fields']['上架状态（RPA回传）'])
                    print('continue005')
                    continue
                print(timestamp_ms_to_datetime(items['fields']['创建日期']))
            except:
                continue
                pass

            #     continue
            #     pass
            if skc == '':
                continue
            if skc:
                skc_datas[skc] = {}
                skc_datas[skc]['pskc'] = pskc
                skc_datas[skc]['材质方向'] = glv['材质方向']
                skc_datas[skc]['创建日期'] = timestamp_ms_to_datetime(items['fields']['创建日期'])
                skc_datas[skc]['是否FBM测款'] = items['fields']['是否FBM测款']
                skc_datas[skc]['风格'] = items['fields']['花型风格']
                skc_datas[skc]['FBM上架店铺'] = items['fields']['FBM上架店铺']
                skc_datas[skc]['record_id'] = items['record_id']
                #import pdb;pdb.set_trace()

                skc_datas[skc]['运营团队'] = []
                for td in items['fields']['运营团队分配']:
                    skc_datas[skc]['运营团队'].append({"id": td['id']})
                skc_datas[skc]['对应运营'] = []
                if '对应运营' in items['fields']:
                    print("对应运营:", items['fields']['对应运营'])
                    for td in items['fields']['对应运营']:
                        skc_datas[skc]['对应运营'].append({"id": td['id']})
                        break
                for key_word in items['fields']:
                    if "的标题" in key_word:
                        #print(key_word)
                        tag = key_word.split("的")[0]
                        print("\n\n\n")
                        skc_datas[skc][tag] = {}
                        color_field = items['fields'].get('颜色', '')
                        if isinstance(color_field, list):
                            color = color_field[0].get('text', '') if color_field else ''
                        elif isinstance(color_field, dict):
                            color = color_field.get('text', '')
                        else:
                            color = color_field or ''
                        try:
                            # 标题文案同时兼容颜色和花型占位符。
                            skc_datas[skc][tag]['标题'] = items['fields'][key_word][0]['text']
                            skc_datas[skc][tag]['标题'] = (
                                re.sub(
                                    r'\bCOLOR\b',
                                    lambda _: color,
                                    skc_datas[skc][tag]['标题'].replace("##颜色##", color),
                                    flags=re.IGNORECASE,
                                )
                                .replace("#####", rp)
                            )
                        except:
                            items['fields'][key_word] = [{"text": ""}]
                            skc_datas[skc][tag]['标题'] = ''
                        # 副标题来自文案库中当前父体、当前尺寸对应的
                        # “印花地毯-Item Highlight {尺寸}”字段。上面已将其保存为
                        # items['fields']['{尺寸}的副标题']，这里读取该字段并替换颜色占位符。
                        subtitle_field = items['fields'].get(f'{tag}的副标题', '')
                        if isinstance(subtitle_field, list):
                            subtitle = subtitle_field[0].get('text', '') if subtitle_field else ''
                        elif isinstance(subtitle_field, dict):
                            subtitle = subtitle_field.get('text', '')
                        else:
                            subtitle = subtitle_field or ''
                        skc_datas[skc][tag]['副标题'] = (
                            re.sub(
                                r'\bCOLOR\b',
                                lambda _: color,
                                subtitle.replace("##颜色##", color),
                                flags=re.IGNORECASE,
                            )
                            .replace("#####", rp)
                        )
                        try:
                            print(tag, 'SKU :', skc + '-' + key_word.split("的")[0])
                        except:
                            import pdb;pdb.set_trace()
                            pass
                        print(tag, '标题:', items['fields'][key_word][0]['text'])
                        skc_datas[skc][tag]['图所在NAS盘地址'] = items['fields']['图所在NAS盘地址'][0]['text'] if '图所在NAS盘地址' in items['fields'] else ''
                        skc_datas[skc][tag]['颜色'] = ''
                        if '颜色' in items['fields']:
                            skc_datas[skc][tag]['颜色'] = items['fields']['颜色'][0]['text']
                        for index, five_point_text in enumerate(five_point_texts):
                            skc_datas[skc][tag][f'文案{index+1}'] = five_point_text
                            print(tag, f'文案{index+1}:', five_point_text)
                        skc_datas[skc][tag]['Product Description'] = product_description
                        skc_datas[skc][tag]['搜索关键词'] = 搜索关键词
                        skc_datas[skc][tag]['设计'] = 设计
                        skc_datas[skc][tag]['pattern'] = pattern
                        # 关键词和关于此艺术品
                        # print(tag, '文案:', items['fields'][key_word])

                        # print("\n\n\n")
            # print("skc=", skc)

        pass
    for skc in skc_datas:
        if '印花地毯' == skc_datas[skc]['材质方向']:
            continue
        print(skc, skc, skc_datas[skc]['材质方向'])
    return skc_datas

def get_guding_info(tk1='Jolyb8QBoaPzj6swf0cc6bqenlf', tk2='tblTer6BHOZRAQkB'):
    # https://wit0jhu6kvu.feishu.cn/base/Jolyb8QBoaPzj6swf0cc6bqenlf?table=tbliG0Lkhn0LeDZx&view=vewpNVnzsf
    shop_brands = {}
    for item in get_all_table_data('Jolyb8QBoaPzj6swf0cc6bqenlf', 'tbliG0Lkhn0LeDZx'):
        shop_brands[item['fields']['店铺'][0]['text']] = {"品牌": item['fields']['品牌'][0]['text'], "制造商": item['fields']['品牌'][0]['text']}
    guding_list = get_all_table_data(tk1, tk2)
    guding_info = {}
    for guding_index, guding_item in enumerate(guding_list):
        if not ('是否取用' in guding_item['fields'] and guding_item['fields']['是否取用'] == '是'):
            continue
        size_text = guding_item['fields']['size_text'][0]['text'].upper()
        guding_info[size_text] = {}
        for field_key in guding_item['fields']:
            try:
                guding_info[size_text][field_key] = guding_item['fields'][field_key][0]['text']
            except:
                guding_info[size_text][field_key] = guding_item['fields'][field_key]
        print(shop_brands)
        guding_info[size_text]['品牌'] = shop_brands[glv['gvar_shop_name']]['品牌']
        guding_info[size_text]['制造商'] = shop_brands[glv['gvar_shop_name']]['制造商']
    return guding_info
def add_fields(data):
    headers = {
        'Authorization': 'Bearer ' + tenant_access_token
    }
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/Jolyb8QBoaPzj6swf0cc6bqenlf/tables/tbl8zJByKrMGUJvM/records/batch_create"
    response = requests.post(url, headers=headers, data=json.dumps(data, ensure_ascii=False).encode('utf-8'))
    print(response.status_code)
    print(response.json())

def trim_text(text, words_list):
    import re
    for word in words_list:
        if word.upper().find(word.upper()) < 0:
            continue
        #text = text.replace(word, '')
        text = re.sub(word, "", text, flags=re.IGNORECASE)
    return text

def get_feishu_field_text(fields, *field_names):
    """按候选字段名读取飞书字段，兼容富文本列表、字典和普通字符串。"""
    for field_name in field_names:
        if field_name not in fields:
            continue
        value = fields[field_name]
        if isinstance(value, list):
            if not value:
                continue
            first_value = value[0]
            if isinstance(first_value, dict):
                result = first_value.get('text', first_value.get('name', ''))
            else:
                result = first_value
        elif isinstance(value, dict):
            result = value.get('text', value.get('name', ''))
        else:
            result = value
        if result is not None and str(result).strip() != '':
            return str(result).strip()
    return ''

def normalize_product_type(value):
    """亚马逊商品类型统一保存为去除首尾空格后的大写值。"""
    if value is None or pd.isna(value):
        return ''
    return str(value).strip().upper()

def build_child_seller_sku(skc, size_text, material_direction):
    """按材质方向生成子体卖家 SKU。"""
    suffix = 'FT-SMZ-POD' if material_direction == '三明治户外垫' else 'FT-POD'
    return f"{str(skc).strip()}-{str(size_text).upper()}{suffix}"

def _generate_single_shop(
    skcs,
    test_mode,
    target_pskc=None,
    preloaded_skc_datas=None,
    material_direction=None,
):
    material_direction = str(material_direction or '').strip()
    skcs = list(dict.fromkeys(
        str(skc).strip() for skc in (skcs or []) if str(skc).strip()
    ))
    if not material_direction:
        raise ValueError("材质方向不能为空")
    dir_path = r"D:\项目文件\AI自动上架\\"
    nas_dir_path = r"D:\NAS_download\\"
    skc_datas = copy.deepcopy(
        preloaded_skc_datas
        if preloaded_skc_datas is not None
        else get_table_data(material_direction=material_direction, skcs=skcs)
    )
    gd_dict = {}
    # guding_info = get_guding_info(tk1='Jolyb8QBoaPzj6swf0cc6bqenlf', tk2='tblTer6BHOZRAQkB')
    # https://wit0jhu6kvu.feishu.cn/base/Jolyb8QBoaPzj6swf0cc6bqenlf?table=tblNqTrr56XFMxYJ&view=vewBwTjUSL
    guding_info = get_guding_info(tk1='Jolyb8QBoaPzj6swf0cc6bqenlf', tk2='tblNqTrr56XFMxYJ')
    gd_dict['印花地毯'] = guding_info
    for st in guding_info:
        if 'Size' in guding_info[st]:
            guding_info[st]['商品尺寸'] = guding_info[st]['Size']
    # 厨房垫 https://wit0jhu6kvu.feishu.cn/base/Jolyb8QBoaPzj6swf0cc6bqenlf?table=tbl18dyMkn1KS8KH&view=vew294Yp4d
    guding_info = get_guding_info(tk1='Jolyb8QBoaPzj6swf0cc6bqenlf', tk2='tbl18dyMkn1KS8KH')
    gd_dict['仿羊绒厨房垫'] = guding_info
    for st in guding_info:
        if 'Size' in guding_info[st]:
            guding_info[st]['商品尺寸'] = guding_info[st]['Size']
    # 三明治 https://wit0jhu6kvu.feishu.cn/base/Jolyb8QBoaPzj6swf0cc6bqenlf?table=tblCffxAeXGnocAY&view=vewzrq6EcO
    guding_info = get_guding_info(tk1='Jolyb8QBoaPzj6swf0cc6bqenlf', tk2='tblCffxAeXGnocAY')
    gd_dict['三明治户外垫'] = guding_info
    for st in guding_info:
        if 'Size' in guding_info[st]:
            guding_info[st]['商品尺寸'] = guding_info[st]['Size']
    try:
        # https://wit0jhu6kvu.feishu.cn/base/Jolyb8QBoaPzj6swf0cc6bqenlf?table=tblMsoJKlpVM7iNJ&view=vew9NA2BTM
        w_list = []
        for word in get_all_table_data('Jolyb8QBoaPzj6swf0cc6bqenlf', 'tblMsoJKlpVM7iNJ'):
            w_list.append(word['fields']['文本'][0]['text'])
        words_list = [{'fields': {'违禁词列表': [{'text': ",".join(w_list)}]}}]
    except:
        words_list = [{'fields': {'违禁词列表': [{'text': 'durable, Persian, safe, adds, Foldable, ideal, atmosphere, elegant, seasonal, elements, elegance, special, unique, Keen, resistance, ruggable, loloi, lahome, nourison, safavieh, plush, luxury, mold, mildew', 'type': 'text'}]}, 'record_id': 'recvggQIvv7TrG'}]
    '''
    [{'fields': {'违禁词列表': [{'text': 'durable, Persian, safe, adds, Foldable, ideal, atmosphere, elegant, seasonal, elements, elegance, special, unique, Keen, resistance, ruggable, loloi, lahome, nourison, safavieh, plush, luxury, mold, mildew', 'type': 'text'}]}, 'record_id': 'recvggQIvv7TrG'}]
    '''
    # print(words_list[0]['fields']['违禁词列表'][0]['text'].split(","))
    # trim_text(text, words_list[0]['fields']['违禁词列表'][0]['text'].split(","))
    print(words_list)
    # https://wit0jhu6kvu.feishu.cn/base/Jolyb8QBoaPzj6swf0cc6bqenlf?table=tbl8zJByKrMGUJvM&view=vew4LsxOAs
    chrildren_list = get_all_table_data(
        'Jolyb8QBoaPzj6swf0cc6bqenlf',
        'tbl8zJByKrMGUJvM',
        material_direction=material_direction,
    )
    # https://wit0jhu6kvu.feishu.cn/base/Jolyb8QBoaPzj6swf0cc6bqenlf?table=tblZOU82tgaLNhAx&view=vewek5eTd0
    parent_list = get_all_table_data('Jolyb8QBoaPzj6swf0cc6bqenlf', 'tblZOU82tgaLNhAx')
    # 循环 skc_datas，获取每个SKC对应的标题、文案等信息
    need_skc = ''
    
    del_c_skc = []
    for c_skc in skc_datas.keys():
        # if skc_datas[c_skc]['pskc'] != first_pskc:
        #     del_c_skc.append(c_skc)
        #     continue
        if c_skc not in skcs:
            del_c_skc.append(c_skc)
            continue
    for c_skc in del_c_skc:
        try:
            if c_skc not in skc_datas:
                continue
        except:
            #import pdb;pdb.set_trace()
            pass
        del skc_datas[c_skc]
    notin_skcs = [skc for skc in skcs if skc not in list(skc_datas.keys())]
    print("notin_skcs = ", notin_skcs)
    # 检查每个skc合法性
    # img_names = [
    #     "120x170单椅",
    #     "2x5走廊",
    #     "5x7卧室",
    #     "8x10客厅",
    #     "switch",
    #     "2x3门口",
    #     "2x5厨房",
    #     "5x7客厅",
    #     "白底图120x170",
    #     "白底图2x3",
    #     "白底图2x5",
    #     "白底图5x7",
    # ]
    # glv['skc_err_message'] = ''
    # for ddk in skc_datas:
    #     print("ddk ~~~", ddk)
    #     if '图所在NAS盘地址' in skc_datas[ddk]['2X3']:
    #         nas_path = skc_datas[ddk]['2X3']['图所在NAS盘地址']
    #         if nas_path.find(ddk) < 0:
    #             glv['skc_err_message'] += f'skc:{ddk} 在【花型开发表】 【图所在NAS盘地址】中 疑似有误 {nas_path}，请修改成正确的nas地址，'
    #         df = pd.read_excel(r"D:\NAS_download\\" + f"{ddk}\\{ddk}data.xlsx")
    #         is_set = False
    #         for i_name in img_names:
    #             for name in list(df['文件名']):
    #                 if name.upper().find(i_name.upper()) >= 0:
    #                     is_set = True
    #                     break
    #             if not is_set:
    #                 glv['skc_err_message'] += f'skc:{ddk} 图片{i_name} 不存在'
            
    #         merge_names = "".join(list(df['文件名']))
    #         if merge_names.upper().find("封面图5X7客厅") >= 0 and merge_names.upper().find("封面图5X7卧室") >= 0:
    #             glv['skc_err_message'] += f'skc:{ddk} 封面图5X7客厅, 封面图5X7卧室 都存在，请修改成正确主图，'
    #         pass
    # if glv['skc_err_message'] != '':
    #     print("skc_err_messageskc_err_messageskc_err_messageskc_err_messageskc_err_messageskc_err_message")
    #     return
  
    pskc_list = []
    for skc in skc_datas.keys():
        pskc_list.append(skc_datas[skc]['pskc'])
    pskc_list = list(set(pskc_list))
    print("pskc_list = ", pskc_list)
    if len(pskc_list) == 0:
        print("DONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONE")
        glv['gvar_module1_done'] = True
        return
    first_pskc = target_pskc or pskc_list[0]
    if first_pskc not in pskc_list:
        print(f"当前批次没有父体 {first_pskc} 对应的 SKC，跳过")
        return
    #first_pskc = "LF8686-REL-NEW"
    print("first_pskc =", first_pskc)

    # 子体数据不一定包含“商品类型”，提前从当前父体记录读取作为兜底值。
    parent_product_type = ''
    for parent_item in parent_list:
        parent_fields = parent_item.get('fields', {})
        parent_sku = get_feishu_field_text(parent_fields, 'SKU', '卖家 SKU')
        if parent_sku != first_pskc:
            continue
        # 优先使用父体的新字段“产品类型”，旧字段“商品类型”仅作为兼容回退。
        parent_product_type = normalize_product_type(
            get_feishu_field_text(parent_fields, '产品类型', '商品类型')
        )
        break

    del_c_skc = []
    for c_skc in skc_datas.keys():
        if skc_datas[c_skc]['pskc'] != first_pskc:
            del_c_skc.append(c_skc)
            continue
    for c_skc in del_c_skc:
        try:
            if c_skc not in skc_datas:
                continue
        except:
            #import pdb;pdb.set_trace()
            pass
        del skc_datas[c_skc]
    if len(skc_datas) == 0:
        print("DONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONEDONE")
        glv['gvar_module1_done'] = True
        return
    # 以上 删除不是第一批psck次的skc
    add_data = {'records': []}
    existing_child_keys = set()
    for child_item in chrildren_list:
        child_fields = child_item.get('fields', {})
        child_skc = get_feishu_field_text(child_fields, 'SKC')
        child_shop = get_feishu_field_text(child_fields, '店铺')
        child_size = get_feishu_field_text(child_fields, 'size_text').upper()
        if child_skc and child_shop == glv['gvar_shop_name'] and child_size:
            existing_child_keys.add((child_skc, child_size))

    for skc in skc_datas.keys():
        skc_material_direction = skc_datas[skc]['材质方向']
        guding_info = gd_dict[skc_material_direction]
        for size_text in guding_info.keys():
            skc = skc.strip()
            size_text = size_text.upper()
            if (skc, size_text) in existing_child_keys:
                continue
            print("SKC", skc)
            seller_sku = build_child_seller_sku(
                skc, size_text, skc_material_direction
            )
            print("卖家 SKU", seller_sku)
            need_skc = skc
            add_data['records'].append({"fields": {
            "SKC": skc,
            "卖家 SKU": seller_sku,
            "风格": skc_datas[skc]['风格'],
            #"店铺": "岚风",
            "店铺": glv['gvar_shop_name'],
            "团队": skc_datas[skc]['运营团队'],
            "对应运营": skc_datas[skc]['对应运营'],
            "花型上架日期（RPA回传)": skc_datas[skc]['创建日期'],
            "商品尺寸": (
                guding_info[size_text].get('商品尺寸')
                or guding_info[size_text].get('Size')
                or size_text
            ),
            "材质方向": skc_datas[skc]['材质方向'],
            "size_text": size_text,

            }})
            print("add_fields~~~~~~~~~~~~~~~~~~~~", add_data)

    if test_mode:
        print(
            f"测试模式：跳过飞书子表新增，共模拟 "
            f"{len(add_data['records'])} 条记录"
        )
        simulated_children = []
        for simulated_index, record in enumerate(add_data['records']):
            simulated_record = copy.deepcopy(record)
            simulated_record['record_id'] = f"test_{simulated_index}"
            fields = simulated_record['fields']
            size_text = str(fields['size_text']).upper()
            fixed_fields = copy.deepcopy(
                gd_dict[fields['材质方向']].get(size_text, {})
            )
            fixed_fields.update(fields)
            fixed_fields['SKC'] = [{'text': fields['SKC']}]
            fixed_fields['size_text'] = [{'text': size_text}]
            fixed_fields['商品尺寸'] = [{
                'text': str(fields.get('商品尺寸', size_text))
            }]
            simulated_record['fields'] = fixed_fields
            simulated_children.append(simulated_record)
        chrildren_list = chrildren_list + simulated_children
    else:
        if add_data['records']:
            add_fields(add_data)
        chrildren_list = get_all_table_data(
            'Jolyb8QBoaPzj6swf0cc6bqenlf',
            'tbl8zJByKrMGUJvM',
            material_direction=material_direction,
        )
    # import pdb;pdb.set_trace()
    COLOR_DICT = {}
    need_skc = need_skc.strip()
    for c in chrildren_list:
        if get_feishu_field_text(c.get('fields', {}), '店铺') != glv['gvar_shop_name']:
            continue
        try:
            if need_skc == c['fields']['SKC'][0]['text']:
                continue
        except:
            continue
        if '颜色' not in c['fields']  or '父 SKU' not in c['fields']:
            continue
        # name = c['fields']['风格'] + c['fields']['团队'][0]['name']
        name = c['fields']['父 SKU'][0]['text']
        if name not in COLOR_DICT:
            COLOR_DICT[name] = {}
            COLOR_DICT[name]['psku'] = []
            COLOR_DICT[name]['colors'] = []
        COLOR_DICT[name]['colors'].append(c['fields']['颜色'][0]['text'])
        COLOR_DICT[name]['psku'].append(c['fields']['父 SKU'][0]['text'])

    # 定义团队风格变量
    style_teams = ''
    # 子类默认值
    data_to_excel_item = {}

    # dd = {need_skc: skc_datas[need_skc]}
    # list(list(dd.values())[0].values())[0]['图所在NAS盘地址'] 写入 nas_dir_path 保存为 nas_path.txt 
    nas_path_list = []
    # import pdb;pdb.set_trace()
    for skc, skc_data in skc_datas.items():
        print("ddk ~~~", skc)
        for size_data in skc_data.values():
            if not isinstance(size_data, dict):
                continue
            nas_path = size_data.get('图所在NAS盘地址', '')
            if nas_path:
                nas_path_list.append(nas_path)
                break
    print("nas_path_list ", nas_path_list)
    open(os.path.join( nas_dir_path, 'nas_path.txt' ), 'w', encoding='utf-8' ).write(",".join(nas_path_list))
    # open(os.path.join( nas_dir_path, './skcinfo/amz_' + skc + ".json"), 'w', encoding='utf-8' ).write(json.dumps(dd, indent=4, ensure_ascii=False))
    # skd_default_data = {need_skc:{}}
    data_to_excel = []
    ing_record_ids = []
    color_record_dict = {}
    for skc, datas in skc_datas.items():
        # if skc == 'ARM025-DARK BLUE':
        for chrildren_index, chrildren_item in enumerate(chrildren_list):
            if 'SKC' not in chrildren_item['fields'] or not chrildren_item['fields']['SKC']:
               continue
            if get_feishu_field_text(
                chrildren_item['fields'], '店铺'
            ) != glv['gvar_shop_name']:
                continue

            if skc == chrildren_item['fields']['SKC'][0]['text']:
                for size_text in datas.keys():
                    if size_text not in guding_info:
                        continue
                    if 'size_text' in chrildren_item['fields'] and chrildren_item['fields']['size_text'] and size_text.lower() == chrildren_item['fields']['size_text'][0]['text']:
                        chrildren_item['fields']['size_text'][0]['text'] = chrildren_item['fields']['size_text'][0]['text'].upper()
                    if 'size_text' in chrildren_item['fields'] and chrildren_item['fields']['size_text'] and size_text == chrildren_item['fields']['size_text'][0]['text']:
                        data = {
                            "fields": {
                                "卖家 SKU": build_child_seller_sku(
                                    skc,
                                    size_text,
                                    datas['材质方向'],
                                ),
                                "颜色": datas[size_text]['颜色'],
                                "商品名称": datas[size_text]['标题'],
                                "商品特性": datas[size_text]['文案1'],
                                "商品特性 (1)": datas[size_text]['文案2'],
                                "商品特性 (2)": datas[size_text]['文案3'],
                                "商品特性 (3)": datas[size_text]['文案4'],
                                "商品特性 (4)": datas[size_text]['文案5'],
                            }
                        }
                        style_teams = ''
                        if '风格' in chrildren_item['fields'] and chrildren_item['fields']['风格']:
                            style_teams += chrildren_item['fields']['风格']
                        if '团队' in chrildren_item['fields'] and chrildren_item['fields']['团队'] and len(chrildren_item['fields']['团队']) > 0:
                            style_teams += chrildren_item['fields']['团队'][0].get('name', '')
                        if '店铺' in chrildren_item['fields'] and chrildren_item['fields']['店铺']:
                            style_teams += chrildren_item['fields']['店铺']

                        data_keys = [
                            "商品类型",
                            "配送模板",
                            "货币",
                        #    "卖家 SKU",
                            "品牌",
                            "更新删除",
                        #    "关于此艺术品",
                            "制造商",
                        #    "商品编码",
                        #    "商品编码类型",
                            "产品类型关键字",
                            "您的价格",
                            "库存数量",
                            "父子关系",
                            "父 SKU",
                            "关系类型",
                            "商品变体主题",
                        #    "搜索关键词",
                        #    "颜色",
                            "色表",
                        #    "商品尺寸",
                            "外壳材料",
                            "背面材料",
                            "缝制品",
                            "箱子数量",
                            "结构类型",
                            "绒毛高度",
                            "商品形状",
                            "商品长边的长度",
                            "商品长度单位",
                            "商品短边的宽度",
                            "商品宽度单位",
                            "强制性警示声明",
                            "面料类型",
                            "原产国/地区",
                            "配送模板",
                            "市场价",
                            "状况",
                            "订单商品最大数量",
                            "特殊功能1",
                            "特殊功能2",
                            "特殊功能3",
                            "特殊功能4",
                            
                            "是否附带电池",
							"此商品是否使用电池或商品本身是电池？",
							"商品重量计量单位",
							"处理时间",
							"物品数量",
                        ]
                        # 根据data_keys填充 data_to_excel_item
                        for data_key in data_keys + ['卖家 SKU', '商品编码', '商品尺寸', '商品编码类型']:
                            if data_key in data_to_excel_item and data_to_excel_item[data_key] != '':
                                continue
                            if data_key not in chrildren_item['fields']:
                                continue
                            try:
                                data_to_excel_item[data_key] = chrildren_item['fields'][data_key][0]['text']
                            except:
                                data_to_excel_item[data_key] = chrildren_item['fields'][data_key]
                        data_item = copy.deepcopy(guding_info[size_text]) if size_text in guding_info else {}
                        for data_key in data_keys:
                            # 有值而且不为空赐跳过
                            if data_key in data_item and data_item[data_key] != '':
                                continue
                            # 空值处理
                            if data_key not in chrildren_item:
                                data_item[data_key] = ''
                                continue
                            try:
                                data_item[data_key] = chrildren_item['fields'][data_key][0]['text']
                            except:
                                data_item[data_key] = chrildren_item['fields'][data_key]
                        # data_item["商品编码类型"] = chrildren_item['fields']['商品编码类型'][0]['text'] if '商品编码类型' in chrildren_item['fields'] else 'ASIN'
                        # import pdb;pdb.set_trace()
                        data_item["商品编码"] = chrildren_item['fields']['商品编码'][0]['text'] if '商品编码' in chrildren_item['fields'] else ''
                        data_item["商品尺寸"] = chrildren_item['fields']['商品尺寸'][0]['text'] if '商品尺寸' in chrildren_item['fields'] else size_text
                        data_item["size_text"] = chrildren_item['fields']['size_text'][0]['text'] if 'size_text' in chrildren_item['fields'] else size_text
                        data_item["SKC"] = skc
                        data_item["卖家 SKU"] = build_child_seller_sku(
                            skc,
                            size_text,
                            datas['材质方向'],
                        )
                        datas[size_text]['标题'] = trim_text(datas[size_text]['标题'], words_list[0]['fields']['违禁词列表'][0]['text'].split(","))
                        data_item["商品名称"] = datas[size_text]['标题']
                        data_item["副标题"] = datas[size_text]['副标题']
                        data_item["颜色"] = trim_text(datas[size_text]['颜色'], words_list[0]['fields']['违禁词列表'][0]['text'].split(","))
                        # data_item["色表"] = trim_text(datas[size_text]['色表'], words_list[0]['fields']['违禁词列表'][0]['text'].split(","))
                        data_item["商品特性"] = trim_text(datas[size_text]['文案1'], words_list[0]['fields']['违禁词列表'][0]['text'].split(","))
                        data_item["商品特性.1"] = trim_text(datas[size_text]['文案2'], words_list[0]['fields']['违禁词列表'][0]['text'].split(","))
                        data_item["商品特性.2"] = trim_text(datas[size_text]['文案3'], words_list[0]['fields']['违禁词列表'][0]['text'].split(","))
                        data_item["商品特性.3"] = trim_text(datas[size_text]['文案4'], words_list[0]['fields']['违禁词列表'][0]['text'].split(","))
                        data_item["商品特性.4"] = trim_text(datas[size_text]['文案5'], words_list[0]['fields']['违禁词列表'][0]['text'].split(","))
                        data_item["搜索关键词"] = datas[size_text]['搜索关键词']
                        # 父体商品类型优先，避免被固定表中的小写旧值覆盖。
                        # 父体为空时才回退到尺寸文案或固定表/子体已有值。
                        data_item["商品类型"] = normalize_product_type(
                            parent_product_type
                            or datas[size_text].get('商品类型')
                            or data_item.get('商品类型')
                        )
                        data_item["FBM上架店铺"] = datas['FBM上架店铺']
                        data_item["Product Description"] = datas[size_text]['Product Description']
                        data_item["设计"] = datas[size_text]['设计']
                        data_item["材质方向"] = skc_datas[skc]['材质方向']
                        # pattern 保存在当前尺寸节点，而不是 SKC 顶层。
                        data_item["pattern"] = datas[size_text]['pattern']
                        # data_item["市场价"] 保留两位小数,去掉第三位小数
                        data_item["您的价格 USD (在亚马逊上出售, US)"] = round(float(data_item["您的价格 USD (在亚马逊上出售, US)"]), 2)

                        # https://wit0jhu6kvu.feishu.cn/base/Jolyb8QBoaPzj6swf0cc6bqenlf?table=tbl8zJByKrMGUJvM&view=vew4LsxOAs
                        data = {}
                        data['fields'] = {}
                        data['fields']['颜色'] = data_item["颜色"]
                        data['fields']['色表'] = data_item["颜色"]
                        # datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        if data_item["卖家 SKU"] not in color_record_dict:
                            color_record_dict[data_item["卖家 SKU"]] = {'record_id': chrildren_item['record_id'], 'data': data, 'skc': skc}
                        if '对应运营' in chrildren_item['fields']:
                            glv["gvar_对应运营id"] = chrildren_item['fields']['对应运营'][0]['id']
                        elif '运营团队分配' in chrildren_item['fields']:
                            glv["gvar_对应运营id"] = chrildren_item['fields']['运营团队分配'][0]['id']

                        data_to_excel.append(data_item)
                        print("skcskcskcskc=", skc, size_text, datas[size_text]['标题'], data['fields']['色表'])
                        # import pdb;pdb.set_trace()
                        pass
    # 补充data_to_excel的空数据
    for key in data_to_excel_item.keys():
        for index in range(len(data_to_excel)):
            # 不在则补充，或者空值则补充
            if key not in data_to_excel[index] or data_to_excel[index][key] == '':
                data_to_excel[index][key] = data_to_excel_item[key]

    parent_data_item = {}
    # 构造父类数据
    for parent_index, parent_item in enumerate(parent_list):
        #if '风格' not in parent_item['fields'] or '团队' not in parent_item['fields'] or '店铺' not in parent_item['fields']:
        #    continue
        parent_fields = parent_item.get('fields', {})
        parent_sku = get_feishu_field_text(parent_fields, 'SKU', '卖家 SKU')
        if first_pskc != parent_sku:
            continue
        parent_data_item['商品类型'] = normalize_product_type(
            get_feishu_field_text(parent_fields, '产品类型', '商品类型')
        )
        parent_data_item['产品类型关键字'] = get_feishu_field_text(parent_fields, '产品类型关键字')
        parent_data_item['制造商'] = get_feishu_field_text(parent_fields, '制造商', 'Manufacturer')
        parent_data_item['卖家 SKU'] = parent_sku
        parent_data_item['品牌'] = get_feishu_field_text(parent_fields, '品牌', 'Brand Name')
        parent_data_item['商品信息操作'] = get_feishu_field_text(parent_fields, '商品信息操作', '更新删除')
        parent_data_item['父条目的等级'] = get_feishu_field_text(parent_fields, '父条目的等级')
        parent_data_item['变体主题名称'] = get_feishu_field_text(parent_fields, '变体主题名称', '商品变体主题')
        # 保留旧 Header，兼容仍使用旧字段名的流程。
        parent_data_item['更新删除'] = parent_data_item['商品信息操作']
        parent_data_item['商品变体主题'] = parent_data_item['变体主题名称']
        parent_data_item['商品名称'] = get_feishu_field_text(parent_fields, '商品名称', 'Item Name')
        parent_data_item['商品编码类型'] = get_feishu_field_text(
            parent_fields, '商品编码类型', '商品编号类型'
        )
        parent_data_item['父子关系'] = get_feishu_field_text(parent_fields, '父子关系')
        parent_data_item['副标题'] = get_feishu_field_text(parent_fields, 'Item Highlight', '副标题')
        parent_data_item['色表'] = ''
        parent_data_item['Product Description'] = get_feishu_field_text(
            parent_fields, 'Product Description', '关于此艺术品'
        ) or f"GENIMO {parent_data_item.get('品牌', 'Rugs')} - Premium Quality Washable Area Rug"
        parent_data_item['Manufacturer'] = get_feishu_field_text(
            parent_fields, 'Manufacturer', '制造商'
        ) or 'GENIMO'
        parent_data_item['商品 ID'] = get_feishu_field_text(
            parent_fields, '商品 ID'
        ) or ''
        parent_data_item['商品编号类型'] = get_feishu_field_text(
            parent_fields, '商品编号类型', '商品编码类型'
        ) or 'ASIN'
        parent_data_item['Brand Name'] = get_feishu_field_text(
            parent_fields, 'Brand Name', '品牌'
        ) or parent_data_item.get('品牌', 'GENIMO')
        message = ''
        for index in range(len(data_to_excel)):
            # data_to_excel[index]['商品名称'] = data_to_excel[index]['品牌'] + " " + data_to_excel[index]['商品名称']
            data_to_excel[index]['商品名称'] = data_to_excel[index]['商品名称']
            '''
            for ii in range(1, 100):
                psku = parent_data_item["卖家 SKU"]
                if psku in COLOR_DICT and  data_to_excel[index]['颜色'] in COLOR_DICT[psku]['colors']:
                    # 设置新颜色，如果设置过了就不用再设置了
                    if color_record_dict[data_to_excel[index]["卖家 SKU"]]['skc'] not in skc_color_seted:
                        data_to_excel[index]['颜色'] = data_to_excel[index]['颜色'].split("-")[0] + "-" + str(ii)
                        message = "SKU:" + data_to_excel[index]["卖家 SKU"] + " 父 SKU:" + psku + " 颜色冲突，已经修改为:" + data_to_excel[index]['颜色']
                        print(message)
                    for rc in color_record_dict:
                        # 设置新颜色
                        if color_record_dict[rc]['skc'] == color_record_dict[data_to_excel[index]["卖家 SKU"]]['skc'] and color_record_dict[data_to_excel[index]["卖家 SKU"]]['skc'] not in skc_color_seted:
                            color_record_dict[rc]['data']['fields']['色表'] = data_to_excel[index]['颜色']
                            color_record_dict[rc]['data']['fields']['颜色'] = data_to_excel[index]['颜色']
                            skc_color_seted.append(color_record_dict[data_to_excel[index]["卖家 SKU"]]['skc'])
                            COLOR_DICT[psku]['colors'].append(data_to_excel[index]['颜色'])
                            print(rc, color_record_dict[rc]['data']['fields']['颜色'])

                            import pdb;pdb.set_trace()
                            pass
                else:
                    break
                '''
            #if color_record_dict[data_to_excel[index]["卖家 SKU"]]['skc'] not in skc_color_seted:
            #    skc_color_seted.append(color_record_dict[data_to_excel[index]["卖家 SKU"]]['skc'])
            #    COLOR_DICT[psku]['colors'].append(data_to_excel[index]['颜色'])
                
            data_to_excel[index]['色表'] = data_to_excel[index]['颜色']
            dataset = {}
            dataset['fields'] = {}
            dataset['fields']['父 SKU'] = parent_data_item["卖家 SKU"]
            dataset['fields']['色表'] = data_to_excel[index]['颜色']
            dataset['fields']['颜色'] = data_to_excel[index]['颜色']
            dataset['fields']['花型上架日期（RPA回传)'] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            dataset['fields']['上架状态（RPA回传）'] = 'ing'
            # if data_to_excel[index]["卖家 SKU"] == 'DT8202-SAFFRON VEIL-2X3FT-POD':
            #     import pdb;pdb.set_trace()
            #     pass
            color_record_dict[data_to_excel[index]["卖家 SKU"]]['data'] = copy.deepcopy(dataset)

            data_to_excel[index]['商品编码'] = ''
            data_to_excel[index]['父 SKU'] = parent_data_item["卖家 SKU"]
            data_to_excel[index]["父条目的库存单位"] = parent_data_item["卖家 SKU"]
                        
        #if message != '':
        #    send(message)
        data_to_excel.insert(0, parent_data_item)
        break
    
    # 判断这个skc颜色是否设置过
    skc_color_seted = []
    for i in range(1, len(data_to_excel)):
        psku = data_to_excel[i]['父 SKU']
        # color_record_dict[data_to_excel[i]['卖家 SKU']]['data']['fields']['颜色']
        # color_record_dict[data_to_excel[i]['卖家 SKU']]['data']['fields']['色表']
        # data_to_excel[i]['颜色']
        if color_record_dict[data_to_excel[i]['卖家 SKU']]['skc'] in skc_color_seted:
            continue
        print(i, psku, data_to_excel[i]['卖家 SKU'], data_to_excel[i]['颜色'])
        skc_color_seted.append(color_record_dict[data_to_excel[i]['卖家 SKU']]['skc'])
        if data_to_excel[i]['父 SKU'] not in COLOR_DICT:
            COLOR_DICT[data_to_excel[i]['父 SKU']] = {'colors': []}
        for ii in range(1, 100):
            if data_to_excel[i]['颜色'] in COLOR_DICT[data_to_excel[i]['父 SKU']]['colors']:
                data_to_excel[i]['颜色'] = data_to_excel[i]['颜色'].split("-")[0] + "-" + str(ii)
                data_to_excel[i]['色表'] = data_to_excel[i]['色表'].split("-")[0] + "-" + str(ii)
            else:
                break
        for rc in color_record_dict:
            # 设置新颜色
            if color_record_dict[rc]['skc'] == color_record_dict[data_to_excel[i]["卖家 SKU"]]['skc']:
                color_record_dict[rc]['data']['fields']['色表'] = data_to_excel[i]['颜色']
                color_record_dict[rc]['data']['fields']['颜色'] = data_to_excel[i]['颜色']
                COLOR_DICT[psku]['colors'].append(data_to_excel[i]['颜色'])
                print(rc, color_record_dict[rc]['data']['fields']['颜色'])
        for fo in range(1, len(data_to_excel)):
            if  data_to_excel[i]["卖家 SKU"].find(color_record_dict[data_to_excel[fo]["卖家 SKU"]]['skc']) >= 0:
                print(fo, data_to_excel[i]["卖家 SKU"], color_record_dict[data_to_excel[fo]["卖家 SKU"]]['skc'], data_to_excel[i]['颜色'])
                data_to_excel[fo]['颜色'] = data_to_excel[i]['颜色']
                data_to_excel[fo]['色表'] = data_to_excel[i]['色表']
    reset_fields_datas = {
        "records": [
        ]
        }
    for k, v in color_record_dict.items():
        # reset_fields(v['data'], v['record_id'], 'Jolyb8QBoaPzj6swf0cc6bqenlf', 'tbl8zJByKrMGUJvM')
        reset_fields_datas['records'].append({"record_id": v['record_id'], "fields": v['data']['fields']})
    if test_mode:
        print(
            f"测试模式：跳过飞书子表更新，共模拟 "
            f"{len(reset_fields_datas['records'])} 条记录"
        )
    elif reset_fields_datas['records']:
        reset_fields_batch_update_v2(reset_fields_datas, "", "")

    safe_shop_name = re.sub(
        r'[<>:"/\\|?*\x00-\x1f]',
        '_',
        str(glv['gvar_shop_name']).strip()
    ).strip().rstrip('.')[:150]
    if not safe_shop_name:
        raise ValueError(f"店铺名称无法生成合法文件名：{glv['gvar_shop_name']!r}")
    safe_parent_sku = re.sub(
        r'[<>:"/\\|?*\x00-\x1f]', '_', str(first_pskc).strip()
    ).strip().rstrip('.')[:150]
    if not safe_parent_sku:
        raise ValueError(f"父体 SKU 无法生成合法文件名：{first_pskc!r}")
    output_feishu_path = os.path.join(
        dir_path,
        f"output_feishu_table_data_{safe_shop_name}_{safe_parent_sku}.xlsx"
    )
    output_df = pd.DataFrame(data_to_excel)
    # 历史固定表中如果仍带有“关于此艺术品”，合并到新字段后删除旧列，
    # 保证父体和子体只写入同一个 Product Description 列。
    if '关于此艺术品' in output_df.columns:
        if 'Product Description' not in output_df.columns:
            output_df['Product Description'] = output_df['关于此艺术品']
        else:
            output_df['Product Description'] = output_df['Product Description'].fillna(
                output_df['关于此艺术品']
            )
        output_df.drop(columns=['关于此艺术品'], inplace=True)
    parent_export_fields = [
        'Product Description', 'Manufacturer', '商品 ID', '商品编号类型', 'Brand Name',
        '商品信息操作', '父条目的等级', '变体主题名称'
    ]
    for field_name in parent_export_fields:
        if field_name not in output_df.columns:
            output_df[field_name] = ''
    if '商品类型' not in output_df.columns:
        raise ValueError("output_feishu_table_data.xlsx 缺少商品类型字段")
    # 最终保存前再次统一大写，防止任何固定表或旧字段在后续流程中覆盖。
    output_df['商品类型'] = output_df['商品类型'].map(normalize_product_type)
    output_df.to_excel(output_feishu_path, index=False)

    # 保存后重新读取，防止字段只存在于内存但没有真正写入文件。
    saved_df = pd.read_excel(output_feishu_path)
    missing_saved_fields = [
        field_name for field_name in parent_export_fields
        if field_name not in saved_df.columns
    ]
    if missing_saved_fields:
        raise ValueError(f"output_feishu_table_data.xlsx 保存后缺少字段：{missing_saved_fields}")
    empty_saved_parent_fields = [
        field_name for field_name in parent_export_fields
        if len(saved_df) == 0
        or pd.isna(saved_df.iloc[0][field_name])
        or str(saved_df.iloc[0][field_name]).strip() == ''
    ]
    if empty_saved_parent_fields:
        raise ValueError(
            f"output_feishu_table_data.xlsx 父体行字段为空：{empty_saved_parent_fields}"
        )
    invalid_product_types = [
        str(value).strip()
        for value in saved_df['商品类型'].dropna().tolist()
        if str(value).strip() != str(value).strip().upper()
    ]
    if invalid_product_types:
        raise ValueError(
            f"output_feishu_table_data.xlsx 商品类型未转为大写：{invalid_product_types}"
        )
    print(
        "商品类型保存验证：",
        sorted(set(saved_df['商品类型'].dropna().astype(str).str.strip().tolist()))
    )
    print("父体字段保存验证：", {
        field_name: saved_df.iloc[0][field_name] if len(saved_df) else ''
        for field_name in parent_export_fields
    })
    print("gvar_对应运营id =", glv['gvar_对应运营id'])
    return output_feishu_path


def m(test_mode=None, material_direction=None):
    """遍历全部店铺，并生成指定材质方向的可处理 SKC。"""
    if test_mode is None:
        test_mode = TEST_MODE
    material_direction = str(material_direction or '').strip()
    if not material_direction:
        raise ValueError(
            "必须指定材质方向；命令行请使用 "
            "--material-direction \"材质方向\""
        )

    skcs_path = r"D:\NAS_download\SKCS.txt"
    if not os.path.isfile(skcs_path):
        raise FileNotFoundError(f"SKC 清单不存在：{skcs_path}")
    with open(skcs_path, 'r', encoding='utf-8-sig') as skcs_file:
        skcs = list(dict.fromkeys(
            line.strip() for line in skcs_file if line.strip()
        ))
    if not skcs:
        raise ValueError(f"SKC 清单为空：{skcs_path}")

    original_shop_name = glv.get('gvar_shop_name', '')
    original_shop_name_list = glv.get('gvar_shop_name_list', [])
    shop_name_list = list(dict.fromkeys(original_shop_name_list))
    if not shop_name_list and original_shop_name:
        shop_name_list = [original_shop_name]
    if not shop_name_list:
        raise ValueError("gvar_shop_name_list 为空，且未设置 gvar_shop_name")

    mode_name = "测试模式（不写飞书子表）" if test_mode else "正式模式（写飞书子表）"
    material_name = material_direction or "全部材质"
    print(
        f"开始批量生成：{mode_name}，店铺 {len(shop_name_list)} 个，"
        f"SKCS.txt 共 {len(skcs)} 个 SKC，材质方向：{material_name}"
    )

    generated_paths = []
    failed_batches = []
    try:
        for shop_index, shop_name in enumerate(shop_name_list, start=1):
            glv['gvar_shop_name'] = shop_name
            print(
                f"\n========== 店铺 [{shop_index}/{len(shop_name_list)}]："
                f"{shop_name} =========="
            )
            shop_skc_datas = get_table_data(
                material_direction=material_direction,
                skcs=skcs,
            )
            matched_skcs = [skc for skc in skcs if skc in shop_skc_datas]
            parent_batches = {}
            for skc in matched_skcs:
                parent_sku = shop_skc_datas[skc]['pskc']
                parent_batches.setdefault(parent_sku, []).append(skc)

            missing_skcs = [skc for skc in skcs if skc not in shop_skc_datas]
            print(
                f"店铺匹配 SKC {len(matched_skcs)} 个，"
                f"未匹配 {len(missing_skcs)} 个，父体批次 {len(parent_batches)} 个"
            )
            for batch_index, (parent_sku, batch_skcs) in enumerate(
                parent_batches.items(), start=1
            ):
                print(
                    f"开始父体批次 [{batch_index}/{len(parent_batches)}]："
                    f"{parent_sku}，SKC {len(batch_skcs)} 个"
                )
                try:
                    output_path = _generate_single_shop(
                        batch_skcs,
                        test_mode,
                        target_pskc=parent_sku,
                        preloaded_skc_datas=shop_skc_datas,
                        material_direction=material_direction,
                    )
                except Exception as exc:
                    failed_batches.append((shop_name, parent_sku, str(exc)))
                    print(
                        f"父体批次失败，继续下一批：{shop_name} / "
                        f"{parent_sku}，原因：{exc}"
                    )
                else:
                    if output_path:
                        generated_paths.append(output_path)
    finally:
        glv['gvar_shop_name'] = original_shop_name
        glv['gvar_shop_name_list'] = original_shop_name_list

    print(
        f"批量生成结束：成功生成 {len(generated_paths)} 个 Excel，"
        f"失败 {len(failed_batches)} 个父体批次"
    )
    if failed_batches:
        failure_message = "；".join(
            f"{shop_name}/{parent_sku}: {reason}"
            for shop_name, parent_sku, reason in failed_batches
        )
        raise RuntimeError(f"以下父体批次生成失败：{failure_message}")
    return generated_paths


def get_material_direction_arg(args):
    """兼容命令行、RPA 字典和直接字符串形式的材质方向入参。"""
    if args is None:
        return None
    if isinstance(args, str):
        return args.strip() or None
    if isinstance(args, dict):
        value = args.get('material_direction', args.get('材质方向'))
        return str(value).strip() if value else None
    value = getattr(args, 'material_direction', None)
    return str(value).strip() if value else None


def main(args=None):
    return m(material_direction=get_material_direction_arg(args))


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="按材质方向批量生成亚马逊上架数据")
    parser.add_argument(
        "--material-direction",
        required=True,
        help="必须指定要精确筛选的材质方向",
    )
    main(parser.parse_args())
    # if not OUT_FILE_RES:
    #     m()
