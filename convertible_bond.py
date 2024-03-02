import argparse
import requests
import json
import openpyxl
from lxml import html
import re


def list_cb(cookie):
    url = "https://www.jisilu.cn/webapi/cb/list/"

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
        # 设置 User-Agent
        "Columns": "1,70,2,3,5,6,11,12,14,15,16,29,30,32,34,44,46,47,50,52,53,54,56,57,58,59,60,62,63,67",
        "Init": "1",
    }

    cookies = {
        "Cookie": cookie,
        # 设置 Cookie
    }

    response = requests.get(url, headers=headers, cookies=cookies)

    if response.status_code == 200:
        data = response.text
        # 处理数据
        # 解析 JSON 字符串
        parsed_json = json.loads(data)

        # 获取数据中的bond_id值
        data_list = parsed_json["data"]
        print('code:', parsed_json["code"], 'msg:', parsed_json["msg"], 'data size:', len(data_list))

        for item in data_list:
            # 移除指定的键
            parsed_json.pop("icons", None)

            put_ytm = detail(cookie, item["bond_id"])
            item["put_ytm"] = put_ytm

            # 将 put_ytm 插入到字典的最前面
            item_keys = list(item.keys())
            item_keys.insert(0, "put_ytm")
            item_values = [item.pop(key) for key in item_keys if key in item]
            item.update(zip(item_keys, item_values))

        # 根据 put_ytm 的值进行排序（降序）
        sorted_bond_data = sorted(data_list, key=lambda x: float(x.get("put_ytm", "0%").rstrip("%")), reverse=True)

        # 打印数据列表到文件
        write_to_excel(sorted_bond_data)

    else:
        print("Error:", response.status_code)


def write_to_excel(data):

    # 创建一个新的 Excel 工作簿
    wb = openpyxl.Workbook()

    # 获取活动的工作表
    ws = wb.active

    # 写入表头
    header = ["到期税后收益率", "代码", "转债名称", "现价", "评级", "到期时间", "剩余年限/年", "涨跌幅"]
    ws.append(header)

    # 将数据写入工作表
    for item in data:
        ws.append(
            [item["put_ytm"], item["bond_id"], item["bond_nm"], item["price"], item["rating_cd"], item["maturity_dt"], str(item["year_left"]), str(item["increase_rt"]) + '%'])

    # 保存工作簿
    wb.save("output.xlsx")

    print("数据已写入到 output.xlsx 文件中")


def write_to_txt(data):
    # 打印数据列表到文本文件，每个值之间添加制表符
    with open("output.txt", "w", encoding="utf-8") as file:
        for item in data:
            # 将字典中的值转换为字符串
            values_as_strings = [str(value) for value in item.values()]
            # 连接值并写入文件
            file.write("\t".join(values_as_strings) + "\n")


def detail(cookie, code):
    url = "https://www.jisilu.cn/data/convert_bond_detail/" + code

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
        # 设置 User-Agent
    }

    cookies = {
        "Cookie": cookie,
        # 设置 Cookie
    }

    response = requests.get(url, headers=headers, cookies=cookies)

    if response.status_code == 200:
        data = response.text
        # 处理数据
        return Yield_to_Maturity_After_Taxes(data)
    else:
        print("Error:", response.status_code)
    return -999


def Yield_to_Maturity_After_Taxes(content):

    # 使用 lxml 解析 HTML
    tree = html.fromstring(content)

    code = tree.xpath('//*[@id="tc_data"]/div/div[1]/table[1]/tr/td/div/div[1]/text()')[1]

    name = tree.xpath('//*[@id="tc_data"]/div/div[1]/table[1]/tr/td/div/div[1]/span')

    # 如果找到了元素
    if name:
        # 获取元素的HTML内容
        element_html = html.tostring(name[0], pretty_print=True, encoding='unicode')
        print(element_html)
        # 获取子字符串 "123044" 在字符串中的位置
        index = element_html.find(code)

        # 如果找到了子字符串
        if index != -1:
            # 截取从字符串开头到子字符串位置的部分
            span = element_html[:index].strip()
            # <span class="font_18">红相转债</span>
            tree_span = html.fromstring(span)

            name = tree_span.xpath('text()')[0]
        else:
            print("未找到子字符串")

    # 使用 XPath 表达式定位到特定的元素
    value = tree.xpath('//*[@id="tc_data"]/div/div[1]/table[1]/tr[3]/td[3]/text()')[0]
    # 输出获取到的值
    print(code, name, value)

    # 使用正则表达式匹配模式
    pattern = r'-?\d+\.\d+%'  # 匹配一个可选的负号，后面跟着一个或多个数字，然后是一个小数点，接着是一个或多个数字，最后是一个百分号
    match = re.search(pattern, value)

    # 如果找到了匹配项
    if match:
        result = match.group()  # 获取匹配到的字符串
        print(f'计算到期税后收益率返回值:{result}')
        return result
    else:
        print("未找到匹配到期税后收益率,返回值为-999")
    return -999


def main(cookie):
    list_cb(cookie)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("cookie", help="jisilu cookie")

    options = parser.parse_args()
    main(options.cookie)
