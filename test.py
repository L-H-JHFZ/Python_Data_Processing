import requests
import json

# 定义接口URL
url = '/px-settlement-infpubquery/export/searchDynamicTableData/1006'

# 定义查询参数和分页参数
params = {
    "hashMap": {
        "varDate": "2024-12-26"  # 替换为实际要查询的日期
    },
    "pageInfo": {
        "pageNum": 1,  # 页码，从1开始
        "pageSize": 10  # 每页条数，根据实际需求调整
    }
}

# 定义请求头
headers = {
    'Content-Type': 'application/json'
}

# 发送POST请求
response = requests.post(url, data=json.dumps(params), headers=headers)

# 解析响应
if response.status_code == 200:
    response_data = response.json()
    if response_data['status'] == 0 and response_data['message'] == 'Success':
        # 获取业务数据
        business_data = response_data['data']
        total_num = business_data['totalNum']  # 总条数
        page_size = business_data['pageSize']  # 每页条数
        page_num = business_data['pageNum']  # 当前页码
        status = business_data['status']  # 数据查询状态
        list_data = business_data['list']  # 当前页数据列表
        
        # 打印响应数据
        print("Total Number of Records:", total_num)
        print("Page Size:", page_size)
        print("Current Page Number:", page_num)
        print("Data Query Status:", status)
        
        for item in list_data:
            print("Title:", item['Title'])
            print("Content:", item['Content'])  # 注意：Content是longtext类型，可能很长
            print("Attachment:", item['Attachment'])
            print("Operate Date:", item['Operate_date'])
    else:
        print("Error:", response_data['message'])
else:
    print("Failed to retrieve data. Status code:", response.status_code)