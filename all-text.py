import os
import pandas as pd

# 获取当前目录下的所有文件名
files = os.listdir()

# 初始化统计数据的字典和日期集合
statistics = {}
all_dates = set()

# 遍历文件名列表，解析文件名并统计
for file_name in files:
    if file_name.endswith('.jpg') or file_name.endswith('.png'):  # 只考虑图片文件，根据实际情况修改后缀名

        parts = file_name.split('-')  # 使用'-'分割文件名

        if len(parts) == 3:  # 确保分割后有三部分
            name = parts[0]  # 第一部分是姓名
            date = parts[1]  # 第二部分是请假日期和节次
            date_parts = date.split('.')  # 使用'.'分割日期和节次

            if len(date_parts) == 2:  # 确保日期部分包含两个部分：月和日
                month = date_parts[0]  # 获取月份部分
                day = date_parts[1]    # 获取日期部分

                # 格式化月和日，保证输出为两位数
                if len(month) == 1:
                    month = '0' + month
                if len(day) == 1:
                    day = '0' + day

                formatted_date = f"{month}.{day}"  # 格式化后的日期

                if name not in statistics:
                    statistics[name] = {'请假次数': 0, '日期': {}}

                # 根据文件名中的信息更新统计数据
                statistics[name]['请假次数'] += 1
                if formatted_date not in statistics[name]['日期']:
                    statistics[name]['日期'][formatted_date] = []

                # 添加请假节次（去掉文件尾）
                times = parts[2].split('.')[0]  # 获取请假节次，去掉文件尾
                statistics[name]['日期'][formatted_date].append(times)
                all_dates.add(formatted_date)

# 将统计结果转换成DataFrame
data = []

# 第一行标题
header = ['姓名', '', '', '请假次数']
header.extend(sorted(all_dates))
data.append(header)

for name, info in statistics.items():
    row = [name, '', '', info['请假次数']]
    dates = info['日期']
    for date in sorted(all_dates):
        if date in dates:
            times = ' '.join(dates[date])
            row.append(times)
        else:
            row.append('')
    data.append(row)

df = pd.DataFrame(data)

# 写入Excel文件到当前目录
excel_file = os.path.join(os.getcwd(), '假条统计结果.xlsx')
with pd.ExcelWriter(excel_file, engine='openpyxl', mode='w') as writer:
    df.to_excel(writer, index=False, header=False)
