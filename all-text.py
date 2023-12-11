import os
import pandas as pd

# 获取当前工作目录
current_directory = os.getcwd()

# 初始化统计数据的字典和日期集合
name_count = {}
date_set = set()

# 遍历当前目录中的文件
for filename in os.listdir(current_directory):
    if filename.lower().endswith(('.jpg', '.png')):  # 假设这里只有图片文件
        # 提取文件名中的姓名和日期部分
        parts = filename.rsplit('-', 1)
        if len(parts) == 2:
            name, date = parts
            date_set.add(date)
            # 统计姓名出现的次数
            if name in name_count:
                if date in name_count[name]:
                    name_count[name][date] += 1
                else:
                    name_count[name][date] = 1
            else:
                name_count[name] = {date: 1}

# 将日期转换为列表，并按照日期排序
dates = sorted(list(date_set))

# 将统计结果转换为DataFrame
data = []
for name, counts in name_count.items():
    row = [name, '', '', sum(counts.values())]  # 初始值为姓名和次数
    for date in dates:
        if date in counts:
            row.append(1)
        else:
            row.append('')
    data.append(row)

# 将DataFrame写入Excel文件到当前目录
output_file = os.path.join(current_directory, '假条统计结果.xlsx')
df = pd.DataFrame(data, columns=['姓名', '', '', '次数'] + [date.rstrip('.jpg').rstrip('.png') for date in dates])

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False)
