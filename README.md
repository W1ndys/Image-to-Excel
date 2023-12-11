# Image-to-Excel
Python实现假条统计并合并到Excel表格中

> 一个简单的Python脚本，全程用ChatGPT生成+改报错，最终完美运行
>
> 请注意，<font color='red'>本项目只适合我个人使用，因为代码非常具有定制化，仅供查阅借鉴</font>

# 更新日志

2023年12月11日19:47:18   push到GitHub上

# 使用方法

首选运行下面安装必备的库

```
pip install pandas
```

然后把上面文件中 `all-text.py`，`一键执行假条统计.bat`以及你要统计的假条图片放在同一个文件夹中

请注意：！！！

假条的命名格式一定按照下面要求命名

**命名格式为：姓名-日期**

**文件后缀一定是jpg或者是png，不能是其他格式**

**例如：张三-02.03，李四-12.31。**

脚本已经集成好了

双击运行 `一键执行假条统计.bat`即可生成表格 `假条统计结果`

# 设计背景

继昨天用ai跑出来了一键修改日期的脚本

另外作为某班委，除了每周交一次统计表之外，另外还需要每周上交班内的请假条

我需要对着请假条一张张的改统计表

这种方法又笨又慢

我又想到，能不能继续借用ai的力量写一个脚本

用来实现把假条信息统计在一个表格中

一开始ai并不聪明，这次报错不多，但是实现的效果并不如意

经过一遍遍的训练

最终终于实现了假条统计，并且很直观的可以看到哪个同学分别在哪一天请过假

~~本来我想继续添加一个功能，用来实现把请的哪一节的课也输出出来，无奈一直没调好，太费时间了，等调好之后再试试~~

果不其然

成功了！

# 效果预览

![img](file:///F:\QQ message\Tencent Files\2769731875\nt_qq\nt_data\Pic\2023-12\Ori\839c87ea435345dc261d1bb8ed67a2e4.png)

# 与ai的对话节选

### 问题需求

要求统计当前目录下图片文件名中姓名部分出现的次数，并将结果输出到Excel表格。图片文件名格式为“姓名-日期”。希望Excel表格的第一列是姓名，第四列是姓名出现的次数。

### 代码实现

```
pythonCopy codeimport os
import pandas as pd

current_directory = os.getcwd()
name_count = {}

# 遍历文件进行统计
for filename in os.listdir(current_directory):
    if filename.lower().endswith(('.jpg', '.png')):
        # 提取文件名中的姓名部分
        name = filename.split('-')[0]
        # 统计姓名出现的次数
        if name in name_count:
            name_count[name] += 1
        else:
            name_count[name] = 1

# 将统计结果转换为DataFrame并输出到Excel表格
data = [{'姓名': name, '姓名出现次数': count} for name, count in name_count.items()]
df = pd.DataFrame(data)

output_file = os.path.join(current_directory, '姓名统计结果.xlsx')
df.to_excel(output_file, index=False, columns=['姓名', '姓名出现次数'])
print(f"已将姓名统计结果写入 {output_file}")
```

### 新需求

增加功能：在Excel表格的第一行输出标题，第一列是姓名，第二列和第三列留空，第四列是姓名出现的次数。第五列以及之后是统计到的日期，如果某姓名在某个日期出现了，则在相应单元格置1。

### 调整后的代码实现

```
pythonCopy codeimport os
import pandas as pd

# 获取当前目录
current_directory = os.getcwd()
name_count = {}
date_set = set()

# 遍历文件进行统计
for filename in os.listdir(current_directory):
    if filename.lower().endswith(('.jpg', '.png')):
        parts = filename.split('-')
        if len(parts) == 2:
            name, date = parts
            date_set.add(date)
            if name in name_count:
                if date in name_count[name]:
                    name_count[name][date] += 1
                else:
                    name_count[name][date] = 1
            else:
                name_count[name] = {date: 1}

# 生成日期列表，并将统计结果输出到Excel表格
dates = sorted(list(date_set))
data = []
for name, counts in name_count.items():
    row = [name, '', '', sum(counts.values())]
    for date in dates:
        if date in counts:
            row.append(1)
        else:
            row.append('')
    data.append(row)

output_file = os.path.join(current_directory, '姓名统计结果.xlsx')
df = pd.DataFrame(data, columns=['姓名', '', '', '次数'] + [date.rstrip('.jpg').rstrip('.png') for date in dates])

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False)
```

# 运行环境

Python

pandas库