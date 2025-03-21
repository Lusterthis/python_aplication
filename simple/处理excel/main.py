import pandas as pd
import random
from collections import defaultdict


def find_header_row(sheet_df):
    """找到包含'序号'的表头行"""
    for idx, row in sheet_df.iterrows():
        if '序号' in str(row.iloc[0]):
            return idx
    return 0


# 读取Excel文件
file_path = "周一晚马原课堂分组_副本.xlsx"
sheets = pd.read_excel(file_path, sheet_name=None, header=None)

# 合并所有sheet的数据
all_data = []
for sheet_name, sheet_df in sheets.items():
    header_row = find_header_row(sheet_df)
    data = sheet_df.iloc[header_row + 1:]
    data = data.iloc[:, :4].dropna(how='all')
    data.columns = ['序号', '学号', '姓名', '班级']
    all_data.append(data)

full_df = pd.concat(all_data).reset_index(drop=True)

# 处理缺失值并转换为字符串
full_df['班级'] = full_df['班级'].fillna('').astype(str)

# 按班级分组（修复NaN问题）
classes = {
    '双学位班': full_df[full_df['班级'].str.contains('双学位', na=False)],
    '中澳班': full_df[full_df['班级'].str.contains('中澳', na=False)],
    '卓越班': full_df[full_df['班级'].str.contains('卓越', na=False)],
    '工业设计班': full_df[full_df['班级'].str.contains('工业设计', na=False)]
}

# 打乱每个班级的学生顺序
for cls in classes:
    classes[cls] = classes[cls].sample(frac=1).reset_index(drop=True)

# 初始化10个分组
groups = defaultdict(list)

# 分配工业设计班学生到随机组
if not classes['工业设计班'].empty:
    target_group = random.randint(0, 9)
    groups[target_group].extend(classes['工业设计班'].to_dict('records'))

# 分配其他班级学生
for cls in ['双学位班', '中澳班', '卓越班']:
    students = classes[cls]
    num_groups = 10

    per_group = len(students) // num_groups
    remainder = len(students) % num_groups

    pointer = 0
    for group_num in range(num_groups):
        count = per_group + (1 if group_num < remainder else 0)
        groups[group_num].extend(
            students.iloc[pointer:pointer + count].to_dict('records'))
        pointer += count

# 打乱各组内部顺序并格式化输出
result = []
for group_num in range(10):
    random.shuffle(groups[group_num])
    for student in groups[group_num]:
        result.append({
            '组别': group_num + 1,
            '学号': student['学号'],
            '姓名': student['姓名'],
            '班级': student['班级']
        })

# 转换为DataFrame并保存
result_df = pd.DataFrame(result)
result_df.to_excel("分组结果.xlsx", index=False)

print("分组完成，结果已保存到 分组结果.xlsx")