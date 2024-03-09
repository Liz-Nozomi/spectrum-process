import pandas as pd
import glob
import sys
import datetime
# 获取所有以"BF2-"开头并以"_processed"结尾的Excel文件列表
excel_files = glob.glob("BF2-*_processed.xlsx")  # 替换为你的文件路径

# 创建一个空的DataFrame，用于存储合并后的数据
merged_data = pd.DataFrame()
header_row=['时间','间隔','异戊烷','23DMB','2MP','3MP','24DMP','223TMB','2MH','23DMP','3MH','224TMP','25DMH','223TMP','24DMH','234TMP','233TMP','33DMH','23DMH','225TMH']

# 逐个读取Excel文件并合并到DataFrame中
for file in excel_files:
    df = pd.read_excel(file, header=None, skiprows=1, nrows=1)
    merged_data = pd.concat([merged_data, df], ignore_index=True)

# 按照第一列的值进行排序

merged_data = merged_data.sort_values(by=int(str(merged_data.columns[0]).replace(':','')))
merged_data= pd.concat([pd.DataFrame([header_row], columns=merged_data.columns), merged_data], ignore_index=True)

# 将合并并排序后的数据保存到新的Excel文件
merged_data.to_excel('合并后数据.xlsx', index=False, header=False)
sys.exit()