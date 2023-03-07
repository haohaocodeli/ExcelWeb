import pandas as pd
from flask import Flask, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
# 创建Flask实例
app = Flask(__name__)
# 路由，用于展示上传表格的页面
@app.route('/')
def index():
    return render_template('upload.html')
# 路由，用于处理上传的表格并进行数据预处理
@app.route('/process', methods=['POST'])
def process():
    # 读取上传的表格1
    file1 = request.files['file1']
    df_plan = pd.read_excel(file1, sheet_name='Sheet1', header=2)
    # 找到第一个出现“学生签名”的位置
    end_index = df_plan[df_plan['课程类型'] == '学生签名：'].index[0]
    # 删除“学生签名”这一行及其之后的行
    df = df_plan.drop(df_plan.index[end_index:])
    # 将“课程编码”列的空值填充为-1，并将其转换为整型
    df['课程编码'] = df['课程编码'].fillna(-1).astype(int)
    # 提取需要的四列数据
    result = df[['课程名称', '课程编码', '学分', '课程类型']]
    # 读取上传的表格2
    file2 = request.files['file2']
    df_completed = pd.read_excel(file2, sheet_name='Sheet1', header=2)
    # 提取需要的四列数据
    new_df = df_completed[['课程名称', '课程编号', '学分', '课程属性']]
    new_df.columns = ['课程名称', '课程编码', '学分', '课程类型']
    course_map = {"必修": "必修课程", "选修": "选修课程"}
    new_df['课程类型'] = new_df['课程类型'].map(course_map)
    # 找到new_df中在result_df中不存在的行
    new_courses= new_df[~new_df['课程名称'].isin(result['课程名称'])]

    # 将处理后的数据保存为Excel文件，并将其作为响应返回给前端页面
    writer = pd.ExcelWriter('result.xlsx')
    result.to_excel(writer, sheet_name='表格1', index=False)
    new_df.to_excel(writer, sheet_name='表格2', index=False)
    writer.save()
    # 合并两个 DataFrame
    merged_df = pd.merge(result, new_df, on=["课程名称"], how="left", indicator=True)
    # 找出不存在于 new_df 中的课程
    not_in_new_df = merged_df.loc[merged_df["_merge"] == "left_only", ["课程名称", "课程编码_x", "学分_x", "课程类型_x"]]
    not_in_new_df.columns = ['课程名称', '课程编码', '学分', '课程类型']
    # 根据课程类型分别存放到三个 DataFrame 中
    required_courses = not_in_new_df.loc[not_in_new_df["课程类型"] == "必修课程"]
    elective_courses = not_in_new_df.loc[not_in_new_df["课程类型"] == "选修课程"]
    practice_courses = not_in_new_df.loc[not_in_new_df["课程类型"] == "实践课程"]
    # print(required_courses)
    credits_sum=0
    required_courses['学分'] = required_courses['学分'].astype(float)
    elective_courses['学分']=elective_courses["学分"].astype(float)
    practice_courses['学分']=practice_courses["学分"].astype(float)
    credits_sum = required_courses['学分'].sum()
    credits_sum_2=elective_courses['学分'].sum()
    credits_sum_3=practice_courses['学分'].sum()
    # 将所有表格合并成一个 DataFrame
    all_df = pd.concat([result, new_df, required_courses, elective_courses, practice_courses], axis=0,
                       ignore_index=True)
    writer = pd.ExcelWriter('课程修读情况.xlsx')
    # return send_file('课程修读情况.xlsx', as_attachment=True)
    # 将数据写入 Excel 文件中
    with pd.ExcelWriter('课程修读情况.xlsx', engine='openpyxl') as writer:
        # 写入合并后的 DataFrame
        result.to_excel(writer, sheet_name='培养方案', index=False)
        new_df.to_excel(writer, sheet_name='已修读所有课程', index=False)
        # 创建一个新的 sheet 来展示 credits_sum
        wb = writer.book
        ws = wb.create_sheet('Credits Sum', 1)
        ws.append(['必修课未修读', credits_sum])
        for r in dataframe_to_rows(required_courses, index=False, header=True):
            ws.append(r)
        ws.append(['选修课程未修读', credits_sum_2])
        for r in dataframe_to_rows(elective_courses, index=False, header=True):
            ws.append(r)
        ws.append(['实践课程未修读', credits_sum_3])
        for r in dataframe_to_rows(practice_courses, index=False, header=True):
            ws.append(r)
        ws.append(['已修读但是未在培养方案课程'])
        for r in dataframe_to_rows(new_courses, index=False, header=True):
            ws.append(r)
        # return send_file('课程修读情况.xlsx', as_attachment=True)
    return render_template('index.html', tables=[result.to_html(classes='data', header="true"),
                                                 new_df.to_html(classes='data', header="true"),required_courses.to_html(classes='data', header="true"),elective_courses.to_html(classes='data', header="true"),practice_courses.to_html(classes='data', header="true"),new_courses.to_html(classes='data', header="true") ]
   ,credits_sum = credits_sum,credits_sum_2=credits_sum_2,credits_sum_3=credits_sum_3)
if __name__ == '__main__':
    app.run()
