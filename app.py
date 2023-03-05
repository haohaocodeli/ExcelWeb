import pandas as pd
from flask import Flask, render_template, request, send_file

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
    # 将处理后的数据保存为Excel文件，并将其作为响应返回给前端页面
    writer = pd.ExcelWriter('result.xlsx')
    result.to_excel(writer, sheet_name='表格1', index=False)
    new_df.to_excel(writer, sheet_name='表格2', index=False)
    # return send_file('result.xlsx', as_attachment=True)
    return render_template('index.html', tables=[result.to_html(classes='data', header="true"),
                                                 new_df.to_html(classes='data', header="true")])
if __name__ == '__main__':
    app.run()
