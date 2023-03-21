import tkinter as tk
from tkinter import filedialog
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from pandastable import Table, TableModel
import tkinter.messagebox as messagebox

class ExcelProcessor:
    try:
        def __init__(self, master):
            self.master = master
            self.master.title("培养方案完成情况分析")
            self.master.geometry("500x400")

            # Create widgets
            self.file_label1 = tk.Label(self.master, text="请上传当前学生培养方案")
            self.file_label1.pack(pady=10)

            self.file_button1 = tk.Button(self.master, text="培养方案上传", command=self.select_file1)
            self.file_button1.pack(pady=10)

            self.file_label2 = tk.Label(self.master, text="请上传当前学生成绩单")
            self.file_label2.pack(pady=10)

            self.file_button2 = tk.Button(self.master, text="成绩单上传", command=self.select_file2)
            self.file_button2.pack(pady=10)

            self.process_button = tk.Button(self.master, text="进行分析", command=self.process_files,
                                            state=tk.DISABLED)
            self.process_button.pack(pady=10)

            self.result_text = tk.Text(self.master)
            self.result_text.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        def select_file1(self):
            file_path = filedialog.askopenfilename()
            if file_path:
                self.file_label1.config(text=file_path)
                if self.file_label2.cget("text") != "No file selected":
                    self.process_button.config(state=tk.NORMAL)

        def select_file2(self):
            file_path = filedialog.askopenfilename()
            if file_path:
                self.file_label2.config(text=file_path)
                if self.file_label1.cget("text") != "No file selected":
                    self.process_button.config(state=tk.NORMAL)

        def display_table(self, dataframe, title):
            # Create a new window to display the data
            table_window = tk.Toplevel(self.master)
            table_window.title(title)

            # Create a Table widget to display the data
            table = Table(table_window, dataframe=dataframe, showtoolbar=True, showstatusbar=True)

            # Add a custom button
            def custom_button_callback():
                # Your custom button callback code here
                pass

            table.add_button('Custom Button', custom_button_callback)
            # Convert the image to a Tkinter-compatible format
            # tk_image = ImageTk.PhotoImage(image)

            # Add the button with the image
            # table.add_button(tk_image, custom_button_callback, row=0, column=0)
            table.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        def process_files(self):
            # Read first Excel file
            file_path = self.file_label1.cget("text")
            df_plan = pd.read_excel(file_path, sheet_name='Sheet1', header=2)
            end_index = df_plan[df_plan['课程类型'] == '学生签名：'].index[0]
            df_plan = df_plan.drop(df_plan.index[end_index:])
            df_plan['课程编码'] = df_plan['课程编码'].fillna(-1).astype(int)
            result = df_plan[['课程名称', '课程编码', '学分', '课程类型']]
            df_plan2 = pd.read_excel(file_path, sheet_name='Sheet1')
            column_name = df_plan2.columns[0]
            new_column_name = column_name.replace(" ", "") + "完成情况"

            # Read second Excel file
            file_path = self.file_label2.cget("text")
            df_completed = pd.read_excel(file_path, header=2)
            df = pd.read_excel(file_path)
            major = df.iloc[2, 8]
            class_name = df.iloc[2, 9]

            new_df = df_completed[['课程名称', '课程编号', '学分', '课程属性']]
            new_df.columns = ['课程名称', '课程编码', '学分', '课程类型']
            course_map = {"必修": "必修课程", "选修": "选修课程"}
            # new_df['课程类型'] = new_df['课程类型'].map(course_map)
            new_df['课程类型'] = new_df['课程类型'].map(course_map)
            new_df.loc[:, '课程类型'] = new_df.loc[:, '课程类型'].map(course_map)
            merged_df = pd.merge(result, new_df, on=["课程名称"], how="left", indicator=True)
            # 找出不存在于 new_df 中的课程
            not_in_new_df = merged_df.loc[
                merged_df["_merge"] == "left_only", ["课程名称", "课程编码_x", "学分_x", "课程类型_x"]]
            not_in_new_df.columns = ['课程名称', '课程编码', '学分', '课程类型']
            # 根据课程类型分别存放到三个 DataFrame 中
            # 找到new_df中在result_df中不存在的行
            new_courses = new_df[~new_df['课程名称'].isin(result['课程名称'])]
            required_courses = not_in_new_df.loc[not_in_new_df["课程类型"] == "必修课程"]
            elective_courses = not_in_new_df.loc[not_in_new_df["课程类型"] == "选修课程"]
            practice_courses = not_in_new_df.loc[not_in_new_df["课程类型"] == "实践课程"]
            required_courses['学分'] = required_courses['学分'].astype(float)
            elective_courses['学分'] = elective_courses["学分"].astype(float)
            practice_courses['学分'] = practice_courses["学分"].astype(float)
            credits_sum = required_courses['学分'].sum()
            credits_sum_2 = elective_courses['学分'].sum()
            credits_sum_3 = practice_courses['学分'].sum()
            # 计算每种课程类型未修读的学分总和
            # credits_sum_untaken = not_in_new_df.groupby('课程类型')['学分'].sum()
            # credits_sum_untaken = credits_sum_untaken.reset_index()
            # credits_sum_untaken.columns = ['课程类型', '未修读学分']
            new_r = pd.DataFrame([[new_column_name, '', '', '']], columns=['课程名称', '课程编码', '学分', '课程类型'])
            new_r1 = pd.DataFrame([['专业', major, '班级', class_name]], columns=['课程名称', '课程编码', '学分', '课程类型'])
            new_row0 = pd.DataFrame([['', '', '', '']], columns=['课程名称', '课程编码', '学分', '课程类型'])
            new_row1 = pd.DataFrame([['必修课程未修读', "合计"+str(credits_sum), '', '']],
                                    columns=['课程名称', '课程编码', '学分', '课程类型'])
            new_row2 = pd.DataFrame([['选修课程未修读', "合计"+str(credits_sum_2), '', '']],
                                    columns=['课程名称', '课程编码', '学分', '课程类型'])
            new_row3 = pd.DataFrame([['实践课程未修读', "合计"+str(credits_sum_3), '', '']],
                                    columns=['课程名称', '课程编码', '学分', '课程类型'])
            new_row = pd.DataFrame([['', '', '', ''],['不在培养方案但是已经修读', '', '', ''] ],
                                   columns=['课程名称', '课程编码', '学分', '课程类型'])
            new_courses = pd.concat([new_row, new_courses])
            # 合并 DataFrame
            concatenated_df = pd.concat(
                [new_r,new_r1,new_row1, required_courses, new_row0, new_row2, elective_courses, new_row0, new_row3, practice_courses,
                 new_row0])
            # # 将三个 DataFrame 拼接在一起
            #
            # # 重置索引，使得索引按照顺序排列
            concatenated_df = concatenated_df.reset_index(drop=True)
            # 将三个学分总和添加到 DataFrame 的末尾
            concatenated_df = concatenated_df.append({'学分': credits_sum, '课程类型': '必修课程未修读合计'},
                                                     ignore_index=True)
            concatenated_df = concatenated_df.append({'学分': credits_sum_2, '课程类型': '选修课程未修读合计'},
                                                     ignore_index=True)
            concatenated_df = concatenated_df.append({'学分': credits_sum_3, '课程类型': '实践课程未修读合计'},
                                                     ignore_index=True)

            concatenated_df = pd.concat([concatenated_df, new_courses])
            # Create a new window to display the data
            data_window = tk.Toplevel(self.master)
            data_window.title(

            )

            # Create a Table widget to display the data
            table = Table(data_window, dataframe=concatenated_df, showtoolbar=True, showstatusbar=True)
            table.show()
    except Exception as e:
            messagebox.showerror("Error", "Error processing files: {}".format(e))
# 创建主窗口
root = tk.Tk()
excel_processor = ExcelProcessor(root)

# Call mainloop after creating and populating result and new_df windows
root.mainloop()