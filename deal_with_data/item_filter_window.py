import tkinter as tk
import threading
from item_filter import ItemFilter

# 在windows上打包参考https://blog.csdn.net/swing123lovecoding/article/details/118998263
# pyinstall在pycharm.ui无法下载，请在pycharm.terminal中命令行处理
# 打包请在terminal中用命令行打包，item_filter_window.py必须放到第一位，否则exe运行不会打开界面
# pyinstaller -F -w item_filter_window.py item_filter.py
# 测试：X86企业级，X86终端
class App(object):
    def __init__(self, window):
        # http://c.biancheng.net/tkinter/layout-manager-method.html
        self.window = window

        # filter notice
        self.label = tk.Label(window, text='过滤词语，多个词语用【中文】逗号间隔', width=40, font=('Arial', 14)).grid(row=0, column=0,
                                                                                                      columnspan=2)

        # filter
        self.var_filter_string = tk.StringVar()
        entry_filter_string = tk.Entry(window, textvariable=self.var_filter_string, font=('Arial', 14))
        entry_filter_string.grid(row=1, column=0, sticky='W')

        # IntVar() 用于处理整数类型的变量
        self.var_categroy = tk.IntVar() # 0是创建多个表；1是创建一个表；2未来是给"合并表格使用的"
        # 根据单选按钮的 value 值来选择相应的选项
        # self.var_categroy.set(1)
        # 使用 variable 参数来关联 IntVar() 的变量 v
        tk.Radiobutton(window, text="创建一个表", variable=self.var_categroy, value=1).grid(row=2, column=0)
        tk.Radiobutton(window, text="创建多个表", variable=self.var_categroy, value=2).grid(row=3, column=0)
        # tk.Radiobutton(window, text="创建多个表", variable=self.var_categroy, value=3).grid(row=3, column=0)

        # 启动按钮
        self.start_button = tk.Button(window, text='开始', font=('Arial', 10),
                                      command=self.prepare_data_for_start_filter)
        self.start_button.grid(row=6, column=1, sticky='W')

    def prepare_data_for_start_filter(self):
        # check
        filter_string = self.var_filter_string.get()
        if len(filter_string) == 0:
            filter_string.messagebox.showerror(title='出错了', message='过滤词语不能为空')
            return

        # 中文逗号
        items: [] = filter_string.split('，')

        if len(items) == 0:
            items.messagebox.showerror(title='出错了', message='过滤用中文逗号切割出错')
            return
        only_one_result_workbook = False
        if self.var_categroy.get() == 1:
            only_one_result_workbook = True

        t1 = threading.Thread(target=self.start_filter, args=(items, only_one_result_workbook))
        t1.start()

        self.start_button.config(text="执行中")
        self.start_button.config(state='disabled')

    def start_filter(self, items, only_one_result_workbook):
        item_filter: ItemFilter = ItemFilter()
        item_filter.ds_sheet_target_items = items
        item_filter.only_one_result_workbook = only_one_result_workbook
        for temp in items:
            print('app要过滤的词语:' + temp)

        print('app要过滤only_one_result_workbook=' + str(only_one_result_workbook))

        if item_filter.only_one_result_workbook:
            print('app 进入只创建一个表格')
            # 只有1个表格
            item_filter.create_result_sheet()
            item_filter.filter_all_target_item_to_one_result_sheet()
            item_filter.hidden_columns()
            items_name = '_'.join(item_filter.ds_sheet_target_items)
            print('items_name = ' + items_name)
            item_filter.save_result_book(items_name)
        else:
            print('app 进入创建n个表格')
            # 每个过滤词对应一个表格
            for item in item_filter.ds_sheet_target_items:
                item_filter.create_result_sheet()
                item_filter.filter_a_item_to_one_result_sheet(item)
                item_filter.hidden_columns()
                item_filter.save_result_book(item)

        self.start_button.config(text="完毕")
        self.start_button.config(state='disabled')

# init
window = tk.Tk()
window.title('0.0.1')
# 设定窗口的大小(长 * 宽)
window.geometry('450x600')
window.resizable(0, 0)
app = App(window)
# 主窗口循环显示
window.mainloop()
