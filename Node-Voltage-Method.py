from numpy import array, reciprocal, linalg
from tkinter.filedialog import askopenfilename
import tkinter as tk
# import SchemDraw
import xlrd


def fun(address):
    # 从excel中获取数据
    global cf, rf, ced
    wb = xlrd.open_workbook(address)  # 打开excel文件
    sheet = wb.sheet_by_index(0)  # 选取第一页
    row = sheet.nrows  # 保存最大行数、列数，为之后提供范围
    col = sheet.ncols

    # print(test)
    bp = False  # 用于break退出双重循环
    for i in range(row):  # 用r,R,g,G表头寻找表格
        for k in range(col):
            if sheet.cell(i, k).value in ['r', 'R', 'g', 'G']:
                print('表头在(', i, ',', k, ')')
                rf = i
                cf = k
                bp = True
                break
        if bp:  # 退出第二重循环
            break

    i = rf + 1  # 单元格位置坐标，从表头右下角一个单元格开始
    k = cf + 1
    ap = []  # 用于存储系数的表格
    tem = []  # 用于存储每行数字
    v = sheet.cell(i, k).value  # 取特定单元格的值
    while isinstance(v, float):  # 假如是数字，才读入
        while isinstance(v, float):
            tem.append(v)  # 将新单元格的值添加到tem的末尾
            k += 1
            if k >= col:  # 假如此表格就在页面的右下角，则需要避免使用value读取超过范围的单元格，所以用if退出循环
                break
            v = sheet.cell(i, k).value  # 用于下次while判断
        # print('tem=', tem)
        ap.append(tem)  # 将本行数据tem添加到a中
        tem = []  # tem置空，用于下一行
        # print('a=',a)
        i += 1
        ced = k
        if i >= col:
            break
        k = cf + 1
        v = sheet.cell(i, k).value
    # print('a=',a)
    a = array(ap, dtype=float)  # 将多维列表转化为numpy库下的矩阵
    if sheet.cell(rf, cf).value in ['r', 'R']:  # 取倒数，将电阻转化为电导
        a = reciprocal(a)
    print('A=', a)
    bp = []
    for sb in range(rf + 1, i):  # 从上到下读取电流源的值，并存储在列矩阵b中
        bp.append([sheet.cell(sb, ced + 1).value])
    b = array(bp, dtype=float)
    print('B=', b)
    # 计算结点电压
    x = linalg.solve(a, b)  # 求解线性方程组
    print('x=', x)
    lt.insert(tk.END, '各节点电压为：')
    for it in x:
        lt.insert(tk.END, it[0])


# 绘出电路图（可选）


# 使用交互界面展示结果
def select_path():
    # 选择文件path_接收文件地址
    path_ = tk.filedialog.askopenfilename()

    # 通过replace函数替换绝对文件地址中的/来使文件可被程序读取
    # 注意：\\转义后为\，所以\\\\转义后为\\
    path_ = path_.replace("/", "\\")
    path.set(path_)     # 将字符串path_赋值给path，在输入框中就会更新path的新值
    fun(path_)


# 建立窗体
win = tk.Tk()
win.title('节点电压法')
win.geometry('400x300')
path = tk.StringVar()

# 生成标签、输入框、按钮和复选栏
null = tk.Label(win).grid(row=0, column=0, padx=30)     # 空元件，用于调整主要元件的位置
et = tk.Entry(win, textvariable=path)
et.grid(row=0, column=2)
tk.Label(win, text="目标路径:").grid(row=0, column=1)
tk.Button(win, text="路径选择", command=select_path).grid(row=0, column=3)
lt = tk.Listbox(win)
lt.grid(row=1, column=2)

win.mainloop()  # 不断刷新窗体，更新内容并保持窗体存在

# 试验场
# 不管是type还是isinstance，都不能用complex来识别所有数字
# python变量名区分大小写
# print('here')
# str(500)将数字500转化为字符串‘500’
# input()
