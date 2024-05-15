import tkinter as tk
from tkinter import filedialog, messagebox
import os
import subprocess
from xmind2xlsopen import transformer

def get_file_extension(file_path):
    # 使用os.path.splitext()分割文件名和扩展名
    base, ext = os.path.splitext(file_path)
    return ext if ext.startswith('.') else None

def select_xmind_file():
    filetypes = (('Excel files', '*.xmind'),('All files', '*.*'),)  # 可选，添加一个所有文件的选项)
    file_path = filedialog.askopenfilename(title="请选择xmind文件",filetypes=filetypes)
    xmind_file_path.set(file_path)
    xmind_path_entry.delete(0, tk.END)  # 清除原有路径
    if get_file_extension(file_path)!='.xmind':
        messagebox.showinfo("警告", "请选择xmind文件")
        return
    xmind_path_entry.insert(0, os.path.normpath(file_path))  # 插入新路径

def select_output_dir():
    dir_path = filedialog.askdirectory()
    output_dir_path.set(dir_path)
    excel_dir_entry.delete(0, tk.END)  # 清除原有目录
    excel_dir_entry.insert(0, os.path.normpath(dir_path))  # 插入新目录

def convert_xmind():
    xmind_path = xmind_file_path.get()
    xmind_path= os.path.normpath(xmind_path)
    output_path = output_dir_path.get()
    output_path = os.path.normpath(output_path)

    # print("xmind路径是： ",xmind_path," excel输出路径： ",output_path)
    #不能为空
    if xmind_path=="" or xmind_path=="." or output_path=="" or output_path==".":
        messagebox.showerror("错误", "请先选择XMind文件和输出目录")
        return
    #使用内嵌属性__file__获取文件路径，实现自定义功能
    #将相对路径替换为生成路径
    pyFileName=get_directory_path(__file__) + "\\" + "xmind2xlsopen.py"

    # runMsg= pyFileName+" xmind path " + xmind_path+" excel path " + output_path
    runMsgStr=pyFileName+" "+xmind_path+" "+output_path
    # print("runMsg完整命令 ", runMsg)
    # commandStr=r"python C:\Users\Administrator\PycharmProjects\xmind2xlsopen\xmind2xlsopen.py D:\WEB_0430.xmind D:"
    commandEXE="python "+runMsgStr
    print("组合的命令行字符串是-----》", commandEXE)
    # 这里调用转换逻辑Python脚本
    # try:
    #     #处理这行的异常报错，是程序运行的关键点
    #     # subprocess.run([r"python", pyFileName, xmind_path, output_path], check=True)
    #     subprocess.run(commandEXE.split(), check=True)
    #
    #     messagebox.showinfo("完成", "转换完成")
    # except subprocess.CalledProcessError as e:
    #     messagebox.showinfo("错误", runMsgStr)
    #     messagebox.showerror("错误", f"转换过程中发生错误: {e}")

    transformer(xmind_path, output_path)
    messagebox.showinfo("完成", "转换完成")

def get_directory_path(file_path):
    # split()方法根据指定的分隔符将字符串分割成列表，
    # os.sep来确保跨平台兼容（os.sep在Windows上为'\', 在Unix/Linux上为'/'）
    parts = file_path.split(os.sep)
    # 去掉最后一个元素（即文件名），然后使用os.sep连接剩下的部分
    directory_path = os.sep.join(parts[:-1])
    return directory_path

# 创建主窗口
root = tk.Tk()
root.title("XMind转换工具")

# 存储XMind文件路径和输出目录路径的变量
xmind_file_path = tk.StringVar()
output_dir_path = tk.StringVar()

# 第一行布局
tk.Label(root, text="请选择xmind文件:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
xmind_path_entry = tk.Entry(root, width=50)
xmind_path_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text=" 浏览 ", width=10,command=select_xmind_file).grid(row=0, column=2, padx=10, pady=10)

# 第二行布局
tk.Label(root, text="请选择输出excel文件的路径:").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
excel_dir_entry = tk.Entry(root, width=50)
excel_dir_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="选择目录", width=10,command=select_output_dir).grid(row=1, column=2, padx=10, pady=10)

# 第三行布局，转换按钮
convert_button = tk.Button(root, text="开始转换", width=20,bg="white",command=convert_xmind)
convert_button.grid(row=2, column=1, columnspan=2, padx=10, pady=20)

# 运行GUI
root.mainloop()


# # 创建选择XMind文件的按钮
# tk.Button(root, text="选择XMind文件", command=select_xmind_file).pack()
#
# # 显示XMind文件路径
# tk.Label(root, textvariable=xmind_file_path).pack()
# # tk.Text(root,)
# # 创建选择输出目录的按钮
# tk.Button(root, text="选择输出目录", command=select_output_dir).pack()
#
# # 显示输出目录路径
# tk.Label(root, textvariable=output_dir_path).pack()
#
# # 创建转换按钮
