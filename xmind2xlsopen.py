'''
xmind测试设计用例转xls用例
'''
from typing import List, Any
import xlwt
from xmindparser import xmind_to_dict
import sys

def resolve_path(dict_, lists, title):
    """
    通过递归取出每个主分支下的所有小分支并将其作为一个列表
    :param dict_:
    :param lists:
    :param title:
    """
    # 去除title首尾空格
    title = title.strip()
    # 若title为空，则直接取value
    if len(title) == 0:
        concat_title = dict_["title"].strip()
    else:
        concat_title = title + "\t" + dict_["title"].strip()
    if not dict_.__contains__("topics"):
        lists.append(concat_title)
    else:
        for d in dict_["topics"]:
            resolve_path(d, lists, concat_title)

def xmind_to_excel(list_,excel_path):
    f = xlwt.Workbook()
    # 生成单sheet的Excel文件，sheet名自取
    sheet = f.add_sheet("登录模块", cell_overwrite_ok=True)

    # 第一行固定的表头标题，严重程度
    row_header = ["序号", "模块", "功能点", "测试描述", "预置条件", "测试步骤","预期结果", "严重程度","实测结果"]
    for i in range(0,len(row_header)):
        sheet.write(0,i,row_header[i])

    # 增量索引
    index = 0

    for h in range(0, len(list_)):
        lists: List[Any] = []
        resolve_path(list_[h], lists, "")
        print(lists)
        print('\n'.join(lists))  # 主分支下的小分支

        for j in range(0, len(lists)):
            # 将主分支下的小分支构成列表
            lists[j] = lists[j].split('\t')
            # print(lists[j])

            for n in range(0, len(lists[j])):
                # 生成第一列的序号
                sheet.write(j + index + 1, 0, j + index + 1)
                sheet.write(j + index + 1, n + 1, lists[j][n])
                # 自定义内容，比如：测试点/用例标题、预期结果、实际结果、操作步骤、严重程度……
                # 这里为了更加灵活，除序号、模块、功能点的标题固定，其余以【自定义+序号】命名，如：自定义1，需生成Excel表格后手动修改
                if n >= 8:
                    sheet.write(0, n + 1, "自定义" + str(n - 1))

            # 遍历完lists并给增量索引赋值，跳出for j循环，开始for h循环
            if j == len(lists) - 1:
                index += len(lists)
    f.save(excel_path)



def transformer(xmind_path,excel_path):
    # 将XMind转化成字典
    xmind_dict = xmind_to_dict(xmind_path)
    # print("将XMind中所有内容提取出来并转换成列表：", xmind_dict)
    # Excel文件与XMind文件保存在同一目录下
    #@zhu:为保障保存的excel文件格式兼容性，这里选择保存为.xls格式
    excel_name = xmind_path.split('\\')[-1].split(".")[0] + ".xls"
    #如果用户没有配置输出路径，则放在与xmind同一目录下

    if excel_path=='':
        excel_path = "\\".join(xmind_path.split('\\')[:-1]) + "\\" + excel_name
    else:
        excel_path = excel_path + "\\" + excel_name
    print(excel_path)
    # print("通过切片得到所有分支的内容：", xmind_dict[0]['topic']['title'],"\n","最终结果是：===》 ",xmind_dict[0]['topic']['topics'])
    xmind_to_excel(xmind_dict[0]['topic']['topics'],excel_path)

# # 创建解析器
# parser = argparse.ArgumentParser(description='这是一个示例脚本，展示如何使用命令行参数。')
#
# # 添加参数
# parser.add_argument('arg1', help='第一个参数')
# parser.add_argument('arg2', help='第二个参数')
# try:
#     xmind_path_ = sys.argv[1]
#     print("xmind path is: ", len(sys.argv), "  ", xmind_path_)
#     excel_out_path= sys.argv[2]
#     print("excel_out_path path is: ", excel_out_path)
#     transformer(xmind_path_, excel_out_path)
# except IndexError:
#     print("Error: Please provide the path to the XMind file as a command line argument.")
#     print("Usage: python xmind2xlsopen.py 'path_to_xmind_file'")
#     # 可以根据需要设置一个默认值或者退出程序
#     sys.exit(1)

# 解析参数
# args = parser.parse_args()

# if __name__ == '__main__':
#     xmind_path_ = arg1#r"D:\WEB_0430.xmind"
#     excel_out_path= arg2#r"D:"
#     transformer(xmind_path_,excel_out_path)


'''
一个简单的创建excel代码
def write2Excel(INFO, excel_path,sheet_name):
    keys ＝ list(INFO.keys()) ＃获取字典的key值     
    values ＝ list(INFO.values()) ＃获取字典的value值  
    
    #print(keys)
    # print(values)
    book= xlwt.Workbook()
    sheet-book.add_sheet(sheet_name) 
    ＃写入列名
    for i in range(0, len(keys)):
        sheet.write(0,i, keys[i]) 
    ＃写入数据
    for j in range(θ, len(values)):
        sheet.write(1,j, values[j]) 
    book.save(excel_path)

＃执行时
write2Excel(INFO,'./report/excel.xls','sheetname') 
'''