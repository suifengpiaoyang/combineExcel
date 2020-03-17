import os
import sys
import time
import openpyxl
import traceback

if not os.path.exists('setting.py'):
    with open('setting.py', 'w') as fl:
        fl.write("FILE_SAVE_PATH = r''\n")
        fl.write("NEED_CONBINE_DICTORY_PATH = r''\n")
        fl.write("MODE = 1\n")
    print('''
因第一次运行，请先填好 setting.py 文件。
可以使用记事本打开。请在引号中填写字段。

FILE_SAVE_PATH 填写合并完后文件保存路径；

NEED_CONBINE_DICTORY_PATH 填写需要合并的文件夹路径；
如果都不填写，则默认合并脚本路径下的后缀名为 .xlsx 的
文件，并且生成的文件也保存在该路径下。后续生成的文件不
会进入合并的列表中。

MODE 表示模式选择，目前只有 0 和 1 两种模式：
模式 0 : 普通合并所有 Excel 文件；
模式 1 : 合并所有 Excel 文件，从第二个文件开始不合并
         文件的第一行，因为有很多情况下这些 Excel 文
         件都是同样形式；

配置完成后，请重新运行本程序。
'''
          )
    os.system('pause')
    sys.exit()

from setting import FILE_SAVE_PATH, NEED_CONBINE_DICTORY_PATH, MODE

if MODE not in (0,1):
    print('请在 setting.py 中设置正确的 MODE 值，')
    print('目前只支持 0 和 1,现在程序停止。')
    os.system('pause')
    sys.exit()

class CombineExcelFiles:

    def __init__(self):

        self.current_time = time.strftime('%Y-%m-%d-%H-%M-%S')
        self.default_path = os.getcwd()
        self.filename = '合并完成({}).xlsx'.format(self.current_time)
        if len(FILE_SAVE_PATH) == 0:
            self.file_path = os.path.join(self.default_path, self.filename)
        else:
            self.file_path = os.path.join(FILE_SAVE_PATH, self.filename)
        if len(NEED_CONBINE_DICTORY_PATH) == 0:
            self.target_path = self.default_path
        else:
            self.target_path = NEED_CONBINE_DICTORY_PATH

    def run(self, mode=MODE):
        try:
            self.createExcel()
            self.combine(mode)
        except Exception as e:
            traceback.print_exc()
            os.system('pause')
        finally:
            self.close()

    def createExcel(self):
        self.wb = openpyxl.Workbook()
        self.ws = self.wb.active
        self.max_row = 0
        self.max_column = 0

    def combine(self, mode=MODE):

        # 不同的 mode 代表着不同的表格合并策略
        # 默认 0 代表着全合并
        # 数字 1 代表着除了第一个表格，其他表格忽略掉第一行
        target_file_list = os.listdir(self.target_path)
        for line in target_file_list:
            if '.' in line and '合并' not in line:
                head, tail = line.rsplit('.', 1)
                if tail == 'xlsx':
                    print(line)
                    abs_path = os.path.join(self.target_path, line)
                    wb = openpyxl.load_workbook(abs_path)
                    ws = wb.active
                    # 得到目标表格最大的行列数
                    max_row = ws.max_row
                    max_column = ws.max_column
                    if mode == 0:
                        for row in range(1, max_row + 1):
                            for column in range(1, max_column + 1):
                                self.ws.cell(
                                    row=self.max_row + 1, column=column).value = ws.cell(row=row, column=column).value
                            self.max_row += 1
                    elif mode == 1:
                        if self.max_row == 0:
                            for row in range(1, max_row + 1):
                                for column in range(1, max_column + 1):
                                    self.ws.cell(
                                        row=self.max_row + 1, column=column).value = ws.cell(row=row, column=column).value
                                self.max_row += 1
                        else:
                            for row in range(2, max_row + 1):
                                for column in range(1, max_column + 1):
                                    self.ws.cell(
                                        row=self.max_row + 1, column=column).value = ws.cell(row=row, column=column).value
                                self.max_row += 1
                    wb.close()
                else:
                    pass
            else:
                pass

    def close(self):
        self.wb.save(self.file_path)
        self.wb.close()

if __name__ == '__main__':
    # pass
    print('提示：只有后缀名为 xlsx 的才能合并。')
    print('正在合并工作表...\n')

    start_time = time.time()
    c = CombineExcelFiles()
    c.run(mode=MODE)
    end_time = time.time()

    print('\n合并完成。共耗时 {} 秒。\n'.format(end_time - start_time))
    print('目标文件位置为：')
    print(c.file_path)
    print()
    os.system('pause')
