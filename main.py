import xlrd
from func_lib import *

root_path=r'C:\Users\18005\Desktop\0511\input\og.xlsx'
opt_dir=r'C:\Users\18005\Desktop\0511\output\opt2-3.xlsx'

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    data = xlrd.open_workbook(root_path).sheet_by_index(5)
    # data=pd.read_excel(root_path,sheet_name='月均值需打散')
    # tmp_data=data[:10]
    a = build_init_xlbook(data, opt_dir)
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
