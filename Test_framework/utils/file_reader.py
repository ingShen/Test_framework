'''
    1.为了读取yaml文件，我们需要一个封装YamlReader类
    2.数据分离，进行参数化，添加ExcelReader类
'''
import yaml
import os
from xlrd import open_workbook

class YamlReader:
    def __init__(self, yamlf):
        if os.path.exists(yamlf):
            self.yamlf = yamlf
        else:
            raise FileNotFoundError('文件不存在！')
        self._data = None

    @property
    def data(self):
        # 如果是第一次调用data，读取yaml文档，否则直接返回之前保存的数据
        if not self._data:
            with open(self.yamlf, 'rb') as f:
                self._data = list(yaml.safe_load_all(f)) # load后是一個generator,用list組成列表
        return self._data

class SheetTypeError(Exception):
    pass


'''
读取excel文件中内容。返回list
如：
excel中内容为：
| A  | B  | C  |
| A1 | B1 | C1 |
| A2 | B2 | C2 |

如果 print(ExcelReader(excel, title_line=True).data)，输出结果：
[{A: A1, B: B1, C:C1}, {A:A2, B:B2, C:C2}]

如果 print(ExcelReader(excel, title_line=False).data)，输出结果：
[[A,B,C], [A1,B1,C1], [A2,B2,C2]]

可以指定sheet，通过index或者name：
ExcelReader(excel, sheet=2)
ExcelReader(excel, sheet='BaiDuTest')
'''
class ExcelReader:
    def __init__(self, excel, sheet=0, title_line=True):
        if os.path.exists(excel):
            self.excel = excel
        else:
            raise FileNotFoundError('文件不存在！')
        self.sheet = sheet
        self.title_line = title_line
        self._data = list()

    @property
    def data(self):
        if not self._data:
            workbook = open_workbook(self.excel)
            if type(self.sheet) not in [int, str]:
                raise SheetTypeError('Please pass in <type int> or <type str>, not {0}'.format(type(self.sheet)))
            elif type(self.sheet) == int:
                s = workbook.sheet_by_index(self.sheet)
            else:
                s = workbook.sheet_by_name(self.sheet)
            
            if self.title_line:
                title = s.row_values(0) # 首行为title
                for col in range(1, s.nrows):
                    # 依次遍历其余行，与首行组成dict，拼到self._data中
                    self._data.append(dict(zip(title, s.row_values(col))))
            else:
                for col in range(0, s.nrows):
                    # 遍历所有行，拼到self._data中
                    self._data.append(s.row_values(col))
        return self._data

if __name__ == '__main__':
    y = 'D:\xxx\Code\Test_framework\config\config.yml' # 改成自己的路径
    reader = YamlReader(y)
    print(reader.data)

    e = 'D:\xxx\doc\Code\Test_framework\data\\baidu.xlsx' # 改成自己的路径
    reader = ExcelReader(e, title_line=True)
    print(reader.data)