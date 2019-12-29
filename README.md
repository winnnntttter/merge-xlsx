# merge-xlsx README

## Features

将两个Excel表格合并成一个（根据第一个表格的第一列），如：

a.xlsx:

| name | age  |
| ---- | ---- |
| Jack | 20   |
| Tom  | 30   |
| John | 40   |

b.xlsx:

| name | score |
| ---- | ----- |
| Tom  | 90    |
| Jack | 80    |
| John | 70    |
| Jame | 60    |

合并成：

result.xlsx:

| name | age  | score |
| ---- | ---- | ----- |
| Jack | 20   | 80    |
| Tom  | 30   | 90    |
| John | 40   | 70    |
| Jame |      | 60    |

按`F1`或`ctrl+shift+p`后输入命令`merge xlsx`根据提示输入当前工作区根目录第一个文件名和第二个文件名