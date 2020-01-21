# merge-xlsx README

## Features

Combine two Excel files to one (order by the first file's first column).
Press`F1`or`ctrl+shift+p`then use command`merge xlsx`, then enter the first file's name, then the second file's name.
The two files' first row & fist column's content must be the same (eg: `name` in the tables below) or both be empty, other column's first row must has content too (eg: `age` and `score` below).

将两个Excel表格合并成一个（根据第一个表格的第一列的顺序），如：
按`F1`或`ctrl+shift+p`后输入命令`merge xlsx`根据提示输入当前工作区根目录第一个文件名和第二个文件名
两个文件的第一行第一列内容必须相同（或都置空），比如下方都是`name`，其余列的第一行也要有表头说明，如下方的`age`和`score`

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

result.xlsx:

| name | age  | score |
| ---- | ---- | ----- |
| Jack | 20   | 80    |
| Tom  | 30   | 90    |
| John | 40   | 70    |
| Jame |      | 60    |

