# File-Renaming

## 筛选重命名

可将文件夹中的文件进行重命名排序
排序方式有多种，主要是可以将文件名称前面加上序号

## 匹配列表重命名

可将文件夹中的文件与待定表格中的内容进行匹配重命名
匹配表格中的某一列信息中的名称与文件夹中的文件名进行匹配，匹配达到置信度则重命名文件夹中的文件，将表格中的另一列附加到文件夹中对应文件的文件名前

pip install pandas openpyxl tkinter
pip install PyQt5

python依赖包安装后可直接运行对应的.py文件，也可通过指令进行打包成可执行程序
pyinstaller --onefile --noconsole 筛选重命名.py
pyinstaller --onefile --noconsole 匹配列表重命名.py