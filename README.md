# excel2pdfconverter
convert excel to pdf 【xls与xlsx批量转化为pdf的东西】

- 原理：通过pywin32模块调用Windows API，批量导出同一文件夹下的.xls/.xlsx文件的pdf版本。

- 注意：需要先安装pywin32，按照提示进行操作。

- 打包为exe：需要安装pyinstaller，在代码目录执行

>pyinstaller -F -n GD_excel2pdf -i icon.ico .\excel2pdf.py

其中-n和-i后面的内容可自定义。


另外，freeze.py是给py2exe用的，无法只生成一个exe文件，不如pyinstaller好用。

