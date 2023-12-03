# excel2pdfconverter
walk through directory iteratively, batch convert excel to pdf.

- simple logic：using pywin32 to call Windows API directly。

- atention: need `pywin32(pip install pywin32)`.

- transfer to exe file: need `pyinstaller(pip install pyinstaller)`，run `pyinstaller` in the code directory:

>pyinstaller -F -n GD_excel2pdf -i icon.ico .\excel2pdf.py

the content after `-n` and `-i` can modified for your own demand.


besids, `freeze.py` is the builder for py2exe, which can not transfer the python code to only ONE exe file. so I choose `pyinstaller`. 

