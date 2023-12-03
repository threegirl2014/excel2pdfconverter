#!Python 3

from py2exe import freeze

freeze(
    console=[{"script":"excel2pdf.py"}],
    windows=[],
    data_files=None,
    zipfile=None,
    options={"includes":["os", "sys", "tkinter", "win32com","pywintypes"], "compressed":1, "bundle_files":3, "optimize":2},
    version_info={}
)