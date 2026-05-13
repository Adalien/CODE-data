import os, sys
here = os.path.dirname(os.path.abspath(__file__))
os.chdir(here)
gui = os.path.join(here, '缺料分析_GUI.py')   # 缺料分析_GUI.py
with open(gui, encoding='utf-8') as f:
    src = f.read()
exec(compile(src, gui, 'exec'), {'__name__': '__main__', '__file__': gui})
