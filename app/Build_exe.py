import PyInstaller.__main__
import PyInstaller.config
import os

distpath ="--distpath=" + str(os.getcwd()) +  "\\builder"
workpath = "--workpath=" + str(os.getcwd()) +  "\\builder" + "\\tempo"
# PyInstaller.config.CONF['distpath'] = path
PyInstaller.__main__.run([
    'main.py',
    '--onedir',
    # '--onefile',
    '-nEIW_Upload_Support',
    # '--hidden-import=pya2l.cgen',
    '--windowed',
    #'--disable-windowed-traceback'
    distpath,
    workpath,
    '--clean',
    '-y'
])
