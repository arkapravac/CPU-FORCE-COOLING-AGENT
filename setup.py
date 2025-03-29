import PyInstaller.__main__
import sys
import os

def create_executable():
    script_path = os.path.join(os.path.dirname(__file__), 'cpu_cooling_agent.py')
    PyInstaller.__main__.run([
        script_path,
        '--onefile',
        '--windowed',
        '--name=CPU_Cooling_Agent',
        '--icon=NONE',
        '--add-data=requirements.txt;.',
        '--hidden-import=sklearn.linear_model',
        '--hidden-import=sklearn.utils._typedefs',
        '--hidden-import=sklearn.utils._heap',
        '--hidden-import=sklearn.utils._sorting',
        '--hidden-import=sklearn.utils._vector_sentinel',
        '--hidden-import=wmi',
        '--hidden-import=comtypes',
        '--hidden-import=win32com.client'
    ])

if __name__ == '__main__':
    create_executable()