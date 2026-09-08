"""双击配套启动文件打开Windows桌面程序。"""
from pathlib import Path
import sys


def main():
    if sys.platform!='win32':
        raise RuntimeError('本程序仅支持Windows。')
    import ctypes
    kernel=ctypes.windll.kernel32
    kernel.CreateMutexW.restype=ctypes.c_void_p
    handle=kernel.CreateMutexW(None,False,'Local\\DocumentSmartFillerDesktop')
    if kernel.GetLastError()==183:
        kernel.CloseHandle(ctypes.c_void_p(handle))
        ctypes.windll.user32.MessageBoxW(None,'程序已经打开，请回到现有窗口。','文档智能填表系统',64)
        return
    try:
        from src.界面 import App
        App(root_dir=Path(__file__).resolve().parent).mainloop()
    except Exception as error:
        ctypes.windll.user32.MessageBoxW(None,'程序未能启动。请先双击“安装依赖.bat”。\n\n'+str(error),'启动说明',16)
        raise
    finally:
        kernel.CloseHandle(ctypes.c_void_p(handle))


if __name__=='__main__':
    main()