"""
A "Hello Word!" example as a Qt based messagebox. 
This is used as a test to ensure MXSPyCOM is working correctly with 3ds Max.
"""

from pymxs import runtime # type: ignore

try:
    import MaxPlus # type: ignore
except ImportError:
    import shiboken2 # type: ignore
    MaxPlus = None

from PySide2.QtWidgets import QMainWindow, QMessageBox, QWidget # type: ignore


def main() -> None:
    if MaxPlus:
        main_window = MaxPlus.GetQMaxMainWindow() # type: QMainWindow
    else:
        main_window_widget = QWidget.find(int(runtime.windows.getMAXHWND())) # type: QWidget
        main_window = shiboken2.wrapInstance(shiboken2.getCppPointer(main_window_widget)[0], QMainWindow) # type: QMainWindow

    msg_box = QMessageBox(main_window)
    msg_box.setText('Hello World!')
    msg_box.show()


if __name__ == '__main__':
    main()
