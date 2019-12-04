from pymxs import runtime

try:
    import MaxPlus
except( ImportError ):
    import shiboken2
    MaxPlus = None

from PySide2.QtWidgets import QMainWindow, QMessageBox, QWidget


def main( ):
    if MaxPlus:
        main_window = MaxPlus.GetQMaxMainWindow( )
    else:
        main_window_widget = QWidget.find( int( runtime.windows.getMAXHWND( ) ) )
        main_window = shiboken2.wrapInstance( shiboken2.getCppPointer( main_window_widget )[ 0 ], QMainWindow )

    msg_box = QMessageBox( main_window )
    msg_box.setText( 'Hello World!' )
    msg_box.show( )



if __name__ == '__main__':
    main( )
