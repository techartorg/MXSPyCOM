from pymxs import runtime

from PySide2.QtWidgets import QMainWindow, QMessageBox, QWidget
import shiboken2

def main( ):
    main_window_widget = QWidget.find( runtime.windows.getMAXHWND( ) )
    main_window = shiboken2.wrapInstance( shiboken2.getCppPointer( main_window_widget )[ 0 ], QMainWindow )                                                            
    msg_box = QMessageBox( main_window )	
    msg_box.setText( 'Hello World!' )
    msg_box.show( )



if __name__ == '__main__':
    main( )
  