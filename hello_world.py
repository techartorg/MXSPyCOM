from PySide2.QtWidgets import QMessageBox

def main( ):
	msg_box = QMessageBox( )	
	msg_box.setText( 'Hello World!' )
	msg_box.show( )



if __name__ == '__main__':
	main( )
