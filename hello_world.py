from PySide import QtGui


class _GCProtector(object):
	widgets = []

app = QtGui.QApplication.instance()
if not app:
	app = QtGui.QApplication([])

def main():
	msg_box = QtGui.QMessageBox(MaxPlus.GetQMaxWindow())
	_GCProtector.widgets.append(msg_box)

	msg_box.setText("Hello World!")
	msg_box.show()



if __name__ == '__main__':
	main()