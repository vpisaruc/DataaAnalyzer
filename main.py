from my_interface import *
import sys

"""Точка входа"""
def main():
    app = QtWidgets.QApplication(sys.argv)
    mainWindow = Window()
    mainWindow.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()