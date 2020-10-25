from PyQt5.QtWidgets import  QLineEdit
from qgis.utils import iface

#CONSTANT
layer = iface.activeLayer()
selection = layer.selectedFeatures()

LineEdit = QLineEdit()
LineEdit.returnPressed.connect(onPressed)
LineEdit.setWindowFlags(QtCore.Qt.Window | QtCore.Qt.CustomizeWindowHint | Qt.WindowStaysOnTopHint)
LineEdit.show()

def onPressed():
    for feature in selection:
        attrs = feature.attributes()
        layer.changeAttributeValue(feature.id(),5,int(LineEdit.text()))
