from PyQt5.QtWidgets import QLineEdit
from qgis.utils import iface

def onPressed():
    layer = iface.activeLayer()
    selection = layer.selectedFeatures()
    for feature in selection:
        attrs = feature.attributes()
        layer.changeAttributeValue(feature.id(),5,int(LineEdit.text()))

LineEdit = QLineEdit()
LineEdit.returnPressed.connect(onPressed)
LineEdit.setWindowFlags(Qt.WindowStaysOnTopHint)
LineEdit.show()
