from PyQt5.QtWidgets import QPushButton
from qgis.utils import iface

def onPressed():
    layer = iface.activeLayer()
    selection = layer.selectedFeatures()
    for feature in selection:
        geom = feature.geometry()
        area = int(geom.area() / 10000)
        layer.changeAttributeValue(feature.id(),2,area)

UpdateArea = QPushButton()
UpdateArea.setText("Update")
UpdateArea.clicked.connect(onPressed)
UpdateArea.setWindowFlags(Qt.WindowStaysOnTopHint)
UpdateArea.show()
