from qgis.utils import iface
from PyQt5.QtWidgets import  QFileDialog
xname = QFileDialog.getOpenFileName(QFileDialog(), "x", "","csv(*.pdf)")
fname = str(xname)
layer = QgsProject.instance().mapLayer('[% @layer_id %]')
layer.changeAttributeValue([% $id %],2,fname[2:-16])
