import os
import shutil
from qgis.utils import iface

layer = iface.activeLayer()
selection = layer.selectedFeatures()
print('Jumlat File Ter-select = '+ str(len(selection)))
dest = 'D:/A/BTL-4/'
for feature in selection:
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    shutil.copy(feature[0],dest)
    print('Copying '+feature[0]+' to '+dest)
print ('...')
print('SELESAI')
