import os
import shutil
from qgis.utils import iface

#CONSTANT
layer = iface.activeLayer()
selection = layer.selectedFeatures()

#RUBAH
folder = '33 BTB 16 Sadap'
saluran = 'Turi Baru'

print('Jumlat File Ter-select = '+ str(len(selection)))
dest = f'D:/mrican/Inventori/{saluran}/{folder}/'
for feature in selection:
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    shutil.copy(feature[0],dest)
    print('Copying '+feature[0]+' to '+dest)
print ('...')
print('SELESAI')
