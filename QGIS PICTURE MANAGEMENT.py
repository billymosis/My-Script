selection = layer.selectedFeatures()
print('Jumlat File Ter-select = '+ str(len(selection)))
dest = 'D:/HAHAHAHA/'
for feature in selection:
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    shutil.copy(feature[0],dest)
    print('Copying '+feature[0]+' to '+dest)
print ('...')
print('SELESAI')
