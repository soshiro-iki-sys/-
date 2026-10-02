import sys, zipfile, shutil, re
src=sys.argv[1]; tmp=src+'.tmp'
zin=zipfile.ZipFile(src); zout=zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data=zin.read(item.filename)
    if item.filename.startswith('ppt/theme/theme') and item.filename.endswith('.xml'):
        x=data.decode('utf-8')
        x=x.replace('<a:ea typeface=""/>','<a:ea typeface="メイリオ"/>')
        x=re.sub(r'<a:font script="Jpan" typeface="[^"]*"/>','<a:font script="Jpan" typeface="メイリオ"/>',x)
        data=x.encode('utf-8')
    zout.writestr(item,data)
zout.close(); zin.close(); shutil.move(tmp,src)
