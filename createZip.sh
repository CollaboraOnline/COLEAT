#!/bin/sh

tstamp=$(TZ=UTC0 date +'%Y%m%d.%H%M').$(git log -1 --pretty=format:%h)
zip=coleat-$tstamp.zip

zip $zip README.txt

(cd bin; zip ../$zip *.exe *.dll)

zip $zip demo.vbs demo.exe demo.frm u1.odt u2.doc
