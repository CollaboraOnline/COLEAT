#!/bin/sh

COLEAT_VERSION_MAJOR=$(grep COLEAT_VERSION_MAJOR include/coleat-version.h | awk '{print $3}')
COLEAT_VERSION_MINOR=$(grep COLEAT_VERSION_MINOR include/coleat-version.h | awk '{print $3}')
COLEAT_GIT_HEAD=$(grep COLEAT_GIT_HEAD include/coleat-git-version.h  | awk '{print substr($3,1,7)}')

tstamp=$(TZ=UTC0 date +'%Y%m%d.%H%M')
prefix=coleat-$COLEAT_VERSION_MAJOR.$COLEAT_VERSION_MINOR-$tstamp.$COLEAT_GIT_HEAD
zip=$prefix.zip
tarball=$prefix.tar.gz

zip $zip README.txt

(cd bin; zip ../$zip *.exe *.dll)

zip $zip demo.vbs demo.exe demo.frm u1.odt u2.docx

git archive --prefix=$prefix/ -o $tarball $COLEAT_GIT_HEAD
