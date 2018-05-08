#!/bin/sh

tstamp=$(TZ=UTC0 date +'%Y%m%d.%H%M').$(git log -1 --pretty=format:%h)

zip coleat-$tstamp.zip bin/*.exe bin/*.dll
