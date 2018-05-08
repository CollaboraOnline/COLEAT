#!/bin/sh

tstamp=$(date +'%Y%m%d').$(git log -1 --pretty=format:%h)

zip coleat-$tstamp.zip bin/*.exe bin/*.dll
