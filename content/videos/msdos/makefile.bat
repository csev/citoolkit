
echo rtsp://www.mel.org/citoolkit/content/videos/%3.rm > %3.ram
echo http://www.mel.org/citoolkit/content/videos/%3.rm > %3.ram

echo file:%3.rm > %3.rpm

sed15 -e s/HHH/%2/ -e s/WWW/%1/ -e s/FFF/%3/ < videoskel.htm > %3.htm
