
echo http://www-personal.umich.edu/~csev/citoolkit/content/videos/$3.rm > $3.ram

echo http://www-personal.umich.edu/~csev/citoolkit/content/videos/$3.rm > $3.rpm

sed -e s/HHH/$2/ -e s/WWW/$1/ -e s/FFF/$3/ < videoskel.htm > $3.htm
