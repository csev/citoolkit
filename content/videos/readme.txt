This directory must be modified depending on whether it is put up on the web
or burned onto a CD.  There are comments in the index.html file

When the web site is hosted on the network, certain files (*.ram) must
have the proper URL.  Two batch scripts have been provided for this purpose.

  dofiles.bat - Run makefile for each file with the appropriate height and width
  makefile.bat - Create the necessary auxiliary files for each .rm file

This process also uses the videoskel.htm file and the sed15.exe file

To put up on the web, you must edit the file makefile.bat and put in the
right URL.  This could either be a rtsp: URL or a http: URL.

Once you have edited makefile run dofiles to put the proper URLs in the *.ram files.
