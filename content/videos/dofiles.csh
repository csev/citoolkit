
foreach i ( video1-160 video2-160 video3-160  video4-160 ) 

  csh -x makefile.csh 160 120 $i

end

foreach i ( video1-320 video2-320 video3-320  video4-320 ) 

  csh -x makefile.csh 320 240 $i

end

