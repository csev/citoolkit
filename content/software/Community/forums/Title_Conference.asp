<!-- #include File="../mapvar.asp" -->
<html>

<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 3.0">

<title>Conference Title</title>

<base target="main">
</head>

<body link="White" vlink="White">
<font SIZE="5" FACE="Arial" COLOR="#000000">

<p align="center"> </p>
</font>

<table border="1" width="100%" cellpadding=3 cellspacing=0>
  <tr>
    <td width=75 bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="../index.asp" target="_top">Home</a></strong></big></font></center></td>
    <td width=75 bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="default.htm" target="_top">Forums</a></strong></big></font></center></td>
    <td width="65%" bgcolor="White"><font face="Verdana" color="Black"><big><strong><%= Request.QueryString("ConferenceTitle") %> Discussion Forum</strong></big></font></td>
    <td width=75 bgcolor="#008080"><center><font face="Verdana" color="Black"><big><strong>
    <%Response.Write "<a vlink=""White"" target=""message"" href=""Login_Guts_Frame.asp?ConferenceID=" & Request.QueryString("ConferenceID") & "&PostType=N&MessageTitle=&ParentMessageID=0"">"%>Post</a></strong></big></font></center></td>
    <% 
    '<a href=<%Response.write """Login_Guts_Frame.asp?ConferenceID=" & Request.QueryString(ConferenceID) & "&PostType=N&MessageTitle=&ParentMessageID=0""" target="article">Post</a></strong></big></font></center></td>  
    %>
  
  </tr>
</table>


</body>
</html>
