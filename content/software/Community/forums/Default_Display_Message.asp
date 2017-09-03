<%@ LANGUAGE="VBSCRIPT" %>
<!-- #include File="../mapvar.asp" -->

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">

<title>Default Display Message</title>

<body>
<basefont name="VERDANA" color="#000000" >

<p><b>Welcome</b> to the <b>&quot;<%Response.Write Request.QueryString("ConferenceTitle")%>&quot;</b> on-line forum.
</p><p>The <strong>center window</strong> shows the subjects of all messages posted on
the forum, as well as the name of the person posting the message.
Click on one of those subjects, and that message will appear in this window. You can
then read the message and <strong>Reply</strong> to it if you wish.</p> 

<p>The menu at the top of the screen allows you to: </p>
<div align="center"><center>

<table border="0" cellpadding="3" cellspacing="3" bordercolor="#000000">
  <tr>
    <td><font size="3" color="#FFFFFF"><a href="../index.asp" target="_top">Return
    to the Community Web Home Page.</a></font></td>
  </tr>
  <tr>
    <td><p align="left"><font size="3"><a href="default.htm" target="message">Return to the Forums Page</a></font></td>
  </tr>
  <tr>
    <td><p align="left"><font size="3"><%Response.Write "<a href=""Login_Guts_Frame.asp?ConferenceID=" & Request.QueryString("ConferenceID") & "&PostType=N&MessageTitle=&ParentMessageID=0"">"%>Start a new Message</a></font></td>
  
  </tr>
</table>
</center></div>

<p><font size="+1"><strong>Note:  </strong></font> You may need to reload this page to see the
latest messages. To reload a page, just click on the button that says &quot;Refresh&quot; or 
&quot;Reload&quot; on the tool bar at the top.  If you have questions, comments, or problems, please contact the system administrator.</p>

</body>
</html>
