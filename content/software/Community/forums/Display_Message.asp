<%@ LANGUAGE="VBSCRIPT" %>
<!-- #include File="../mapvar.asp" -->

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">

<title>Display Message</title>
<body bgcolor="White" color="Black" link="Black" vlink="Black">
<basefont name="ARIAL" color="#000000">

<%
	DIM RSusername
	DIM RSmessage

	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open  Session("Connection1_ConnectionString"),"",""
	
	sql = "SELECT * FROM Messages where MessageID ="& cstr(Request.QueryString("MessageID"))
	Set RSmessage = Server.CreateObject("ADODB.Recordset")
	RSmessage.Open sql, conn, adOpenKeyset, , adCmdText
	
	sql = "SELECT FirstName, LastName FROM Users where UserID = " & cstr(RSmessage.fields("UserID"))
	Set RSusername = Server.CreateObject("ADODB.Recordset")
	RSusername.Open sql, conn, adOpenKeyset, , adCmdText

  Response.Write "<p><a href=""Login_Guts_Frame.asp?ConferenceID=" & RSmessage.Fields("ConferenceID") & "&PostType=N&MessageTitle=&ParentMessageID=0""> New Message</a>&nbsp;&nbsp;" & vblf
  'Response.Write "<p><a href=""Post.asp?ConferenceID=" & RSmessage.Fields("ConferenceID") & "&PostType=N&MessageTitle=&ParentMessageID=0""> New Message</a>&nbsp;&nbsp;" & vblf
  Response.Write "<a href=""Login_Guts_Frame.asp?ConferenceID=" & RSmessage.Fields("ConferenceID") & "&PostType=R&MessageTitle=" & RSmessage.Fields("MessageTitle") & "&ParentMessageID=" & RSmessage.Fields("MessageID")& """> Reply Message</a></p>&nbsp;&nbsp;" & vblf
  'Response.Write "<a href=""Post.asp?ConferenceID=" & RSmessage.Fields("ConferenceID") & "&PostType=R&MessageTitle=" & RSmessage.Fields("MessageTitle") & "&ParentMessageID=" & RSmessage.Fields("MessageID")& """> Reply Message</a></p>&nbsp;&nbsp;" & vblf
  '<a href="prev_post.asp">&lt;&lt; Previous</a>&nbsp;&nbsp;<a href="next_post.asp">Next &gt;&gt;</a></p><br> 

  response.write "<i><b>" & RSmessage.Fields("MessageTitle") & "</b></i><br><br>"
  response.write "From: " & RSusername.Fields("FirstName") & " " & RSusername.Fields("LastName")
  Response.write "&nbsp; Date: " & RSmessage.Fields("MessageDate") & "<br><br>"
  Response.Write RSmessage.Fields("MessageBody") & "<br>"

%>
</body>
</html>
