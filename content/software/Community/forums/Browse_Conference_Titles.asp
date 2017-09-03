<html>

<head>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=windows-1252">
<title>Library Conferences</title>
<body link="White" vlink="White">
<basefont name="ARIAL" color="#000000" size="-1"><%
If IsObject(Session("Site")) Then
    Set conn = Session("Site")
Else
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.open Session("Connection1_ConnectionString"),"",""
    Set Session("Site") = conn
End If

    sql = "SELECT * FROM Conferences"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn
%>


<%
On Error Resume Next

if rs.EOF then
	Response.Write "<font size=4>Sorry No Forums available.</font>"
else
	%>
<table BORDERCOLOR="#000000" BORDER="1" CELLSPACING="0" cellpadding="5">
<font FACE="Arial" COLOR="Black"><TBODY>

	<%
	do while Not rs.eof
 %>

  <tr VALIGN="TOP" BORDERCOLOR="#a9a9a9">
    <td bgcolor="#008080" ALIGN="RIGHT"><a
    href="http:View_Conference.asp?ConferenceID=<%=Server.HTMLEncode(rs.Fields("ConferenceID"))%>&ConferenceTitle=<%=Server.HTMLEncode(rs.Fields("ConferenceName"))%>&ConferenceModerated=<%=Server.HTMLEncode(rs.Fields("ConferenceModerated"))%>"
    target="_parent"><font name="Arial" size="+1"><%=Server.HTMLEncode(Trim(rs.Fields("ConferenceName").Value)) %><br>
    </font></td>
    <td ALIGN="LEFT"><font name="Arial" size="-1"><%=Server.HTMLEncode(Trim(rs.Fields("ConferenceDescription").Value)) %><br>
    </font></td>
  </tr>
<%
rs.MoveNext
loop%>
</TBODY>
</table>
<%
end if
%>
</a></font>
</body>
</html>
