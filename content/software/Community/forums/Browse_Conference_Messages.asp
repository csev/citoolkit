<%@ Language=VBScript %>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
<!-- #INCLUDE FILE="../adovbs.inc" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>

<BODY bgcolor="White" color="Black" link="Black" vlink="Black">

<%

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("Connection1_ConnectionString")

sql = "SELECT * FROM Messages WHERE ConferenceID = " & Request.QueryString("ConferenceID") & " AND ParentMessageID = 0"

if Request.QueryString("ConferenceModerated") = TRUE then
  sql = sql & " AND Moderator_Approved = TRUE"
end if

Set RecordSet1 = Server.CreateObject("ADODB.Recordset")
conn.CursorLocation = adUseClient
RecordSet1.Open sql, conn, adOpenKeyset, , adCmdText
Set Recordset1.ActiveConnection = Nothing

if RecordSet1.RecordCount = 0 then
  Response.Write "<p><b>   Currently, there are no messages in this forum.</b><br>"
else
  'Start formating of the list
  Response.Write vbcrlf & vbcrlf & "<dl><p>" & vbcrlf & vbcrlf

  call recurse_messages (RecordSet1)
end if

Recordset1.Close
conn.close

%>


<script runat = server language=vbscript>

sub recurse_messages(RecordSetMain)

Dim Recordset2

Set Recordset2 = RecordsetMain

Do While (Not RecordSet2.eof) 
    
    Dim Recordset3
    Dim RSusername
    
    sql = "SELECT FirstName, LastName from Users where USERID = " & Recordset2.fields("UserID")
    
    Set RSusername = Server.CreateObject("ADODB.Recordset")
	RSusername.Open sql, conn, adOpenKeyset, , adCmdText
    
    If not RSusername.EOF then
		Response.Write "<dt>" & "<b>" & "<a HREF=""http:Display_Message.asp?MessageID=" & Cint(Recordset2.fields("MessageID")) & """ target=message> "  & Recordset2.fields("MessageTitle") & "</a></b>"
    
		'set italics and display the user ID and Message Date
		Response.Write "<i>  " & RSusername.fields("FirstName") & "  " & RSusername.fields("LastName") & " " & Recordset2.fields("MessageDate") & "</i></dt>"
    end if
		
	sql = "SELECT * FROM Messages WHERE ConferenceID = " & Request.QueryString("ConferenceID") & " AND ParentMessageID = " & Recordset2.fields("MessageID")

	if Request.QueryString("ConferenceModerated") = TRUE then
      sql = sql & " AND Moderator_Approved = TRUE"
    end if

	Set Recordset3 = Server.CreateObject("ADODB.Recordset")
	Recordset3.Open sql, conn, adOpenKeyset, , adCmdText
	
	if Recordset3.RecordCount > 0 then
	    Response.Write vbcrlf & "<dd><dl>"
		call recurse_messages(Recordset3)
		Response.Write "</dd></dl>" & vbcrlf & vbcrlf
	End If	

	Recordset2.moveNext
		
loop

Recordset3.Close
	
end sub


</script>
</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>
