<%@ Language=VBScript %>
<% Response.buffer = true %>
<!-- #include File="../adovbs.inc" -->
<%
	Set conn = server.CreateObject("ADODB.Connection")
	conn.Open Application("Connection1_ConnectionString")
	Set rs = server.CreateObject("ADODB.Recordset")
	rs.Open "SiteInfo", Conn, adOpenStatic,adLockOptimistic, adCmdTable
	
	rs.Movefirst
	
	If Request.QueryString("action") = "change" then
		rs.MoveFirst
		For each fld in rs.Fields
			'Response.Write fld.name & "<br>"
			If IsEmpty(Request.Form(fld.name)) Then
				If fld.type = adBoolean then
					fld.value = "False"
				End If
				
			Else	
				If fld.Value = Request.Form(fld.name) then
					'
				Else
					Select Case fld.Type
					Case adChar
						fld.Value = Request.Form(fld.name)
					Case adVarChar
						fld.Value = Request.Form(fld.name)
					Case adSmallInt
						fld.Value = cint(Request.Form(fld.name))
					Case adInteger
						fld.Value = cint(Request.Form(fld.name))
					Case adUnsignedTinyInt
						fld.Value = cint(Request.Form(fld.name))
					Case adDBTimeStamp
						fld.Value = Request.Form(fld.name)
					Case adBoolean
						'Response.Write(fld.name & ":" &  Request.Form(fld.name) & "<br>")
						fld.Value = "False"
						If Request.Form(fld.name) <> "" then
							fld.Value = "True"
						End If
					End Select
				End If
			End If
		Next
		rs.Update
	End If
	rs.Requery
	rs.movefirst
	
	fldCtr = 1
%>
<html>

<head>
<title>Admin-Site Information</title>
</head>

<body vLink=White aLink=White Link=White>
<%
	If rs("AllowRemoteSiteAdmin") = True Then
		If Request.ServerVariables("REMOTE_ADDR") <> Request.ServerVariables("LOCAL_ADDR") Then
			Response.Write("<Body><Font Size=5 Face=Verdana>Remote Site Administration is Forbidden</font></body>")
			Response.End
		Else
			'Response.Write "Admin Allowed"
		End If
	End If			
	
%>
<table border="1" width="350" cellpadding=3 cellspacing=0>
  <tr>
	<td width=180 bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="../index.asp" target="_top" vlink="White">Home</a></strong></big></font></center></td>
    <td width=180 bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="AdminSecure.asp" target="_top" vlink="White">Administration</a></strong></big></font></center></td>
    <td ><center><font face="Verdana"><strong>Site&nbsp;Information</strong></font></center></td>
  </tr>
</table>
<strong>

<form method="post" action="siteinfo.asp?action=change">
 <table border="0" width="400" bordercolorlight="#000000" cellpadding=1 cellspacing=1>
    <tr>
      <td width="60%" colspan="2" bgcolor="#008080"><font color="#ffffff" face="Verdana"><strong>Organization Info</strong></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Organization Name</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Address</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">City</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">State</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Zip</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Phone</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Fax</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Contact Person</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Contact Email</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Comments</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Home Page</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Include Copyright on each page?</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Copyright file to include</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Organization Logo</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"></td>
      <td width="50%"></td>
    </tr>
    <tr>
      <td width="60%"></td>
      <td width="50%"></td>
    </tr>
    <tr>
      <td width="60%" colspan="2" bgcolor="#008080"><font color="#ffffff" face="Verdana"><strong>Calendar Info</strong></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Default Mode</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Allow List View</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Allow Big Month View</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Allow Month View</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    
    <tr>
      <td width="60%"><font face="Verdana">Number of items per page</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Support Audience Type?</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Support Event Type?</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Require Approval of Event?</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Send Email to all approvers?</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Email for Approval notification if not all</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"></td>
      <td width="50%"></td>
    </tr>
    <tr>
      <td width="110%" colspan="2" bgcolor="#008080"><font face="Verdana" color="#ffffff"><strong>Security Information</strong></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Only Allow Event Admin at local computer? </font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Only Allow Site Admin at local computer?</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Only Allow User Admin at local computer?</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    
    <tr>
      <td width="60%"><font face="Verdana">Number of attempts before account is locked out:</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="110%" colspan="2" bgcolor="#008080"><font face="Verdana" color="#ffffff"><strong>Other Information</strong></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Site Custom Field #1 Label:</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Site Custom Field #2 Label:</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Site Custom Field #3 Label:</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
        <tr>
      <td width="60%"><font face="Verdana">Home Button URL (for Calendar pages):</font></td>
      <td width="50%"><font face="Verdana"><%=GetField()%></font></td>
    </tr>
    <tr>
      <td width="60%"></td>
      <td width="50%"></td>
    </tr>
    <tr>
      <td width="60%"></td>
      <td width="50%"></td>
    </tr>
  </table>
  <p><font face="Verdana"><input type="submit" value="Change Site Information" name="Submit"></font></p>

</form>
<!-- #include File="../copyright.asp" -->
</body>
</html>
<%
	rs.Close
	Set rs = nothing
	conn.close
	set conn = nothing

%>
<script runat=server language=vbscript>
Public function GetField

	Set fld = rs.Fields(fldCtr)
	
	Select Case fld.Type
    Case adChar
		FieldType = " Type=Text "
		FieldValue = " value=" & chr(34) & fld.Value & chr(34)
	Case adVarChar
		FieldType = " Type=Text "      
		FieldValue = " value=" & chr(34) & fld.Value & chr(34)
	Case adSmallInt
		FieldType = " Type=Text "      
		FieldValue = " value=" & chr(34) & fld.Value & chr(34)
	Case adInteger
		FieldType = " Type=Text "      
		FieldValue = " value=" & chr(34) & fld.Value & chr(34)
	Case adUnsignedTinyInt
		FieldType = " Type=Text "      
		FieldValue = " value=" & chr(34) & fld.Value & chr(34)
	Case adDBTimeStamp
		FieldType = " Type=Text "      
		FieldValue = " value=" & chr(34) & fld.Value & chr(34)
	Case adBoolean
		FieldType = " Type=CheckBox "
		FieldValue = "  value=" & chr(34) & fld.Value & chr(34)
		IF fld.Value = True then FieldValue = FieldValue & " checked "
	End Select
	
	FieldName = " name=" & chr(34) & fld.name & chr(34) & " id=" & cstr(fld.type) & " "
	If fld.DefinedSize > 20 then
		If fld.DefinedSize > 40 then
			FieldSize = " size=20 maxlength=" & cstr(fld.DefinedSize) 
		Else
			FieldSize = " size=20 maxlength=" & cstr(fld.DefinedSize) 
		End If
	Else
		FieldSize = " size=" & cstr(fld.DefinedSize) & "  maxlength=" & cstr(fld.DefinedSize) 
	End If

	GetField = "<Input " & FieldType & " " & FieldValue & FieldName & FieldSize & ">"

	If fld.name = "CalDefaultMode" then
		GetField = "<Select Name=CalDefaultMode>"
		Select Case fld.value
			Case "List"
				GetField = GetField & "<option></option><option selected>List</option><option>Month</option><option>Big Month</option>"
			Case "Month"
				GetField = GetField & "<option></option><option>List</option><option selected>Month</option><option>Big Month</option>"
			Case "Big Month"
				GetField = GetField & "<option></option><option>List</option><option>Month</option><option selected>Big Month</option>"
			Case Else
				GetField = GetField & "<option selected></option><option>List</option><option>Month</option><option>Big Month</option>"
		End Select
		GetField = GetField & "</select>"
	
	End If
	
	fldCtr = fldCtr + 1

end function

</script>