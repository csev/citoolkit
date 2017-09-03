<%@ Language=VBScript%>
<% Response.buffer = true %>
<!-- #include File="../adovbs.inc" -->
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
<html>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:8CC35CD6-E98B-11D0-B218-00A0C92764F5" id=PageObject1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="4233">
	<PARAM NAME="ExtentY" VALUE="1508">
	<PARAM NAME="State" VALUE="(ObjectName_Unmatched=\quserrole\q,NavigateMethods=(Rows=0),ExecuteMethods=(Rows=0),Properties=(Rows=0),References=(Rows=0))"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=SERVER>
/* VIPM PAGE DESCRIPTION
<DSC NAME="userrole">
   <OBJECT NAME="navigate">
      <METHOD NAME="show" SCENARIOS="CLIENT,SERVER"/>
   </OBJECT>
</DSC>
VIPM PAGE DESCRIPTION */
</SCRIPT>
<%
Sub [_PO_OutputClientCode]()
%>
<SCRIPT LANGUAGE=JavaScript>
if (typeof userrole_onbeforeserverevent == 'function' || typeof userrole_onbeforeserverevent == 'unknown')
	thisPage.advise('onbeforeserverevent', 'userrole_onbeforeserverevent()');

userrole = thisPage;
userrole.location = "../Admin/userrole.asp";
userrole.navigate = new Object;
userrole.navigate.show = Function('thisPage.invokeMethod("", "show", this.show.arguments);');
</SCRIPT>
<%
End Sub
%>

<SCRIPT LANGUAGE=JavaScript RUNAT=SERVER>
function _PO_getClientAccessor(serverValue)
{
	if (serverValue == null)
		return 'null';
	return 'unescape("' + escape(serverValue) + '")';
}

function _PO_ctor()
{
	thisPage.getClientAccessor = _PO_getClientAccessor;

userrole = thisPage;
userrole.location = "../Admin/userrole.asp";
userrole.navigate = new Object;
userrole.navigate.show = Function('return;');

	thisPage._objEventManager.adviseDefaultHandler('userrole','onenter');
	thisPage._objEventManager.adviseDefaultHandler('userrole','onexit');
	thisPage._objEventManager.adviseDefaultHandler('userrole','onshow');
	thisPage.registerVTable(thisPage.navigate, PAGE_NAVIGATE);
}

function _PO_dtor()
{
if (thisPage._redirect == '')
	_PO_OutputClientCode();
}

</SCRIPT>


<!--METADATA TYPE="DesignerControl" endspan-->
<head>
<title>Admin-User Roles</title>
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

Sub userrole_onenter()
		If Request.QueryString("UserID") <> "" then
			UserID = Request.QueryString("UserID")
		Else
			UserID = Request.Form("UserID")
		End IF
		Recordset1.setSQLText("SELECT * FROM Users WHERE USERID="&cstr(UserID))
		Recordset1.open
		Recordset2.setSQLText("SELECT RoleID, RoleCategory, RoleName FROM Roles WHERE RoleID IN (SELECT RoleID FROM UserRoles WHERE UserID=" & cstr(UserID) & ")")
		Recordset2.open
		Recordset3.setSQLText("SELECT RoleID, RoleName FROM Roles WHERE RoleID NOT IN (SELECT RoleID FROM UserRoles WHERE UserID="&cstr(UserID)&")")
		Recordset3.open
		
End Sub


Sub Button2_onclick()
	If Request.QueryString("UserID") <> "" then
		UserID = Request.QueryString("UserID")
	Else
		UserID = Request.Form("UserID")
	End IF

	Set conn = server.CreateObject("ADODB.Connection")
	conn.Open Application("Connection1_ConnectionString")
	
	Set rs = conn.execute("DELETE FROM UserRoles WHERE USERID="&Cstr(userID) & " AND ROLEID IN ("&Request.Form("ROLEIDDELETE")&")")

	Set rs = nothing
	conn.close
	set conn = nothing

	Response.Redirect "userrole.asp?UserID=" & cstr(UserID)
	
End Sub

Sub Button1_onclick()
	If Request.QueryString("UserID") <> "" then
		UserID = Request.QueryString("UserID")
	Else
		UserID = Request.Form("UserID")
	End IF

	Set conn = server.CreateObject("ADODB.Connection")
	conn.Open Application("Connection1_ConnectionString")
	
	For each fld in Request.Form("NewRole")
		set rs = conn.execute("INSERT INTO UserRoles (UserID,RoleID) VALUES("&UserID & ","&fld &")")
	Next	
	
	Set rs = nothing
	conn.close
	set conn = nothing

	Response.Redirect "userrole.asp?UserID=" & cstr(UserID)
	
End Sub

</SCRIPT>
</head>

<body vLink=White aLink=White Link=White>

<%
	Set conn = server.CreateObject("ADODB.Connection")
	conn.Open Application("Connection1_ConnectionString")
	Set rs = server.CreateObject("ADODB.Recordset")
	rs.Open "SiteInfo", Conn, adOpenStatic,adLockOptimistic, adCmdTable

	If rs("AllowRemoteUserAdmin") = True Then
		If Request.ServerVariables("REMOTE_ADDR") <> Request.ServerVariables("LOCAL_ADDR") Then
			Response.Write("<Body><Font Size=5 Face=Verdana>Remote User Administration is Forbidden</font></body>")
			Response.End
		Else
			'Response.Write "Admin Allowed"
		End If
	End If	

	rs.Close
	Set rs = nothing
	conn.close
	set conn = nothing
		
	
%>

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Recordset1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject=\qTables\q,TCDBObjectName=\qUsers\q,TCControlID_Unmatched=\qRecordset1\q,TCPPConn=\qConnection1\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qUsers\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCNoCache\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initRecordset1()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('Connection1_ConnectionTimeout');
	DBConn.CommandTimeout = Application('Connection1_CommandTimeout');
	DBConn.CursorLocation = Application('Connection1_CursorLocation');
	DBConn.Open(Application('Connection1_ConnectionString'), Application('Connection1_RuntimeUserName'), Application('Connection1_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 2;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = '`Users`';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	Recordset1.setRecordSource(rsTmp);
}
function _Recordset1_ctor()
{
	CreateRecordset('Recordset1', _initRecordset1, null);
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<table border="1" width="350" cellpadding=3 cellspacing=0>
  <tr>
	<td width=180 bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="../index.asp" target="_top" vlink="White">Home</a></strong></big></font></center></td>
    <td width=180 bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="AdminSecure.asp" target="_top" vlink="White">Administration</a></strong></big></font></center></td>
    <td bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="Users.asp" >User&nbsp;Information</a></strong></big></font></center></td>
	<td><strong><font face="Verdana"><center>User&nbsp;Roles</center></font></strong></td>
  </tr>
</table>
<strong>
<br>
<font face=verdana>User Roles for 
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=Label1 style="HEIGHT: 17px; LEFT: 10px; TOP: 170px; WIDTH: 52px" 
	width=52>
	<PARAM NAME="_ExtentX" VALUE="1376">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="Label1">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="FirstName">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Label.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel1()
{
	Label1.setDataSource(Recordset1);
	Label1.setDataField('FirstName');
}
function _Label1_ctor()
{
	CreateLabel('Label1', _initLabel1, null);
}
</script>
<% Label1.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
&nbsp;
 <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=Label2 style="HEIGHT: 17px; LEFT: 10px; TOP: 187px; WIDTH: 51px" 
	width=51>
	<PARAM NAME="_ExtentX" VALUE="1349">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="Label2">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="LastName">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel2()
{
	Label2.setDataSource(Recordset1);
	Label2.setDataField('LastName');
}
function _Label2_ctor()
{
	CreateLabel('Label2', _initLabel2, null);
}
</script>
<% Label2.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
(
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=Label3 style="HEIGHT: 17px; LEFT: 10px; TOP: 204px; WIDTH: 53px" 
	width=53>
	<PARAM NAME="_ExtentX" VALUE="1402">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="Label3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="UserName">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel3()
{
	Label3.setDataSource(Recordset1);
	Label3.setDataField('UserName');
}
function _Label3_ctor()
{
	CreateLabel('Label3', _initLabel3, null);
}
</script>
<% Label3.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
)</font>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Recordset2 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject=\qTables\q,TCDBObjectName=\qRoles\q,TCControlID_Unmatched=\qRecordset2\q,TCPPConn=\qConnection1\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qRoles\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCNoCache\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initRecordset2()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('Connection1_ConnectionTimeout');
	DBConn.CommandTimeout = Application('Connection1_CommandTimeout');
	DBConn.CursorLocation = Application('Connection1_CursorLocation');
	DBConn.Open(Application('Connection1_ConnectionString'), Application('Connection1_RuntimeUserName'), Application('Connection1_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 2;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = '`Roles`';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	Recordset2.setRecordSource(rsTmp);
}
function _Recordset2_ctor()
{
	CreateRecordset('Recordset2', _initRecordset2, null);
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" height=147 id=Grid1 style="HEIGHT: 147px; LEFT: 0px; TOP: 0px; WIDTH: 350px" 
	width=350>
	<PARAM NAME="_ExtentX" VALUE="9260">
	<PARAM NAME="_ExtentY" VALUE="3889">
	<PARAM NAME="DataConnection" VALUE="">
	<PARAM NAME="SourceType" VALUE="">
	<PARAM NAME="Recordset" VALUE="Recordset2">
	<PARAM NAME="CtrlName" VALUE="Grid1">
	<PARAM NAME="UseAdvancedOnly" VALUE="0">
	<PARAM NAME="AdvAddToStyles" VALUE="-1">
	<PARAM NAME="AdvTableTag" VALUE="">
	<PARAM NAME="AdvHeaderRowTag" VALUE="">
	<PARAM NAME="AdvHeaderCellTag" VALUE="">
	<PARAM NAME="AdvDetailRowTag" VALUE="">
	<PARAM NAME="AdvDetailCellTag" VALUE="">
	<PARAM NAME="ScriptLanguage" VALUE="1">
	<PARAM NAME="ScriptingPlatform" VALUE="0">
	<PARAM NAME="EnableRowNav" VALUE="0">
	<PARAM NAME="HiliteColor" VALUE="">
	<PARAM NAME="RecNavBarHasNextButton" VALUE="-1">
	<PARAM NAME="RecNavBarHasPrevButton" VALUE="-1">
	<PARAM NAME="RecNavBarNextText" VALUE="   >   ">
	<PARAM NAME="RecNavBarPrevText" VALUE="   <   ">
	<PARAM NAME="ColumnsNames" VALUE="&quot;RoleCategory&quot;,&quot;RoleName&quot;,&quot;='&lt;input type=checkbox name=RoleIDDelete value=' + [RoleID] + '&gt;'&quot;">
	<PARAM NAME="columnIndex" VALUE="0,1,2">
	<PARAM NAME="displayWidth" VALUE="68,68,68">
	<PARAM NAME="Coltype" VALUE="1,1,1">
	<PARAM NAME="formated" VALUE="0,0,0">
	<PARAM NAME="DisplayName" VALUE='"Role Category","Role Capabilities","Delete"'>
	<PARAM NAME="DetailAlignment" VALUE=",,">
	<PARAM NAME="HeaderAlignment" VALUE=",,">
	<PARAM NAME="DetailBackColor" VALUE=",,">
	<PARAM NAME="HeaderBackColor" VALUE=",,">
	<PARAM NAME="HeaderFont" VALUE=",,">
	<PARAM NAME="HeaderFontColor" VALUE=",,">
	<PARAM NAME="HeaderFontSize" VALUE=",,">
	<PARAM NAME="HeaderFontStyle" VALUE=",,">
	<PARAM NAME="DetailFont" VALUE=",,">
	<PARAM NAME="DetailFontColor" VALUE=",,">
	<PARAM NAME="DetailFontSize" VALUE=",,">
	<PARAM NAME="DetailFontStyle" VALUE=",,">
	<PARAM NAME="ColumnCount" VALUE="3">
	<PARAM NAME="CurStyle" VALUE="Teal Titles">
	<PARAM NAME="TitleFont" VALUE="Verdana">
	<PARAM NAME="titleFontSize" VALUE="3">
	<PARAM NAME="TitleFontColor" VALUE="0">
	<PARAM NAME="TitleBackColor" VALUE="32896">
	<PARAM NAME="TitleFontStyle" VALUE="1">
	<PARAM NAME="TitleAlignment" VALUE="0">
	<PARAM NAME="RowFont" VALUE="Arial">
	<PARAM NAME="RowFontColor" VALUE="0">
	<PARAM NAME="RowFontStyle" VALUE="0">
	<PARAM NAME="RowFontSize" VALUE="2">
	<PARAM NAME="RowBackColor" VALUE="16777215">
	<PARAM NAME="RowAlignment" VALUE="0">
	<PARAM NAME="HighlightColor3D" VALUE="12632256">
	<PARAM NAME="ShadowColor3D" VALUE="8421504">
	<PARAM NAME="PageSize" VALUE="20">
	<PARAM NAME="MoveFirstCaption" VALUE="    |<    ">
	<PARAM NAME="MoveLastCaption" VALUE="    >|    ">
	<PARAM NAME="MovePrevCaption" VALUE="    <<    ">
	<PARAM NAME="MoveNextCaption" VALUE="    >>    ">
	<PARAM NAME="BorderSize" VALUE="2">
	<PARAM NAME="BorderColor" VALUE="268435455">
	<PARAM NAME="GridBackColor" VALUE="12632256">
	<PARAM NAME="AltRowBckgnd" VALUE="268435455">
	<PARAM NAME="CellSpacing" VALUE="1">
	<PARAM NAME="WidthSelectionMode" VALUE="1">
	<PARAM NAME="GridWidth" VALUE="350">
	<PARAM NAME="EnablePaging" VALUE="0">
	<PARAM NAME="ShowStatus" VALUE="-1">
	<PARAM NAME="StyleValue" VALUE="439241">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<!--#INCLUDE FILE="../_ScriptLibrary/RSNavBar.ASP"-->
<!--#INCLUDE FILE="../_ScriptLibrary/DataGrid.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initGrid1()
{
Grid1.pageSize = 0;
Grid1.setDataSource(Recordset2);
Grid1.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolordark=Gray bordercolorlight=Silver bgcolor=Silver border=2 cols=3 rules=ROWS WIDTH=350 nowrap';
Grid1.headerAttributes = '   bgcolor=Teal align=Left nowrap';
Grid1.headerWidth[0] = ' WIDTH=68';
Grid1.headerWidth[1] = ' WIDTH=68';
Grid1.headerWidth[2] = ' WIDTH=68';
Grid1.headerFormat = '<Font face="Verdana" size=3 color=Black> <b>';
Grid1.colHeader[0] = '\'Role Category\'';
Grid1.colHeader[1] = '\'Role Capabilities\'';
Grid1.colHeader[2] = '\'Delete\'';
Grid1.rowAttributes[0] = '  bgcolor = White align=Left nowrap  bordercolordark=Gray bordercolorlight=Silver';
Grid1.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid1.colAttributes[0] = '  WIDTH=68';
Grid1.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[0] = 'Recordset2.fields.getValue(\'RoleCategory\')';
Grid1.colAttributes[1] = '  WIDTH=68';
Grid1.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[1] = 'Recordset2.fields.getValue(\'RoleName\')';
Grid1.colAttributes[2] = '  WIDTH=68';
Grid1.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[2] = '\'<input type=checkbox name=RoleIDDelete value=\' + Recordset2.fields.getValue(\'RoleID\') + \'>\'';
}
function _Grid1_ctor()
{
	CreateDataGrid('Grid1',_initGrid1);
}
</SCRIPT>

<%	Grid1.display %>


<!--METADATA TYPE="DesignerControl" endspan-->
<br>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=Button2 style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 163px" 
	width=163>
	<PARAM NAME="_ExtentX" VALUE="4313">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="Button2">
	<PARAM NAME="Caption" VALUE="Delete Marked Roles">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initButton2()
{
	Button2.value = 'Delete Marked Roles';
	Button2.setStyle(0);
}
function _Button2_ctor()
{
	CreateButton('Button2', _initButton2, null);
}
</script>
<% Button2.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<p>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Recordset3 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject=\qTables\q,TCDBObjectName=\qRoles\q,TCControlID_Unmatched=\qRecordset3\q,TCPPConn=\qConnection1\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qRoles\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCNoCache\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initRecordset3()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = Application('Connection1_ConnectionTimeout');
	DBConn.CommandTimeout = Application('Connection1_CommandTimeout');
	DBConn.CursorLocation = Application('Connection1_CursorLocation');
	DBConn.Open(Application('Connection1_ConnectionString'), Application('Connection1_RuntimeUserName'), Application('Connection1_RuntimePassword'));
	var cmdTmp = Server.CreateObject('ADODB.Command');
	var rsTmp = Server.CreateObject('ADODB.Recordset');
	cmdTmp.ActiveConnection = DBConn;
	rsTmp.Source = cmdTmp;
	cmdTmp.CommandType = 2;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = '`Roles`';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	Recordset3.setRecordSource(rsTmp);
}
function _Recordset3_ctor()
{
	CreateRecordset('Recordset3', _initRecordset3, null);
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->

<font face=verdana>Roles that can be added:</font><br>
<%



	If Recordset1.EOF then
		Response.Write "No roles to add..."
	Else
	
		If not Recordset3.EOF then
			Response.Write "<SELECT Name=NewRole  multiple>"
			Recordset3.moveFirst
			Do while not Recordset3.EOF
				Response.Write("<option value=" & cstr(Recordset3.fields.getvalue("RoleID")) & ">")
				Response.Write(cstr(Recordset3.fields.getvalue("RoleName")) & "</option>" & vbcrlf)
				Recordset3.moveNext
			Loop
			Response.Write "</select>"
		End If
	End If
	Recordset3.moveFirst
	if not Recordset3.EOF then
%>

<p>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=Button1 style="HEIGHT: 27px; LEFT: 10px; TOP: 574px; WIDTH: 147px" 
	width=147>
	<PARAM NAME="_ExtentX" VALUE="3889">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="Button1">
	<PARAM NAME="Caption" VALUE="Add Selected Role">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initButton1()
{
	Button1.value = 'Add Selected Role';
	Button1.setStyle(0);
}
function _Button1_ctor()
{
	CreateButton('Button1', _initButton1, null);
}
</script>
<% Button1.display %>

<!--METADATA TYPE="DesignerControl" endspan-->

<% End if %>
<input type=hidden name=userID value="<%=Cstr(UserID)%>">
<!-- #include File="../copyright.asp" -->
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</html>
