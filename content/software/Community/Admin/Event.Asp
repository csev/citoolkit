<%@ Language=VBScript%>
<!-- #include File="../mapvar.asp" -->
<!--#include file="checksecure.asp"-->
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>

<html>

<head>
<title>Admin-Events</title>
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

Sub EventInfo_onenter()
	'rsUserRole.SetSQLText("SELECT Roles.RoleName FROM Roles, UserRoles WHERE Roles.RoleID = UserRoles.RoleID AND     (UserRoles.UserID = "&session("UserID")& " )")
	'rsUserRole.Open
	if EventInfo.firstEntered then
		purgeDate.value = cstr(dateadd("m",-6,date()))
	End If
End Sub


Sub btnPurge_onclick()
	Recordset2.setSQLText("DELETE FROM EVENTS WHERE EventDate < #" & Request.Form("PurgeDate") & "#")
	Recordset2.open
	Recordset2.close
	vPurge = True
End Sub

</SCRIPT>
</head>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:8CC35CD6-E98B-11D0-B218-00A0C92764F5" id=PageObject1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="4233">
	<PARAM NAME="ExtentY" VALUE="1508">
	<PARAM NAME="State" VALUE="(ObjectName_Unmatched=\qEventInfo\q,NavigateMethods=(Rows=0),ExecuteMethods=(Rows=0),Properties=(Rows=0),References=(Rows=0))"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=SERVER>
/* VIPM PAGE DESCRIPTION
<DSC NAME="EventInfo">
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
if (typeof EventInfo_onbeforeserverevent == 'function' || typeof EventInfo_onbeforeserverevent == 'unknown')
	thisPage.advise('onbeforeserverevent', 'EventInfo_onbeforeserverevent()');

EventInfo = thisPage;
EventInfo.location = "../Admin/Event.Asp";
EventInfo.navigate = new Object;
EventInfo.navigate.show = Function('thisPage.invokeMethod("", "show", this.show.arguments);');
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

EventInfo = thisPage;
EventInfo.location = "../Admin/Event.Asp";
EventInfo.navigate = new Object;
EventInfo.navigate.show = Function('return;');

	thisPage._objEventManager.adviseDefaultHandler('EventInfo','onenter');
	thisPage._objEventManager.adviseDefaultHandler('EventInfo','onexit');
	thisPage._objEventManager.adviseDefaultHandler('EventInfo','onshow');
	thisPage.registerVTable(thisPage.navigate, PAGE_NAVIGATE);
}

function _PO_dtor()
{
if (thisPage._redirect == '')
	_PO_OutputClientCode();
}

</SCRIPT>


<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Recordset2 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject=\qTables\q,TCDBObjectName=\qEvents\q,TCControlID_Unmatched=\qRecordset2\q,TCPPConn=\qConnection1\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qEvents\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
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
	cmdTmp.CommandText = '`Events`';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	Recordset2.setRecordSource(rsTmp);
	if (thisPage.getState('pb_Recordset2') != null)
		Recordset2.setBookmark(thisPage.getState('pb_Recordset2'));
}
function _Recordset2_ctor()
{
	CreateRecordset('Recordset2', _initRecordset2, null);
}
function _Recordset2_dtor()
{
	Recordset2._preserveState();
	thisPage.setState('pb_Recordset2', Recordset2.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=rsSiteInfo style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject=\qTables\q,TCDBObjectName=\qSiteInfo\q,TCControlID_Unmatched=\qrsSiteInfo\q,TCPPConn=\qConnection1\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qSiteInfo\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initrsSiteInfo()
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
	cmdTmp.CommandText = '`SiteInfo`';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	rsSiteInfo.setRecordSource(rsTmp);
	rsSiteInfo.open();
	if (thisPage.getState('pb_rsSiteInfo') != null)
		rsSiteInfo.setBookmark(thisPage.getState('pb_rsSiteInfo'));
}
function _rsSiteInfo_ctor()
{
	CreateRecordset('rsSiteInfo', _initrsSiteInfo, null);
}
function _rsSiteInfo_dtor()
{
	rsSiteInfo._preserveState();
	thisPage.setState('pb_rsSiteInfo', rsSiteInfo.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<body vLink=White aLink=White Link=White>

<table border="1" width="350" cellpadding=3 cellspacing=0>
  <tr>
	<td width=180 bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="../index.asp" target="_top" vlink="White">Home</a></strong></big></font></center></td>
    <td width=180 bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="AdminSecure.asp" target="_top" vlink="White">Administration</a></strong></big></font></center></td>
    <td><strong><font face="Verdana"><center>Events</center></font></strong></td>
  </tr>
</table>

<% 
	If vPurge = True then
		Response.Write "<font color=red>Database purged before " & Request.form("purgeDate") & "</font><br><br>"
	End if
	
	' If the site requires approval then show events needing
	' approval to any users who can approve them
	If rsSiteInfo.fields.getValue("CalApproval") = True Then 

		rsUserRole.SetSQLText("SELECT * From UserRoles Where UserID = "&session("UserID")& " AND RoleID=2")
		rsUserRole.Open
		If not rsUserRole.eof then
%>

<p><font face="Verdana"><strong>Events Pending Approval</strong></font>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Recordset1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSELECT\s*\sFROM\sEVENTS\sWHERE\sAPPROVEUserID\sIS\sNULL\sOR\sApproveUserID\s=0\sORDER\sBY\sEventPublishDate\sDESC,\sEventDate\sDESC\r\n\q,TCControlID_Unmatched=\qRecordset1\q,TCPPConn=\qConnection1\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qEvents\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSELECT\s*\sFROM\sEVENTS\sWHERE\sAPPROVEUserID\sIS\sNULL\sOR\sApproveUserID\s=0\sORDER\sBY\sEventPublishDate\sDESC,\sEventDate\sDESC\r\n\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
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
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
	cmdTmp.CommandText = 'SELECT * FROM EVENTS WHERE APPROVEUserID IS NULL OR ApproveUserID =0 ORDER BY EventPublishDate DESC, EventDate DESC ';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	Recordset1.setRecordSource(rsTmp);
	Recordset1.open();
	if (thisPage.getState('pb_Recordset1') != null)
		Recordset1.setBookmark(thisPage.getState('pb_Recordset1'));
}
function _Recordset1_ctor()
{
	CreateRecordset('Recordset1', _initRecordset1, null);
}
function _Recordset1_dtor()
{
	Recordset1._preserveState();
	thisPage.setState('pb_Recordset1', Recordset1.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<br>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" id=Grid1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="8043">
	<PARAM NAME="_ExtentY" VALUE="3889">
	<PARAM NAME="DataConnection" VALUE="">
	<PARAM NAME="SourceType" VALUE="">
	<PARAM NAME="Recordset" VALUE="Recordset1">
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
	<PARAM NAME="ColumnsNames" VALUE='"=[EventName]","=[EventPublishDate]","=[EventDate]","=getButton([EventID])"'>
	<PARAM NAME="columnIndex" VALUE="0,1,2,3">
	<PARAM NAME="displayWidth" VALUE="68,68,68,68">
	<PARAM NAME="Coltype" VALUE="1,1,1,1">
	<PARAM NAME="formated" VALUE="0,0,0,0">
	<PARAM NAME="DisplayName" VALUE='"Event Name","Date to show event","Event Date","Action"'>
	<PARAM NAME="DetailAlignment" VALUE=",,,">
	<PARAM NAME="HeaderAlignment" VALUE=",,,">
	<PARAM NAME="DetailBackColor" VALUE=",,,">
	<PARAM NAME="HeaderBackColor" VALUE=",,,">
	<PARAM NAME="HeaderFont" VALUE=",,,">
	<PARAM NAME="HeaderFontColor" VALUE=",,,">
	<PARAM NAME="HeaderFontSize" VALUE=",,,">
	<PARAM NAME="HeaderFontStyle" VALUE=",,,">
	<PARAM NAME="DetailFont" VALUE=",,,">
	<PARAM NAME="DetailFontColor" VALUE=",,,">
	<PARAM NAME="DetailFontSize" VALUE=",,,">
	<PARAM NAME="DetailFontStyle" VALUE=",,,">
	<PARAM NAME="ColumnCount" VALUE="4">
	<PARAM NAME="CurStyle" VALUE="Teal Titles">
	<PARAM NAME="TitleFont" VALUE="Arial">
	<PARAM NAME="titleFontSize" VALUE="4">
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
	<PARAM NAME="WidthSelectionMode" VALUE="2">
	<PARAM NAME="GridWidth" VALUE="100">
	<PARAM NAME="EnablePaging" VALUE="-1">
	<PARAM NAME="ShowStatus" VALUE="-1">
	<PARAM NAME="StyleValue" VALUE="455625">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<!--#INCLUDE FILE="../_ScriptLibrary/RSNavBar.ASP"-->
<!--#INCLUDE FILE="../_ScriptLibrary/DataGrid.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initGrid1()
{
Grid1.pageSize = 20;
Grid1.setDataSource(Recordset1);
Grid1.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolordark=Gray bordercolorlight=Silver bgcolor=Silver border=2 cols=4 rules=ROWS WIDTH=100%';
Grid1.headerAttributes = '   bgcolor=Teal align=Left';
Grid1.headerWidth[0] = ' WIDTH=22%';
Grid1.headerWidth[1] = ' WIDTH=22%';
Grid1.headerWidth[2] = ' WIDTH=22%';
Grid1.headerWidth[3] = ' WIDTH=22%';
Grid1.headerFormat = '<Font face="Arial" size=4 color=Black> <b>';
Grid1.colHeader[0] = '\'Event Name\'';
Grid1.colHeader[1] = '\'Date to show event\'';
Grid1.colHeader[2] = '\'Event Date\'';
Grid1.colHeader[3] = '\'Action\'';
Grid1.rowAttributes[0] = '  bgcolor = White align=Left bordercolordark=Gray bordercolorlight=Silver';
Grid1.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid1.colAttributes[0] = '  WIDTH=22%';
Grid1.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[0] = 'Recordset1.fields.getValue(\'EventName\')';
Grid1.colAttributes[1] = '  WIDTH=22%';
Grid1.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[1] = 'Recordset1.fields.getValue(\'EventPublishDate\')';
Grid1.colAttributes[2] = '  WIDTH=22%';
Grid1.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[2] = 'Recordset1.fields.getValue(\'EventDate\')';
Grid1.colAttributes[3] = '  WIDTH=22%';
Grid1.colFormat[3] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[3] = 'getButton(Recordset1.fields.getValue(\'EventID\'))';
Grid1.navbarAlignment = 'Right';
var objPageNavbar = Grid1.showPageNavbar(170,1);
objPageNavbar.getButton(0).value = '    |<    ';
objPageNavbar.getButton(1).value = '    <<    ';
objPageNavbar.getButton(2).value = '    >>    ';
objPageNavbar.getButton(3).value = '    >|    ';
Grid1.hasPageNumber = true;
}
function _Grid1_ctor()
{
	CreateDataGrid('Grid1',_initGrid1);
}
</SCRIPT>

<%	Grid1.display %>


<!--METADATA TYPE="DesignerControl" endspan-->
<%
			If Recordset1.EOF then
				Response.Write " -- No Pending Events..."
			End If	
		End If
		rsUserRole.Close

	End If
%>

<br>
<font face="Verdana"><strong>Event Actions</strong></font>
<br>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=rsUserRole style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\qSelect\s*\sFrom\sUserRoles\sWhere\sUserID=\s?\q,TCControlID_Unmatched=\qrsUserRole\q,TCPPConn=\qConnection1\q,TCPPDBObject=\qDE\sCommands\q,TCPPDBObjectName=\qSiteInfo\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\qSelect\s*\sFrom\sUserRoles\sWhere\sUserID=\s?\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q100\q,TCCommTimeout_Unmatched=\q30\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCNoCache\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initrsUserRole()
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
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 30;
	cmdTmp.CommandText = 'Select * From UserRoles Where UserID= ?';
	rsTmp.CacheSize = 100;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	rsUserRole.setRecordSource(rsTmp);
}
function _rsUserRole_ctor()
{
	CreateRecordset('rsUserRole', _initrsUserRole, null);
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<%

	' Allow a user to add events if they have been granted permssion to do so
	rsUserRole.SetSQLText("SELECT * From UserRoles Where UserID = "&session("UserID")& " AND RoleID=6")
	rsUserRole.Open
	If not rsUserRole.eof then
		'Response.Write "- <A href=eventadd.asp?action=new>Add an Event</A><br>"
		Response.Write "<input type=button onclick=""window.location.href='eventadd.asp?action=new'"" value=""Add an Event""><br>"
	End If
	rsUserRole.Close
	
	' Allow a user to edit events if they have been granted permssion to do so
	rsUserRole.SetSQLText("SELECT * From UserRoles Where UserID = "&session("UserID")& " AND RoleID=4")
	rsUserRole.Open
	If not rsUserRole.eof then
		'Response.Write "- <A href=eventedit.asp>Modify an Event</a><br>"
		Response.Write "<input type=button onclick=""window.location.href='editchoose.asp'"" value=""Edit/Delete Events"" id=button1 name=button1><br>"
	End If
	rsUserRole.Close
	
	' Allow a user to edit events if they have been granted permssion to do so
	rsUserRole.SetSQLText("SELECT * From UserRoles Where UserID = "&session("UserID")& " AND RoleID=9")
	rsUserRole.Open
	If not rsUserRole.eof then %>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnPurge 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 108px" width=108>
	<PARAM NAME="_ExtentX" VALUE="2858">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnPurge">
	<PARAM NAME="Caption" VALUE="Purge Events">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnPurge()
{
	btnPurge.value = 'Purge Events';
	btnPurge.setStyle(0);
}
function _btnPurge_ctor()
{
	CreateButton('btnPurge', _initbtnPurge, null);
}
</script>
<% btnPurge.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<font face=verdana size=-1>Purge Date:</font>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" id=purgeDate style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="purgeDate">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="20">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initpurgeDate()
{
	purgeDate.setStyle(TXT_TEXTBOX);
	purgeDate.setMaxLength(20);
	purgeDate.setColumnCount(20);
}
function _purgeDate_ctor()
{
	CreateTextbox('purgeDate', _initpurgeDate, null);
}
</script>
<% purgeDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<br>

	
	<%
		
	End If
	rsUserRole.Close
	
	
%>

<p>&nbsp;</p>
<!-- #include File="../copyright.asp" -->
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</html>
<script runat=server language=vbscript>
function getbutton(ID)

	getbutton1 = "<input type=button value=" & chr(34) & "Approve" & chr(34)
	getbutton2 = " onclick=" & chr(34) & "window.location.href='approve.asp?action=approve&EventID=" & cstr(ID) & "'" & chr(34) & ">"

	getbutton = getbutton1 & getbutton2

End Function
</script>