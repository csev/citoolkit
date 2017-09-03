<%@ Language=VBScript%>
<!--#include file="../mapvar.asp"-->
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
<html>

<head>
<title>Admin-Forums</title>
</head>
<script ID="serverEventHandlersVBS" LANGUAGE="vbscript" RUNAT="Server">

Sub Forums_onenter()
UserID = session("UserID")

Recordset1.setSQLText("SELECT Users.*, Messages.* FROM Messages, Users, ConferenceModertors "_
& "WHERE ConferenceModertors.UserID = " &  UserID _
& " AND Messages.ConferenceID = ConferenceModertors.ConferenceID" _
& " AND Users.UserID = " & UserID & " AND Messages.UserID = Users.UserID" _
& " AND (Messages.ApproveUserID = 0 OR Messages.ApproveUserID IS NULL)" )

Recordset1.open
End Sub

Sub btnPurge_onclick()
	Recordset2.setSQLText("DELETE FROM Messages WHERE MessageDate < #" & Request.Form("PurgeDate") & "#")
	Recordset2.open
	Recordset2.close
	vPurge = True
End Sub

</script>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:8CC35CD6-E98B-11D0-B218-00A0C92764F5" id=PageObject1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="4233">
	<PARAM NAME="ExtentY" VALUE="1508">
	<PARAM NAME="State" VALUE="(ObjectName_Unmatched=\qForums\q,NavigateMethods=(Rows=0),ExecuteMethods=(Rows=0),Properties=(Rows=0),References=(Rows=0))"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=SERVER>
/* VIPM PAGE DESCRIPTION
<DSC NAME="Forums">
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
if (typeof Forums_onbeforeserverevent == 'function' || typeof Forums_onbeforeserverevent == 'unknown')
	thisPage.advise('onbeforeserverevent', 'Forums_onbeforeserverevent()');

Forums = thisPage;
Forums.location = "../Admin/Forums.asp";
Forums.navigate = new Object;
Forums.navigate.show = Function('thisPage.invokeMethod("", "show", this.show.arguments);');
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

Forums = thisPage;
Forums.location = "../Admin/Forums.asp";
Forums.navigate = new Object;
Forums.navigate.show = Function('return;');

	thisPage._objEventManager.adviseDefaultHandler('Forums','onenter');
	thisPage._objEventManager.adviseDefaultHandler('Forums','onexit');
	thisPage._objEventManager.adviseDefaultHandler('Forums','onshow');
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Recordset1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject_Unmatched=\qSQL\sStatement\q,TCDBObjectName_Unmatched=\q\q,TCControlID_Unmatched=\qRecordset1\q,TCPPConn=\qConnection1\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qAdminFunctions\q,RCDBObject=\qRCSQLStatement\q,TCSQLStatement_Unmatched=\q\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
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
	cmdTmp.CommandType = 1;
	cmdTmp.CommandTimeout = 10;
//Recordset DTC error: Failed to get command text
	cmdTmp.CommandText = '';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	Recordset1.setRecordSource(rsTmp);
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
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Recordset2 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject=\qTables\q,TCDBObjectName=\qMessages\q,TCControlID_Unmatched=\qRecordset2\q,TCPPConn=\qConnection1\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qMessages\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
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
	cmdTmp.CommandText = '`Messages`';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	Recordset2.setRecordSource(rsTmp);
	Recordset2.open();
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
<body vLink=White aLink=White Link=White>

<table border="1" cellpadding=3 cellspacing=0>
  <tr>
    <td bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="../index.asp" target="_top" vlink="White">Home</a></strong></big></font></center></td>
    <td bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="AdminSecure.asp" target="_top" vlink="White">Administration</a></strong></big></font></center></td>
    <td><strong><font face="Verdana"><center>Manage Forums</center></font></strong>
  </tr>
</table>
<strong>
<br>


<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" height=175 id=Grid1 style="HEIGHT: 175px; LEFT: 0px; TOP: 0px; WIDTH: 750px" 
	width=750>
	<PARAM NAME="_ExtentX" VALUE="15875">
	<PARAM NAME="_ExtentY" VALUE="3704">
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
	<PARAM NAME="ColumnsNames" VALUE='"=[FirstName] + "" "" + [LastName]","MessageTitle","MessageBody","MessageDate","=getbutton([MessageID])"'>
	<PARAM NAME="columnIndex" VALUE="0,1,2,3,4">
	<PARAM NAME="displayWidth" VALUE="90,120,270,55,35">
	<PARAM NAME="Coltype" VALUE="1,1,1,1,1">
	<PARAM NAME="formated" VALUE="0,0,0,0,0">
	<PARAM NAME="DisplayName" VALUE='"Name","Subject","Message","Date","Action"'>
	<PARAM NAME="DetailAlignment" VALUE=",,,,">
	<PARAM NAME="HeaderAlignment" VALUE=",,,,">
	<PARAM NAME="DetailBackColor" VALUE=",,,,">
	<PARAM NAME="HeaderBackColor" VALUE=",,,,">
	<PARAM NAME="HeaderFont" VALUE=",,,,">
	<PARAM NAME="HeaderFontColor" VALUE=",,,,">
	<PARAM NAME="HeaderFontSize" VALUE=",,,,">
	<PARAM NAME="HeaderFontStyle" VALUE=",,,,">
	<PARAM NAME="DetailFont" VALUE=",,,,">
	<PARAM NAME="DetailFontColor" VALUE=",,,,">
	<PARAM NAME="DetailFontSize" VALUE=",,,,">
	<PARAM NAME="DetailFontStyle" VALUE=",,,,">
	<PARAM NAME="ColumnCount" VALUE="5">
	<PARAM NAME="CurStyle" VALUE="Basic Navy">
	<PARAM NAME="TitleFont" VALUE="Arial">
	<PARAM NAME="titleFontSize" VALUE="2">
	<PARAM NAME="TitleFontColor" VALUE="16777215">
	<PARAM NAME="TitleBackColor" VALUE="32896">
	<PARAM NAME="TitleFontStyle" VALUE="1">
	<PARAM NAME="TitleAlignment" VALUE="0">
	<PARAM NAME="RowFont" VALUE="Arial">
	<PARAM NAME="RowFontColor" VALUE="0">
	<PARAM NAME="RowFontStyle" VALUE="0">
	<PARAM NAME="RowFontSize" VALUE="2">
	<PARAM NAME="RowBackColor" VALUE="16777215">
	<PARAM NAME="RowAlignment" VALUE="0">
	<PARAM NAME="HighlightColor3D" VALUE="268435455">
	<PARAM NAME="ShadowColor3D" VALUE="268435455">
	<PARAM NAME="PageSize" VALUE="20">
	<PARAM NAME="MoveFirstCaption" VALUE="    |<    ">
	<PARAM NAME="MoveLastCaption" VALUE="    >|    ">
	<PARAM NAME="MovePrevCaption" VALUE="    <<    ">
	<PARAM NAME="MoveNextCaption" VALUE="    >>    ">
	<PARAM NAME="BorderSize" VALUE="1">
	<PARAM NAME="BorderColor" VALUE="13421772">
	<PARAM NAME="GridBackColor" VALUE="8421504">
	<PARAM NAME="AltRowBckgnd" VALUE="12632256">
	<PARAM NAME="CellSpacing" VALUE="0">
	<PARAM NAME="WidthSelectionMode" VALUE="1">
	<PARAM NAME="GridWidth" VALUE="750">
	<PARAM NAME="EnablePaging" VALUE="-1">
	<PARAM NAME="ShowStatus" VALUE="-1">
	<PARAM NAME="StyleValue" VALUE="453613">
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
Grid1.tableAttributes = ' cellpadding=2  cellspacing=0 bordercolor=#cccccc bgcolor=Gray border=1 cols=5 rules=ALL WIDTH=750';
Grid1.headerAttributes = '   bgcolor=Teal align=Left';
Grid1.headerWidth[0] = ' WIDTH=90';
Grid1.headerWidth[1] = ' WIDTH=120';
Grid1.headerWidth[2] = ' WIDTH=270';
Grid1.headerWidth[3] = ' WIDTH=55';
Grid1.headerWidth[4] = ' WIDTH=35';
Grid1.headerFormat = '<Font face="Arial" size=2 color=White> <b>';
Grid1.colHeader[0] = '\'Name\'';
Grid1.colHeader[1] = '\'Subject\'';
Grid1.colHeader[2] = '\'Message\'';
Grid1.colHeader[3] = '\'Date\'';
Grid1.colHeader[4] = '\'Action\'';
Grid1.rowAttributes[0] = '  bgcolor = White align=Left bordercolor=#cccccc';
Grid1.rowAttributes[1] = '  bgcolor = Silver align=Left bordercolor=#cccccc';
Grid1.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid1.colAttributes[0] = '  WIDTH=90';
Grid1.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[0] = 'Recordset1.fields.getValue(\'FirstName\') + " " + Recordset1.fields.getValue(\'LastName\')';
Grid1.colAttributes[1] = '  WIDTH=120';
Grid1.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[1] = 'Recordset1.fields.getValue(\'MessageTitle\')';
Grid1.colAttributes[2] = '  WIDTH=270';
Grid1.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[2] = 'Recordset1.fields.getValue(\'MessageBody\')';
Grid1.colAttributes[3] = '  WIDTH=55';
Grid1.colFormat[3] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[3] = 'Recordset1.fields.getValue(\'MessageDate\')';
Grid1.colAttributes[4] = '  WIDTH=35';
Grid1.colFormat[4] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[4] = 'getbutton(Recordset1.fields.getValue(\'MessageID\'))';
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
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=rsUserRole style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12192">
	<PARAM NAME="ExtentY" VALUE="1799">
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

	' Allow a user to manage forums if they have been granted permssion to do so
	rsUserRole.SetSQLText("SELECT * From UserRoles Where UserID = "&session("UserID")& " AND RoleID=8")
	rsUserRole.Open
	If not rsUserRole.eof then
		Response.Write "<input type=button onclick=""window.location.href='ForumsAdd.asp'"" value=""Add a Forum"" id=button1 name=button1><br>"
		Response.Write "<input type=button onclick=""window.location.href='ForumsEditChoose.asp'"" value=""Edit/Delete Forums""><br>"
	End If
	rsUserRole.Close %>
		
	<p>
	Purge Date:
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" id=purgeDate style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="2963">
	<PARAM NAME="_ExtentY" VALUE="508">
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
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnPurge 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 108px" width=108>
	<PARAM NAME="_ExtentX" VALUE="2688">
	<PARAM NAME="_ExtentY" VALUE="699">
	<PARAM NAME="id" VALUE="btnPurge">
	<PARAM NAME="Caption" VALUE="Purge Posts">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnPurge()
{
	btnPurge.value = 'Purge Posts';
	btnPurge.setStyle(0);
}
function _btnPurge_ctor()
{
	CreateButton('btnPurge', _initbtnPurge, null);
}
</script>
<% btnPurge.display %>

<!--METADATA TYPE="DesignerControl" endspan-->

<p>&nbsp;</p>
<!-- #include File="../copyright.asp" -->
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</html>
<script runat=server language=vbscript>
function getbutton(ID)

	getbutton1 = "<input type=button value=" & chr(34) & "OK" & chr(34)
	getbutton2 = " onclick=" & chr(34) & "window.location.href='ForumsApprove.asp?action=approve&MessageID=" & cstr(ID) & "'" & chr(34) & " id=button1 name=button1>"

	getbutton = getbutton1 & getbutton2

End Function
</script>