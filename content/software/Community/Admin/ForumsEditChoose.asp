<%@ Language=VBScript%>
<% Response.Expires=0 %>
<!-- #include File="../mapvar.asp" -->
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
<head>
<title>Admin-Forums</title>
</head>

<body vLink=white aLink=white Link=white>
<table border="1" cellpadding=3 cellspacing=0>
  <tr>
    <td bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="../index.asp" target=_top  vlink="White">Home</A></strong></big></font></center></td>
    <td bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="AdminSecure.asp" target=_top  vlink="White">Administration</A></strong></big></font></center></td>
    <td bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="Forums.asp" target=_top  vlink="White">Manage Forums</A></strong></big></font></center></td>
    <td><strong><font face="Verdana"><center>Edit</center></font></strong></td>
  </tr>
</table>
<strong>
<br>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:8CC35CD6-E98B-11D0-B218-00A0C92764F5" id=PageObject1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="3387">
	<PARAM NAME="ExtentY" VALUE="1270">
	<PARAM NAME="State" VALUE="(ObjectName_Unmatched=\qForumsEditChoose\q,NavigateMethods=(Rows=0),ExecuteMethods=(Rows=0),Properties=(Rows=0),References=(Rows=0))"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=SERVER>
/* VIPM PAGE DESCRIPTION
<DSC NAME="ForumsEditChoose">
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
if (typeof ForumsEditChoose_onbeforeserverevent == 'function' || typeof ForumsEditChoose_onbeforeserverevent == 'unknown')
	thisPage.advise('onbeforeserverevent', 'ForumsEditChoose_onbeforeserverevent()');

ForumsEditChoose = thisPage;
ForumsEditChoose.location = "../Admin/ForumsEditChoose.asp";
ForumsEditChoose.navigate = new Object;
ForumsEditChoose.navigate.show = Function('thisPage.invokeMethod("", "show", this.show.arguments);');
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

ForumsEditChoose = thisPage;
ForumsEditChoose.location = "../Admin/ForumsEditChoose.asp";
ForumsEditChoose.navigate = new Object;
ForumsEditChoose.navigate.show = Function('return;');

	thisPage._objEventManager.adviseDefaultHandler('ForumsEditChoose','onenter');
	thisPage._objEventManager.adviseDefaultHandler('ForumsEditChoose','onexit');
	thisPage._objEventManager.adviseDefaultHandler('ForumsEditChoose','onshow');
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=rsForums style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12192">
	<PARAM NAME="ExtentY" VALUE="1799">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject=\qTables\q,TCDBObjectName=\qConferences\q,TCControlID_Unmatched=\qrsForums\q,TCPPConn=\qConnection1\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qConferences\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCNoCache\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initrsForums()
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
	cmdTmp.CommandText = '`Conferences`';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	rsForums.setRecordSource(rsTmp);
	rsForums.open();
}
function _rsForums_ctor()
{
	CreateRecordset('rsForums', _initrsForums, null);
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->

<br>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:277FC3F2-E90F-11D0-B767-0000F81E081D" height=184 id=Grid1 style="HEIGHT: 184px; LEFT: 0px; TOP: 0px; WIDTH: 667px" 
	width=667>
	<PARAM NAME="_ExtentX" VALUE="14118">
	<PARAM NAME="_ExtentY" VALUE="3895">
	<PARAM NAME="DataConnection" VALUE="">
	<PARAM NAME="SourceType" VALUE="">
	<PARAM NAME="Recordset" VALUE="rsForums">
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
	<PARAM NAME="ColumnsNames" VALUE='"ConferenceName","ConferenceDescription","ConferenceModerated","=getEditButton([ConferenceID]) + getDeleteButton([ConferenceID]) "'>
	<PARAM NAME="columnIndex" VALUE="0,1,2,3">
	<PARAM NAME="displayWidth" VALUE="110,278,68,123">
	<PARAM NAME="Coltype" VALUE="1,1,1,1">
	<PARAM NAME="formated" VALUE="0,0,0,0">
	<PARAM NAME="DisplayName" VALUE='"Name","Description","Moderated","Action"'>
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
	<PARAM NAME="titleFontSize" VALUE="2">
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
	<PARAM NAME="GridWidth" VALUE="667">
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
Grid1.setDataSource(rsForums);
Grid1.tableAttributes = ' cellpadding=2  cellspacing=1 bordercolordark=Gray bordercolorlight=Silver bgcolor=Silver border=2 cols=4 rules=ROWS WIDTH=667';
Grid1.headerAttributes = '   bgcolor=Teal align=Left';
Grid1.headerWidth[0] = ' WIDTH=110';
Grid1.headerWidth[1] = ' WIDTH=278';
Grid1.headerWidth[2] = ' WIDTH=68';
Grid1.headerWidth[3] = ' WIDTH=123';
Grid1.headerFormat = '<Font face="Arial" size=2 color=Black> <b>';
Grid1.colHeader[0] = '\'Name\'';
Grid1.colHeader[1] = '\'Description\'';
Grid1.colHeader[2] = '\'Moderated\'';
Grid1.colHeader[3] = '\'Action\'';
Grid1.rowAttributes[0] = '  bgcolor = White align=Left bordercolordark=Gray bordercolorlight=Silver';
Grid1.rowFormat[0] = ' <Font face="Arial" size=2 color=Black >';
Grid1.colAttributes[0] = '  WIDTH=110';
Grid1.colFormat[0] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[0] = 'rsForums.fields.getValue(\'ConferenceName\')';
Grid1.colAttributes[1] = '  WIDTH=278';
Grid1.colFormat[1] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[1] = 'rsForums.fields.getValue(\'ConferenceDescription\')';
Grid1.colAttributes[2] = '  WIDTH=68';
Grid1.colFormat[2] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[2] = 'rsForums.fields.getValue(\'ConferenceModerated\')';
Grid1.colAttributes[3] = '  WIDTH=123';
Grid1.colFormat[3] = '<Font Size=2 Face="Arial" Color=Black >';
Grid1.colData[3] = 'getEditButton(rsForums.fields.getValue(\'ConferenceID\')) + getDeleteButton(rsForums.fields.getValue(\'ConferenceID\')) ';
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

<p>

<!-- #include File="../copyright.asp" --></p></strong>
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</html>

<script runat=server language=vbscript>
function getEditButton(ID)

	getbutton1 = "<input type=button value=" & chr(34) & "Edit" & chr(34)
	getbutton2 = " onclick=" & chr(34) & "window.location.href='ForumsEdit.asp?ForumID=" & cstr(ID) & "'" & chr(34) & ">"

	getEditbutton = getbutton1 & getbutton2

End Function

function getDeleteButton(ID)

	getbutton1 = "<input type=button value=" & chr(34) & "Delete" & chr(34)
	getbutton2 = " onclick=" & chr(34) & "window.location.href='ForumsDelete.asp?ForumID=" & cstr(ID) & "'" & chr(34) & " id=button1 name=button1>"

	getDeletebutton = getbutton1 & getbutton2
End Function


</script>