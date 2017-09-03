<%@ Language=VBScript %>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

Sub ForumsAdd_onenter()
	if ForumsAdd.firstEntered then
		Recordset1.addRecord
	end if
End Sub

Sub btnSave_onclick()
	Recordset1.updateRecord
	Response.redirect "forums.asp"
End Sub

</SCRIPT>
</HEAD>
<body vLink="white" aLink="white" Link="white">
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:8CC35CD6-E98B-11D0-B218-00A0C92764F5" id=PageObject1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="4233">
	<PARAM NAME="ExtentY" VALUE="1508">
	<PARAM NAME="State" VALUE="(ObjectName_Unmatched=\qForumsAdd\q,NavigateMethods=(Rows=0),ExecuteMethods=(Rows=0),Properties=(Rows=0),References=(Rows=0))"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=SERVER>
/* VIPM PAGE DESCRIPTION
<DSC NAME="ForumsAdd">
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
if (typeof ForumsAdd_onbeforeserverevent == 'function' || typeof ForumsAdd_onbeforeserverevent == 'unknown')
	thisPage.advise('onbeforeserverevent', 'ForumsAdd_onbeforeserverevent()');

ForumsAdd = thisPage;
ForumsAdd.location = "../Admin/ForumsAdd.asp";
ForumsAdd.navigate = new Object;
ForumsAdd.navigate.show = Function('thisPage.invokeMethod("", "show", this.show.arguments);');
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

ForumsAdd = thisPage;
ForumsAdd.location = "../Admin/ForumsAdd.asp";
ForumsAdd.navigate = new Object;
ForumsAdd.navigate.show = Function('return;');

	thisPage._objEventManager.adviseDefaultHandler('ForumsAdd','onenter');
	thisPage._objEventManager.adviseDefaultHandler('ForumsAdd','onexit');
	thisPage._objEventManager.adviseDefaultHandler('ForumsAdd','onshow');
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
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject=\qTables\q,TCDBObjectName=\qConferences\q,TCControlID_Unmatched=\qRecordset1\q,TCPPConn=\qConnection1\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qConferences\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
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
	cmdTmp.CommandText = '`Conferences`';
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
<table border="1" width="750" cellpadding="3" cellspacing="0">
  <tr>
    <td width="180" bgcolor="#008080"><center><font face="Verdana"><big><strong><a href="../index.asp" target="_top" vlink="White">Home</a></strong></big></font></center></td>
    <td width="180" bgcolor="#008080"><center><font face="Verdana"><big><strong><a href="AdminSecure.asp" target="_top" vlink="White">Administration</a></strong></big></font></center></td>
    <td width="180" bgcolor="#008080"><center><font face="Verdana"><big><strong><a href="Forums.asp" target="_top" vlink="White">Moderate Forums</a></strong></big></font></center></td>
    <td><strong><font face="Verdana"><center>Manage Forums</center></font></strong></td>
  </tr>
</table>
<strong>
<br>

<br><br><table BORDER="0" CELLSPACING="1" CELLPADDING="1">
	<tr>
		<td><font face="Verdana" size="2">Conference 
Name</font> </td>
		<td><font face="Verdana" size="2">
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=Textbox1 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 300px" width=300>
	<PARAM NAME="_ExtentX" VALUE="7938">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox1">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="ConferenceName">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="50">
	<PARAM NAME="DisplayWidth" VALUE="50">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox1()
{
	Textbox1.setStyle(TXT_TEXTBOX);
	Textbox1.setDataSource(Recordset1);
	Textbox1.setDataField('ConferenceName');
	Textbox1.setMaxLength(50);
	Textbox1.setColumnCount(50);
}
function _Textbox1_ctor()
{
	CreateTextbox('Textbox1', _initTextbox1, null);
}
</script>
<% Textbox1.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
</td>
	</tr>
	<tr>
		<td><font face="Verdana" size="2">Conference 
      Description</font> </td>
		<td><font face="Verdana" size="2">
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=39 id=Textbox2 
	style="HEIGHT: 39px; LEFT: 0px; TOP: 0px; WIDTH: 300px" width=300>
	<PARAM NAME="_ExtentX" VALUE="7938">
	<PARAM NAME="_ExtentY" VALUE="1032">
	<PARAM NAME="id" VALUE="Textbox2">
	<PARAM NAME="ControlType" VALUE="1">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="ConferenceDescription">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="20">
	<PARAM NAME="DisplayWidth" VALUE="50">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox2()
{
	Textbox2.setStyle(TXT_TEXTAREA);
	Textbox2.setDataSource(Recordset1);
	Textbox2.setDataField('ConferenceDescription');
	Textbox2.setRowCount(3);
	Textbox2.setColumnCount(50);
}
function _Textbox2_ctor()
{
	CreateTextbox('Textbox2', _initTextbox2, null);
}
</script>
<% Textbox2.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
</td>
	</tr>
	<tr>
		<td><font face="Verdana" size="2">Moderated</font></td>
		<td><font face="Verdana" size="2">
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E46C-DC5F-11D0-9846-0000F8027CA0" height=27 id=Checkbox1 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 29px" width=29>
	<PARAM NAME="_ExtentX" VALUE="767">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="Checkbox1">
	<PARAM NAME="Caption" VALUE="">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="ConferenceModerated">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/CheckBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initCheckbox1()
{
	Checkbox1.setDataSource(Recordset1);
	Checkbox1.setDataField('ConferenceModerated');
}
function _Checkbox1_ctor()
{
	CreateCheckbox('Checkbox1', _initCheckbox1, null);
}
</script>
<% Checkbox1.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
</td>
	</tr> 
</table>

<br>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnSave style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 52px" 
	width=52>
	<PARAM NAME="_ExtentX" VALUE="1376">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnSave">
	<PARAM NAME="Caption" VALUE="Save">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnSave()
{
	btnSave.value = 'Save';
	btnSave.setStyle(0);
}
function _btnSave_ctor()
{
	CreateButton('btnSave', _initbtnSave, null);
}
</script>
<% btnSave.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>
