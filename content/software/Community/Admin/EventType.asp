<%@ Language=VBScript%>
<!-- #include File="../mapvar.asp" -->
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
<html>

<head>
<title>Admin-Event Types</title>
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

Sub Button1_onclick()
	Recordset1.moveLast
	Recordset1.addRecord
End Sub

Sub btn_delete_onclick()
	Recordset1.deleteRecord
	If Recordset1.EOF then
		Recordset1.moveprevious
	Else
		Recordset1.movenext
	End If
End Sub

Sub btn_go_onclick()
	Recordset1.moveAbsolute(cint(Request.Form("listbox1")))
End Sub


</SCRIPT>
</head>

<body>

<table border="0" width="100">
  <tr>
    <td width="100%" bgcolor="#008080"><font face="Verdana" color="#FFFFFF"><big><strong>Administration</strong></big></font></td>
  </tr>
</table>
<strong><font face="Verdana">
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Recordset1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject=\qTables\q,TCDBObjectName=\qEventType\q,TCControlID_Unmatched=\qRecordset1\q,TCPPConn=\qConnection1\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qEventType\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
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
	cmdTmp.CommandText = '`EventType`';
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
<p>Event Types</font></strong> </p>
<A href="AdminSecure.asp" target="_top" ><IMG alt="" border=0 height=15 src="../images/bullet_arrow.gif" width=20 >
<FONT face=Verdana>Administration Menu</font></a>
<br><br>

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Recordset2 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject=\qTables\q,TCDBObjectName=\qEventType\q,TCControlID_Unmatched=\qRecordset2\q,TCPPConn=\qConnection1\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qEventType\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
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
	cmdTmp.CommandText = '`EventType`';
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
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 id=Listbox1 
	style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 96px" width=96>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="Listbox1">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="0">
	<PARAM NAME="RowSource" VALUE="Recordset2">
	<PARAM NAME="BoundColumn" VALUE="EventTypeID">
	<PARAM NAME="ListField" VALUE="EventTypeName">
	<PARAM NAME="LookupPlatform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/ListBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initListbox1()
{
	Recordset2.advise(RS_ONDATASETCOMPLETE, 'Listbox1.setRowSource(Recordset2, \'EventTypeName\', \'EventTypeID\');');
}
function _Listbox1_ctor()
{
	CreateListbox('Listbox1', _initListbox1, null);
}
</script>
<% Listbox1.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btn_go style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 40px" 
	width=40>
	<PARAM NAME="_ExtentX" VALUE="1058">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btn_go">
	<PARAM NAME="Caption" VALUE="Go!">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtn_go()
{
	btn_go.value = 'Go!';
	btn_go.setStyle(0);
}
function _btn_go_ctor()
{
	CreateButton('btn_go', _initbtn_go, null);
}
</script>
<% btn_go.display %>

<!--METADATA TYPE="DesignerControl" endspan-->


  <table border="0" bordercolorlight="#000000">
    <tr>
      <td><font face="Verdana">Event Type Name</font></td>
      <td ><font face="Verdana">
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=Textbox1 
	style="HEIGHT: 19px; LEFT: 10px; TOP: 113px; WIDTH: 210px" width=210>
	<PARAM NAME="_ExtentX" VALUE="5556">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox1">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="EventTypeName">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="50">
	<PARAM NAME="DisplayWidth" VALUE="35">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox1()
{
	Textbox1.setStyle(TXT_TEXTBOX);
	Textbox1.setDataSource(Recordset1);
	Textbox1.setDataField('EventTypeName');
	Textbox1.setMaxLength(50);
	Textbox1.setColumnCount(35);
}
function _Textbox1_ctor()
{
	CreateTextbox('Textbox1', _initTextbox1, null);
}
</script>
<% Textbox1.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</font></td>
    </tr>
    <tr>
      <td><font face="Verdana">Event Type Image URL</font></td>
      <td><font face="Verdana">
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=Textbox2 
	style="HEIGHT: 19px; LEFT: 10px; TOP: 132px; WIDTH: 210px" width=210>
	<PARAM NAME="_ExtentX" VALUE="5556">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox2">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="EventTypeImageURL">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="50">
	<PARAM NAME="DisplayWidth" VALUE="35">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox2()
{
	Textbox2.setStyle(TXT_TEXTBOX);
	Textbox2.setDataSource(Recordset1);
	Textbox2.setDataField('EventTypeImageURL');
	Textbox2.setMaxLength(50);
	Textbox2.setColumnCount(35);
}
function _Textbox2_ctor()
{
	CreateTextbox('Textbox2', _initTextbox2, null);
}
</script>
<% Textbox2.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</font></td>
    </tr>
    <tr>
      <td width="60%"><font face="Verdana">Image</font></td>
      <td width="50%"><% 
	If Textbox2.value <> "" then
		Response.Write "<img src=" & chr(34) & Textbox2.value & chr(34) & ">"
	End If
%></td>
    </tr>
  </table>
  

<p>
  <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:58F3D268-FEDF-11D0-9C7F-0060081840F3" id=RecordsetNavbar1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="4075">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="RecordsetNavbar1">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="UpdateOnMove" VALUE="-1">
	<PARAM NAME="FirstCaption" VALUE=" |< ">
	<PARAM NAME="MoveFirst" VALUE="-1">
	<PARAM NAME="FirstImage" VALUE="0">
	<PARAM NAME="PrevCaption" VALUE="  <  ">
	<PARAM NAME="MovePrev" VALUE="-1">
	<PARAM NAME="PrevImage" VALUE="0">
	<PARAM NAME="NextCaption" VALUE="  >  ">
	<PARAM NAME="MoveNext" VALUE="-1">
	<PARAM NAME="NextImage" VALUE="0">
	<PARAM NAME="LastCaption" VALUE=" >| ">
	<PARAM NAME="MoveLast" VALUE="-1">
	<PARAM NAME="LastImage" VALUE="0">
	<PARAM NAME="Alignment" VALUE="1">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/RSNavBar.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initRecordsetNavbar1()
{
	RecordsetNavbar1.setAlignment(1);
	RecordsetNavbar1.setButtonStyles(170);
	RecordsetNavbar1.setDataSource(Recordset1);
	RecordsetNavbar1.getButton(0).value = ' |< ';
	RecordsetNavbar1.getButton(1).value = '  <  ';
	RecordsetNavbar1.getButton(2).value = '  >  ';
	RecordsetNavbar1.getButton(3).value = ' >| ';
}
function _RecordsetNavbar1_ctor()
{
	CreateRecordsetNavbar('RecordsetNavbar1', _initRecordsetNavbar1, null);
}
</script>
<% RecordsetNavbar1.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=Button1 style="HEIGHT: 27px; LEFT: 10px; TOP: 178px; WIDTH: 46px" 
	width=46>
	<PARAM NAME="_ExtentX" VALUE="1217">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="Button1">
	<PARAM NAME="Caption" VALUE="New">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initButton1()
{
	Button1.value = 'New';
	Button1.setStyle(0);
}
function _Button1_ctor()
{
	CreateButton('Button1', _initButton1, null);
}
</script>
<% Button1.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btn_delete 
	style="HEIGHT: 27px; LEFT: 10px; TOP: 205px; WIDTH: 62px" width=62>
	<PARAM NAME="_ExtentX" VALUE="1640">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btn_delete">
	<PARAM NAME="Caption" VALUE="Delete">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtn_delete()
{
	btn_delete.value = 'Delete';
	btn_delete.setStyle(0);
}
function _btn_delete_ctor()
{
	CreateButton('btn_delete', _initbtn_delete, null);
}
</script>
<% btn_delete.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<!-- #include File="../copyright.asp" -->
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</html>
