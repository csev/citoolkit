<%@ Language=VBScript %>
<% Response.expires=0 %>
<!-- #include File="../mapvar.asp" -->
<!--#include file="checksecure.asp"-->
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

Sub Recordset1_onbeforeupdate()
	Recordset1.fields.setValue "EditUserID",Cint(Session("UserID"))
	Recordset1.fields.setValue "EditDate",Cstr(Date())
End Sub


Sub EventEdit_onenter()
	Recordset1.setSQLText("SELECT * FROM EVENTS WHERE EventID="& cstr(Request.QueryString("EventID")))
	Recordset1.open
End Sub

Sub btnUpdate_onclick()
	Recordset1.updateRecord
End Sub

</SCRIPT>
</HEAD>

<body vLink=White aLink=White Link=White>

<table border="1" width="350" cellpadding=3 cellspacing=0>
  <tr>
	<td width=180 bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="../index.asp" target="_top" vlink="White">Home</a></strong></big></font></center></td>
    <td width=180 bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="AdminSecure.asp" target="_top" vlink="White">Administration</a></strong></big></font></center></td>
    <td bgcolor="#008080"><strong><font face="Verdana"><center><big><A HREF="Event.Asp">Events</A></big></center></font></strong></td>
    <td><strong><font face="Verdana"><center>Edit&nbsp;Event</center></font></strong></td>
  </tr>
</table><br>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:8CC35CD6-E98B-11D0-B218-00A0C92764F5" id=PageObject1 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="4233">
	<PARAM NAME="ExtentY" VALUE="1508">
	<PARAM NAME="State" VALUE="(ObjectName_Unmatched=\qEventEdit\q,NavigateMethods=(Rows=0),ExecuteMethods=(Rows=0),Properties=(Rows=0),References=(Rows=0))">
	
 
</OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=SERVER>
/* VIPM PAGE DESCRIPTION
<DSC NAME="EventEdit">
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
if (typeof EventEdit_onbeforeserverevent == 'function' || typeof EventEdit_onbeforeserverevent == 'unknown')
	thisPage.advise('onbeforeserverevent', 'EventEdit_onbeforeserverevent()');

EventEdit = thisPage;
EventEdit.location = "../Admin/EventEdit.asp";
EventEdit.navigate = new Object;
EventEdit.navigate.show = Function('thisPage.invokeMethod("", "show", this.show.arguments);');
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

EventEdit = thisPage;
EventEdit.location = "../Admin/EventEdit.asp";
EventEdit.navigate = new Object;
EventEdit.navigate.show = Function('return;');

	thisPage._objEventManager.adviseDefaultHandler('EventEdit','onenter');
	thisPage._objEventManager.adviseDefaultHandler('EventEdit','onexit');
	thisPage._objEventManager.adviseDefaultHandler('EventEdit','onshow');
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Recordset1 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject=\qTables\q,TCDBObjectName=\qEvents\q,TCControlID_Unmatched=\qRecordset1\q,TCPPConn=\qConnection1\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qEvents\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
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
	cmdTmp.CommandText = '`Events`';
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
<FONT face=Verdana></FONT>
<TABLE CellPadding=1 CellSpacing=2 Cols=2>
<TR>
	<TD><FONT face=Verdana>Event Date </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" 
            id=Textbox1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox1">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="EventDate">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="19">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox1()
{
	Textbox1.setStyle(TXT_TEXTBOX);
	Textbox1.setDataSource(Recordset1);
	Textbox1.setDataField('EventDate');
	Textbox1.setMaxLength(19);
	Textbox1.setColumnCount(20);
}
function _Textbox1_ctor()
{
	CreateTextbox('Textbox1', _initTextbox1, null);
}
</script>
<% Textbox1.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD>
<TR>
	<TD><FONT face=Verdana>Event Name </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=Textbox2 style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" 
            width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox2">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="EventName">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="50">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox2()
{
	Textbox2.setStyle(TXT_TEXTBOX);
	Textbox2.setDataSource(Recordset1);
	Textbox2.setDataField('EventName');
	Textbox2.setMaxLength(50);
	Textbox2.setColumnCount(30);
}
function _Textbox2_ctor()
{
	CreateTextbox('Textbox2', _initTextbox2, null);
}
</script>
<% Textbox2.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD>
<TR>
	<TD><FONT face=Verdana>Event Description </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=39 
            id=Textbox3 style="HEIGHT: 39px; LEFT: 0px; TOP: 0px; WIDTH: 240px" 
            width=240>
	<PARAM NAME="_ExtentX" VALUE="6350">
	<PARAM NAME="_ExtentY" VALUE="1032">
	<PARAM NAME="id" VALUE="Textbox3">
	<PARAM NAME="ControlType" VALUE="1">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="EventDescription">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="20">
	<PARAM NAME="DisplayWidth" VALUE="40">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox3()
{
	Textbox3.setStyle(TXT_TEXTAREA);
	Textbox3.setDataSource(Recordset1);
	Textbox3.setDataField('EventDescription');
	Textbox3.setRowCount(3);
	Textbox3.setColumnCount(40);
}
function _Textbox3_ctor()
{
	CreateTextbox('Textbox3', _initTextbox3, null);
}
</script>
<% Textbox3.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT>
</TD>
</TR>
<TR>
	<TD><FONT face=Verdana>Event Time Start </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" 
            id=Textbox4 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox4">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="EventTimeStart">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="20">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox4()
{
	Textbox4.setStyle(TXT_TEXTBOX);
	Textbox4.setDataSource(Recordset1);
	Textbox4.setDataField('EventTimeStart');
	Textbox4.setMaxLength(20);
	Textbox4.setColumnCount(20);
}
function _Textbox4_ctor()
{
	CreateTextbox('Textbox4', _initTextbox4, null);
}
</script>
<% Textbox4.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT>
</TD>
</TR>
<TR>
	<TD><FONT face=Verdana>Event Time End </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" 
            id=Textbox5 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox5">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="EventTimeEnd">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="20">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox5()
{
	Textbox5.setStyle(TXT_TEXTBOX);
	Textbox5.setDataSource(Recordset1);
	Textbox5.setDataField('EventTimeEnd');
	Textbox5.setMaxLength(20);
	Textbox5.setColumnCount(20);
}
function _Textbox5_ctor()
{
	CreateTextbox('Textbox5', _initTextbox5, null);
}
</script>
<% Textbox5.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT>
</TD>
</TR>
<TR>
	<TD><FONT face=Verdana>Event Publish Date 
</FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" 
            id=Textbox6 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox6">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="EventPublishDate">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="19">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox6()
{
	Textbox6.setStyle(TXT_TEXTBOX);
	Textbox6.setDataSource(Recordset1);
	Textbox6.setDataField('EventPublishDate');
	Textbox6.setMaxLength(19);
	Textbox6.setColumnCount(20);
}
function _Textbox6_ctor()
{
	CreateTextbox('Textbox6', _initTextbox6, null);
}
</script>
<% Textbox6.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT>
</TD>
</TR>
<TR>
	<TD><FONT face=Verdana>Location </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=Textbox7 style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" 
            width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox7">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="Location">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="255">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox7()
{
	Textbox7.setStyle(TXT_TEXTBOX);
	Textbox7.setDataSource(Recordset1);
	Textbox7.setDataField('Location');
	Textbox7.setMaxLength(255);
	Textbox7.setColumnCount(30);
}
function _Textbox7_ctor()
{
	CreateTextbox('Textbox7', _initTextbox7, null);
}
</script>
<% Textbox7.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT>
</TD>
</TR>
<TR>
	<TD><FONT face=Verdana>Contact </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" 
            id=Textbox8 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox8">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="Contact">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="50">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox8()
{
	Textbox8.setStyle(TXT_TEXTBOX);
	Textbox8.setDataSource(Recordset1);
	Textbox8.setDataField('Contact');
	Textbox8.setMaxLength(50);
	Textbox8.setColumnCount(20);
}
function _Textbox8_ctor()
{
	CreateTextbox('Textbox8', _initTextbox8, null);
}
</script>
<% Textbox8.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT>
</TD>
</TR>
<TR>
	<TD><FONT face=Verdana>Contact Phone </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" 
            id=Textbox9 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox9">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="ContactPhone">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="50">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox9()
{
	Textbox9.setStyle(TXT_TEXTBOX);
	Textbox9.setDataSource(Recordset1);
	Textbox9.setDataField('ContactPhone');
	Textbox9.setMaxLength(50);
	Textbox9.setColumnCount(20);
}
function _Textbox9_ctor()
{
	CreateTextbox('Textbox9', _initTextbox9, null);
}
</script>
<% Textbox9.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT>
</TD>
</TR>
<TR>
	<TD><FONT face=Verdana>Contact Email </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=Textbox10 style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" 
            width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox10">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="ContactEmail">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="50">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox10()
{
	Textbox10.setStyle(TXT_TEXTBOX);
	Textbox10.setDataSource(Recordset1);
	Textbox10.setDataField('ContactEmail');
	Textbox10.setMaxLength(50);
	Textbox10.setColumnCount(30);
}
function _Textbox10_ctor()
{
	CreateTextbox('Textbox10', _initTextbox10, null);
}
</script>
<% Textbox10.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT>
</TD>
</TR>
<TR>
	<TD><FONT face=Verdana>Admission Charge? </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E46C-DC5F-11D0-9846-0000F8027CA0" height=27 
            id=Checkbox1 
            style="HEIGHT: 27px; LEFT: 10px; TOP: 323px; WIDTH: 29px" width=29>
	<PARAM NAME="_ExtentX" VALUE="767">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="Checkbox1">
	<PARAM NAME="Caption" VALUE="">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="AdmissionCharge">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             
            </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/CheckBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initCheckbox1()
{
	Checkbox1.setDataSource(Recordset1);
	Checkbox1.setDataField('AdmissionCharge');
}
function _Checkbox1_ctor()
{
	CreateCheckbox('Checkbox1', _initCheckbox1, null);
}
</script>
<% Checkbox1.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT>
</TD>
</TR>
<TR>
	<TD><FONT face=Verdana>Admission Amount </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" 
            id=Textbox11 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox11">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="AdmissionAmount">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="50">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox11()
{
	Textbox11.setStyle(TXT_TEXTBOX);
	Textbox11.setDataSource(Recordset1);
	Textbox11.setDataField('AdmissionAmount');
	Textbox11.setMaxLength(50);
	Textbox11.setColumnCount(20);
}
function _Textbox11_ctor()
{
	CreateTextbox('Textbox11', _initTextbox11, null);
}
</script>
<% Textbox11.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT>
</TD>
</TR>
<TR>
	<TD><FONT face=Verdana>Text Directions </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=39 
            id=Textbox12 style="HEIGHT: 39px; LEFT: 0px; TOP: 0px; WIDTH: 240px" 
            width=240>
	<PARAM NAME="_ExtentX" VALUE="6350">
	<PARAM NAME="_ExtentY" VALUE="1032">
	<PARAM NAME="id" VALUE="Textbox12">
	<PARAM NAME="ControlType" VALUE="1">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="TextDirections">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="255">
	<PARAM NAME="DisplayWidth" VALUE="40">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox12()
{
	Textbox12.setStyle(TXT_TEXTAREA);
	Textbox12.setDataSource(Recordset1);
	Textbox12.setDataField('TextDirections');
	Textbox12.setRowCount(3);
	Textbox12.setColumnCount(40);
}
function _Textbox12_ctor()
{
	CreateTextbox('Textbox12', _initTextbox12, null);
}
</script>
<% Textbox12.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT>
</TD>
</TR>
<TR>
	<TD><FONT face=Verdana>URL to Map </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=Textbox13 style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" 
            width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox13">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="URLMap">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="50">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox13()
{
	Textbox13.setStyle(TXT_TEXTBOX);
	Textbox13.setDataSource(Recordset1);
	Textbox13.setDataField('URLMap');
	Textbox13.setMaxLength(50);
	Textbox13.setColumnCount(30);
}
function _Textbox13_ctor()
{
	CreateTextbox('Textbox13', _initTextbox13, null);
}
</script>
<% Textbox13.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT>
</TD>
</TR>
<TR>
	<TD><FONT face=Verdana>URL to Info </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=Textbox14 style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" 
            width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox14">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="URLInfo">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="50">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox14()
{
	Textbox14.setStyle(TXT_TEXTBOX);
	Textbox14.setDataSource(Recordset1);
	Textbox14.setDataField('URLInfo');
	Textbox14.setMaxLength(50);
	Textbox14.setColumnCount(30);
}
function _Textbox14_ctor()
{
	CreateTextbox('Textbox14', _initTextbox14, null);
}
</script>
<% Textbox14.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT>
</TD>
</TR>
<TR>
	<TD><FONT face=Verdana>URL Name for Info </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=Textbox15 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox15">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="URLName">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="50">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox15()
{
	Textbox15.setStyle(TXT_TEXTBOX);
	Textbox15.setDataSource(Recordset1);
	Textbox15.setDataField('URLName');
	Textbox15.setMaxLength(50);
	Textbox15.setColumnCount(30);
}
function _Textbox15_ctor()
{
	CreateTextbox('Textbox15', _initTextbox15, null);
}
</script>
<% Textbox15.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT>
</TD>
</TR>
<TR>
	<TD><FONT face=Verdana>URL to Photo </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 
            id=Textbox16 style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 180px" 
            width=180>
	<PARAM NAME="_ExtentX" VALUE="4763">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox16">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="URLPhoto">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="50">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox16()
{
	Textbox16.setStyle(TXT_TEXTBOX);
	Textbox16.setDataSource(Recordset1);
	Textbox16.setDataField('URLPhoto');
	Textbox16.setMaxLength(50);
	Textbox16.setColumnCount(30);
}
function _Textbox16_ctor()
{
	CreateTextbox('Textbox16', _initTextbox16, null);
}
</script>
<% Textbox16.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD>
<TR>
	<TD><FONT face=Verdana>Event Type </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 id=Listbox1 
	style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 96px" width=96>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="Listbox1">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="EventTypeID">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="0">
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
	Listbox1.setDataSource(Recordset1);
	Listbox1.setDataField('EventTypeID');
}
function _Listbox1_ctor()
{
	CreateListbox('Listbox1', _initListbox1, null);
}
</script>
<% Listbox1.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT>
</TD>
</TR>
<TR>
	<TD><FONT face=Verdana>Audience Type </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 id=Listbox2 
	style="HEIGHT: 21px; LEFT: 0px; TOP: 0px; WIDTH: 96px" width=96>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="Listbox2">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="AudienceID">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="UsesStaticList" VALUE="0">
	<PARAM NAME="RowSource" VALUE="Recordset3">
	<PARAM NAME="BoundColumn" VALUE="AudienceID">
	<PARAM NAME="ListField" VALUE="AudienceName">
	<PARAM NAME="LookupPlatform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initListbox2()
{
	Recordset3.advise(RS_ONDATASETCOMPLETE, 'Listbox2.setRowSource(Recordset3, \'AudienceName\', \'AudienceID\');');
	Listbox2.setDataSource(Recordset1);
	Listbox2.setDataField('AudienceID');
}
function _Listbox2_ctor()
{
	CreateListbox('Listbox2', _initListbox2, null);
}
</script>
<% Listbox2.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT>
</TD>
</TR>
    <TR>
        <TD><FONT face=Verdana>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=Label1 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 79px" 
	width=79>
	<PARAM NAME="_ExtentX" VALUE="2090">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="Label1">
	<PARAM NAME="DataSource" VALUE="rsSiteInfo">
	<PARAM NAME="DataField" VALUE="UserField1Label">
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
	Label1.setDataSource(rsSiteInfo);
	Label1.setDataField('UserField1Label');
}
function _Label1_ctor()
{
	CreateLabel('Label1', _initLabel1, null);
}
</script>
<% Label1.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
&nbsp; </FONT>
        <TD><FONT face=Verdana>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" id=Textbox17 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox17">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="UserField1">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="20">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox17()
{
	Textbox17.setStyle(TXT_TEXTBOX);
	Textbox17.setDataSource(Recordset1);
	Textbox17.setDataField('UserField1');
	Textbox17.setMaxLength(20);
	Textbox17.setColumnCount(20);
}
function _Textbox17_ctor()
{
	CreateTextbox('Textbox17', _initTextbox17, null);
}
</script>
<% Textbox17.display %>

<!--METADATA TYPE="DesignerControl" endspan-->            
</FONT></TD></TR>
    <TR>
        <TD><FONT face=Verdana>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=Label2 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 79px" 
	width=79>
	<PARAM NAME="_ExtentX" VALUE="2090">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="Label2">
	<PARAM NAME="DataSource" VALUE="rsSiteInfo">
	<PARAM NAME="DataField" VALUE="UserField2Label">
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
	Label2.setDataSource(rsSiteInfo);
	Label2.setDataField('UserField2Label');
}
function _Label2_ctor()
{
	CreateLabel('Label2', _initLabel2, null);
}
</script>
<% Label2.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
&nbsp; </FONT>
        <TD><FONT face=Verdana>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" id=Textbox18 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox18">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="UserField2">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="20">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox18()
{
	Textbox18.setStyle(TXT_TEXTBOX);
	Textbox18.setDataSource(Recordset1);
	Textbox18.setDataField('UserField2');
	Textbox18.setMaxLength(20);
	Textbox18.setColumnCount(20);
}
function _Textbox18_ctor()
{
	CreateTextbox('Textbox18', _initTextbox18, null);
}
</script>
<% Textbox18.display %>

<!--METADATA TYPE="DesignerControl" endspan-->            
</FONT></TD></TR>
    <TR>
        <TD><FONT face=Verdana>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=Label3 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 79px" 
	width=79>
	<PARAM NAME="_ExtentX" VALUE="2090">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="Label3">
	<PARAM NAME="DataSource" VALUE="rsSiteInfo">
	<PARAM NAME="DataField" VALUE="UserField3Label">
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
	Label3.setDataSource(rsSiteInfo);
	Label3.setDataField('UserField3Label');
}
function _Label3_ctor()
{
	CreateLabel('Label3', _initLabel3, null);
}
</script>
<% Label3.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
&nbsp; </FONT>
        <TD><FONT face=Verdana>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" id=Textbox19 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox19">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="UserField3">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="20">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox19()
{
	Textbox19.setStyle(TXT_TEXTBOX);
	Textbox19.setDataSource(Recordset1);
	Textbox19.setDataField('UserField3');
	Textbox19.setMaxLength(20);
	Textbox19.setColumnCount(20);
}
function _Textbox19_ctor()
{
	CreateTextbox('Textbox19', _initTextbox19, null);
}
</script>
<% Textbox19.display %>

<!--METADATA TYPE="DesignerControl" endspan-->            
</FONT></TD></TR>
</TABLE><FONT face=Verdana>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Recordset2 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject=\qTables\q,TCDBObjectName=\qEventType\q,TCControlID_Unmatched=\qRecordset2\q,TCPPConn=\qConnection1\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qEventType\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" height=79 
id=Recordset3 style="HEIGHT: 79px; LEFT: 0px; TOP: 0px; WIDTH: 461px" width=461>
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject=\qTables\q,TCDBObjectName=\qAudience\q,TCControlID_Unmatched=\qRecordset3\q,TCPPConn=\qConnection1\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qAudience\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
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
	cmdTmp.CommandText = '`Audience`';
	rsTmp.CacheSize = 10;
	rsTmp.CursorType = 3;
	rsTmp.CursorLocation = 3;
	rsTmp.LockType = 3;
	Recordset3.setRecordSource(rsTmp);
	Recordset3.open();
	if (thisPage.getState('pb_Recordset3') != null)
		Recordset3.setBookmark(thisPage.getState('pb_Recordset3'));
}
function _Recordset3_ctor()
{
	CreateRecordset('Recordset3', _initRecordset3, null);
}
function _Recordset3_dtor()
{
	Recordset3._preserveState();
	thisPage.setState('pb_Recordset3', Recordset3.getBookmark());
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
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" id=btnUpdate 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="1773">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnUpdate">
	<PARAM NAME="Caption" VALUE="Update">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnUpdate()
{
	btnUpdate.value = 'Update';
	btnUpdate.setStyle(0);
}
function _btnUpdate_ctor()
{
	CreateButton('btnUpdate', _initbtnUpdate, null);
}
</script>
<% btnUpdate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT>

<P>&nbsp;</P>
<!-- #include File="../copyright.asp" -->
</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>
