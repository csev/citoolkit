<%@ Language=VBScript%>
<% Response.buffer = true %>
<!-- #include File="../adovbs.inc" -->
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<% Response.Expires=0 %>
<FORM name=thisForm METHOD=post>
<html>
<body  vlink=White   link=White alink=#008080>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:8CC35CD6-E98B-11D0-B218-00A0C92764F5" id=PageObject1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="4233">
	<PARAM NAME="ExtentY" VALUE="1508">
	<PARAM NAME="State" VALUE="(ObjectName_Unmatched=\qUsers\q,NavigateMethods=(Rows=0),ExecuteMethods=(Rows=0),Properties=(Rows=0),References=(Rows=0))"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=SERVER>
/* VIPM PAGE DESCRIPTION
<DSC NAME="Users">
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
if (typeof Users_onbeforeserverevent == 'function' || typeof Users_onbeforeserverevent == 'unknown')
	thisPage.advise('onbeforeserverevent', 'Users_onbeforeserverevent()');

Users = thisPage;
Users.location = "../Admin/Users.asp";
Users.navigate = new Object;
Users.navigate.show = Function('thisPage.invokeMethod("", "show", this.show.arguments);');
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

Users = thisPage;
Users.location = "../Admin/Users.asp";
Users.navigate = new Object;
Users.navigate.show = Function('return;');

	thisPage._objEventManager.adviseDefaultHandler('Users','onenter');
	thisPage._objEventManager.adviseDefaultHandler('Users','onexit');
	thisPage._objEventManager.adviseDefaultHandler('Users','onshow');
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
<title>Admin-User Menu</title>
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

Sub Button1_onclick()
	Recordset1.moveLast
	Recordset1.AddRecord
End Sub

Sub Recordset1_onbeforeupdate()
	If Request.Form("password") <> "" Then
		If Request.Form("password") = Request.Form("verifypassword") then
			recordset1.fields.setValue "Password", Request.Form("password")
		Else
			Response.Write "<font size=+2 face=verdana><b>Passwords did not match</b></font>"
			recordset1.cancelUpdate
		End If
	End If
		
End Sub

Sub btnUpdate_onclick()
	recordset1.updateRecord
	Response.Redirect "UsersComplete.asp"
End Sub

Sub btnGoto_onclick()
	recordset1.close
	recordset1.setSQLText("SELECT * FROM Users WHERE UserID="&Request.Form("lstUsers"))
	recordset1.open
End Sub

Sub btnDelete_onclick()
	recordset1.deleteRecord
	Response.Redirect "UsersComplete.asp"
End Sub

Sub RecordsetNavbar1_onnextclick()
	rsUsers.requery
End Sub

Sub RecordsetNavbar1_onfirstclick()
	rsUsers.requery
End Sub

Sub RecordsetNavbar1_onlastclick()
	rsUsers.requery
End Sub

Sub RecordsetNavbar1_onpreviousclick()
	rsUsers.requery
End Sub

Sub Users_onenter()
	rsUsers.close
	rsUsers.open
	rsUsers.requery
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
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Recordset1 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject=\qTables\q,TCDBObjectName=\qUsers\q,TCControlID_Unmatched=\qRecordset1\q,TCPPConn=\qConnection1\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qUsers\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../">
	
 </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initRecordset1()
{
	var DBConn = Server.CreateObject('ADODB.Connection');
	DBConn.ConnectionTimeout = session('Connection1_ConnectionTimeout');
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
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=rsUsers style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject=\qTables\q,TCDBObjectName=\qUsers\q,TCControlID_Unmatched=\qrsUsers\q,TCPPConn=\qConnection1\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qUsers\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initrsUsers()
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
	rsUsers.setRecordSource(rsTmp);
	rsUsers.open();
	if (thisPage.getState('pb_rsUsers') != null)
		rsUsers.setBookmark(thisPage.getState('pb_rsUsers'));
}
function _rsUsers_ctor()
{
	CreateRecordset('rsUsers', _initrsUsers, null);
}
function _rsUsers_dtor()
{
	rsUsers._preserveState();
	thisPage.setState('pb_rsUsers', rsUsers.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<table border="1" width="350" cellpadding=3 cellspacing=0>
  <tr>
	<td width=180 bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="../index.asp" target="_top" vlink="White">Home</a></strong></big></font></center></td>
    <td width=180 bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="AdminSecure.asp" target="_top" vlink="White">Administration</a></strong></big></font></center></td>
    <td><strong><font face="Verdana"><center>User&nbsp;Information</center></font></strong>
  </tr>
</table>
<br>

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnGoto style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 57px" 
	width=57>
	<PARAM NAME="_ExtentX" VALUE="1508">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnGoto">
	<PARAM NAME="Caption" VALUE="Go to ">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnGoto()
{
	btnGoto.value = 'Go to ';
	btnGoto.setStyle(0);
}
function _btnGoto_ctor()
{
	CreateButton('btnGoto', _initbtnGoto, null);
}
</script>
<% btnGoto.display %>

<!--METADATA TYPE="DesignerControl" endspan-->

<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E450-DC5F-11D0-9846-0000F8027CA0" height=21 id=lstUsers 
	style="HEIGHT: 21px; LEFT: 10px; TOP: 170px; WIDTH: 96px" width=96>
	<PARAM NAME="_ExtentX" VALUE="2540">
	<PARAM NAME="_ExtentY" VALUE="556">
	<PARAM NAME="id" VALUE="lstUsers">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="ControlStyle" VALUE="0">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="UsesStaticList" VALUE="0">
	<PARAM NAME="RowSource" VALUE="rsUsers">
	<PARAM NAME="BoundColumn" VALUE="UserID">
	<PARAM NAME="ListField" VALUE="UserName">
	<PARAM NAME="LookupPlatform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/ListBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlstUsers()
{
	rsUsers.advise(RS_ONDATASETCOMPLETE, 'lstUsers.setRowSource(rsUsers, \'UserName\', \'UserID\');');
}
function _lstUsers_ctor()
{
	CreateListbox('lstUsers', _initlstUsers, null);
}
</script>
<% lstUsers.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
&nbsp;&nbsp;&nbsp;
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
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=Button1 style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 46px" 
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
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" id=btnUpdate style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="1773">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnUpdate">
	<PARAM NAME="Caption" VALUE="Update">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
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
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btnDelete 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 62px" width=62>
	<PARAM NAME="_ExtentX" VALUE="1640">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btnDelete">
	<PARAM NAME="Caption" VALUE="Delete">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtnDelete()
{
	btnDelete.value = 'Delete';
	btnDelete.setStyle(0);
}
function _btnDelete_ctor()
{
	CreateButton('btnDelete', _initbtnDelete, null);
}
</script>
<% btnDelete.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<input type=button onclick='window.location.href="userrole.asp?userid=<%=Recordset1.fields.getValue("UserID") %>"' value="Edit User Roles">
<br>
<br>
<TABLE CellPadding=1 CellSpacing=1 Cols=2>
<TR>
	<TD align=right><FONT size=-1 face=Verdana>User Name </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" 
            id=Textbox2 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox2">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="UserName">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="20">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox2()
{
	Textbox2.setStyle(TXT_TEXTBOX);
	Textbox2.setDataSource(Recordset1);
	Textbox2.setDataField('UserName');
	Textbox2.setMaxLength(20);
	Textbox2.setColumnCount(20);
}
function _Textbox2_ctor()
{
	CreateTextbox('Textbox2', _initTextbox2, null);
}
</script>
<% Textbox2.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD></tr>
<TR>
	<TD align=right><FONT size=-1 face=Verdana>First Name </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" 
            id=Textbox3 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox3">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="FirstName">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="30">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox3()
{
	Textbox3.setStyle(TXT_TEXTBOX);
	Textbox3.setDataSource(Recordset1);
	Textbox3.setDataField('FirstName');
	Textbox3.setMaxLength(30);
	Textbox3.setColumnCount(20);
}
function _Textbox3_ctor()
{
	CreateTextbox('Textbox3', _initTextbox3, null);
}
</script>
<% Textbox3.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD>
	<TD align=right><FONT size=-1 face=Verdana>Last Name </FONT>
	<TD colspan=3><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" 
            id=Textbox4 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox4">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="LastName">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="30">
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
	Textbox4.setDataField('LastName');
	Textbox4.setMaxLength(30);
	Textbox4.setColumnCount(20);
}
function _Textbox4_ctor()
{
	CreateTextbox('Textbox4', _initTextbox4, null);
}
</script>
<% Textbox4.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD></tr><tr>
	<TD align=right><FONT size=-1 face=Verdana>Address </FONT>
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
	<PARAM NAME="DataField" VALUE="Address">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="40">
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
	Textbox5.setDataField('Address');
	Textbox5.setMaxLength(40);
	Textbox5.setColumnCount(20);
}
function _Textbox5_ctor()
{
	CreateTextbox('Textbox5', _initTextbox5, null);
}
</script>
<% Textbox5.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD></tr>
<TR>
	<TD align=right><FONT size=-1 face=Verdana>City </FONT>
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
	<PARAM NAME="DataField" VALUE="City">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="40">
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
	Textbox6.setDataField('City');
	Textbox6.setMaxLength(40);
	Textbox6.setColumnCount(20);
}
function _Textbox6_ctor()
{
	CreateTextbox('Textbox6', _initTextbox6, null);
}
</script>
<% Textbox6.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD>
	<TD align=right><FONT size=-1 face=Verdana>State </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=Textbox7 
	style="HEIGHT: 19px; LEFT: 10px; TOP: 392px; WIDTH: 24px" width=24>
	<PARAM NAME="_ExtentX" VALUE="635">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox7">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="State">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="2">
	<PARAM NAME="DisplayWidth" VALUE="4">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox7()
{
	Textbox7.setStyle(TXT_TEXTBOX);
	Textbox7.setDataSource(Recordset1);
	Textbox7.setDataField('State');
	Textbox7.setMaxLength(2);
	Textbox7.setColumnCount(4);
}
function _Textbox7_ctor()
{
	CreateTextbox('Textbox7', _initTextbox7, null);
}
</script>
<% Textbox7.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD>
	<TD align=right><FONT size=-1 face=Verdana>Zip </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=Textbox8 
	style="HEIGHT: 19px; LEFT: 10px; TOP: 411px; WIDTH: 60px" width=60>
	<PARAM NAME="_ExtentX" VALUE="1588">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox8">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="Zip">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="50">
	<PARAM NAME="DisplayWidth" VALUE="10">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox8()
{
	Textbox8.setStyle(TXT_TEXTBOX);
	Textbox8.setDataSource(Recordset1);
	Textbox8.setDataField('Zip');
	Textbox8.setMaxLength(50);
	Textbox8.setColumnCount(10);
}
function _Textbox8_ctor()
{
	CreateTextbox('Textbox8', _initTextbox8, null);
}
</script>
<% Textbox8.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD></tr>
<TR>
	<TD align=right><FONT size=-1 face=Verdana>Phone </FONT>
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
	<PARAM NAME="DataField" VALUE="Phone">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="15">
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
	Textbox9.setDataField('Phone');
	Textbox9.setMaxLength(15);
	Textbox9.setColumnCount(20);
}
function _Textbox9_ctor()
{
	CreateTextbox('Textbox9', _initTextbox9, null);
}
</script>
<% Textbox9.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD></tr>
<TR>
	<TD align=right><FONT size=-1 face=Verdana>Email </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" 
            id=Textbox10 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox10">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="Email">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="35">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox10()
{
	Textbox10.setStyle(TXT_TEXTBOX);
	Textbox10.setDataSource(Recordset1);
	Textbox10.setDataField('Email');
	Textbox10.setMaxLength(35);
	Textbox10.setColumnCount(20);
}
function _Textbox10_ctor()
{
	CreateTextbox('Textbox10', _initTextbox10, null);
}
</script>
<% Textbox10.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD></tr>
<TR>
	<TD align=right><FONT size=-1 face=Verdana>Expire Date </FONT>
	<TD colspan=3><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" 
            id=Textbox11 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox11">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="ExpireDate">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="19">
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
	Textbox11.setDataField('ExpireDate');
	Textbox11.setMaxLength(19);
	Textbox11.setColumnCount(20);
}
function _Textbox11_ctor()
{
	CreateTextbox('Textbox11', _initTextbox11, null);
}
</script>
<% Textbox11.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD><td></td><td></td></tr>
<TR>
	<TD><FONT size=-1 face=Verdana>Last Logon Date </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" 
            id=Textbox12 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox12">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="LastLogonDate">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="19">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox12()
{
	Textbox12.setStyle(TXT_TEXTBOX);
	Textbox12.setDataSource(Recordset1);
	Textbox12.setDataField('LastLogonDate');
	Textbox12.setMaxLength(19);
	Textbox12.setColumnCount(20);
}
function _Textbox12_ctor()
{
	CreateTextbox('Textbox12', _initTextbox12, null);
}
</script>
<% Textbox12.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD>
	<TD><FONT size=-1 face=Verdana>Last Logon Attempt </FONT>
	<TD colspan=3><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" 
            id=Textbox13 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox13">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="LastLogonAttempt">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="19">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox13()
{
	Textbox13.setStyle(TXT_TEXTBOX);
	Textbox13.setDataSource(Recordset1);
	Textbox13.setDataField('LastLogonAttempt');
	Textbox13.setMaxLength(19);
	Textbox13.setColumnCount(20);
}
function _Textbox13_ctor()
{
	CreateTextbox('Textbox13', _initTextbox13, null);
}
</script>
<% Textbox13.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD></tr>
<TR>
	<TD><FONT size=-1 face=Verdana>Failed Logons </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" 
            id=Textbox14 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="Textbox14">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="FailedLogons">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="10">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../">
	
             </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initTextbox14()
{
	Textbox14.setStyle(TXT_TEXTBOX);
	Textbox14.setDataSource(Recordset1);
	Textbox14.setDataField('FailedLogons');
	Textbox14.setMaxLength(10);
	Textbox14.setColumnCount(20);
}
function _Textbox14_ctor()
{
	CreateTextbox('Textbox14', _initTextbox14, null);
}
</script>
<% Textbox14.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD></tr>
<% 
	If Recordset1.fields.getValue("Password") <> "" then
		Response.Write ("<tr><td colspan=2><font size=-1 face=verdana color=green>This user has a password</font></td></tr>")
	Else
		Response.Write ("<tr><td colspan=2><font face=verdana color=red>This user has a BLANK password</font></td></tr>")
	End If
%>
<TR>
	<TD><FONT size=-1 face=Verdana> New Password </FONT>
	<TD><FONT face=Verdana>
<INPUT type="password" id=password1 name=password>
</FONT></TD>
	<TD><FONT size=-1 face=Verdana> Verify Password </FONT>
	<TD colspan=3><FONT face=Verdana>
<INPUT type="password" id=password1 name=verifypassword>            
</FONT></TD></tr>

<TR>
	<TD><FONT size=-1 face=Verdana>AccountLocked </FONT>
	<TD><FONT face=Verdana>
            <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E46C-DC5F-11D0-9846-0000F8027CA0" height=27 id=Checkbox2 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 29px" width=29>
	<PARAM NAME="_ExtentX" VALUE="767">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="Checkbox2">
	<PARAM NAME="Caption" VALUE="">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="AccountLocked">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/CheckBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initCheckbox2()
{
	Checkbox2.setDataSource(Recordset1);
	Checkbox2.setDataField('AccountLocked');
}
function _Checkbox2_ctor()
{
	CreateCheckbox('Checkbox2', _initCheckbox2, null);
}
</script>
<% Checkbox2.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD>
</TR>
</TABLE><br>
</FONT></p>
<!-- #include File="../copyright.asp" -->
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</html>
