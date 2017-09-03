<%@ Language=VBScript %>
<!-- #include File="../mapvar.asp" -->
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
<!--#include file="../ADOVBS.INC"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

Sub btn_Login_OnClick()
		
	quote = chr(39)

	RS_Logon.setSQLText("Select *  from Users where Username=" & quote & txtbox_Username.value & quote)
	RS_Logon.open
	
	RS_Site.setSQLText("Select * from SiteInfo where SiteID=1")
	RS_Site.open
	
	SiteAccountLockoutDefault = RS_Site.fields.getValue("SiteAccountLockout")

	RS_Site.Close

	if (RS_Logon.EOF) then
		' bad userID
		' Reload page with mode=2
		Response.Redirect "Login_Guts_Frame.asp?mode=2"
	elseif ((RS_Logon.fields.getValue("Password") = txtbox_password.value) AND (RS_Logon.fields.getValue("Deleted") = FALSE) AND (RS_Logon.fields.getValue("AccountLocked") = FALSE)) then
		' good password, account NOT deleted and not Locked
		' update the users table with current login info
		RS_Logon.fields.setValue"LastLogonDate", cstr(Now())
		RS_Logon.fields.setValue"LastLogonAttempt", cstr(Now())
		RS_Logon.fields.setValue"FailedLogons", "0"
		' load the ADMIN default.htm page
		RS_Logon.updateRecord
		Session("UserID") = rs_Logon.fields.getValue("UserID")
		'Response.Redirect "AdminSecure.asp?UserID=" & rs_Logon.fields.getValue("UserID")
		'Set cal = server.CreateObject("Community.Calendar")
		'Cal.LogEvent 4,"Admin logged in successfully:"& cstr(txtbox_Username.value)
		'Set Cal = Nothing
	
		Response.Redirect "AdminSecure_Guts_Frame.asp"		 		
		
	else
		' bad Password or Deleted account or Locked account 
		' increment Failed logons and update LastLogon Attempt
		RS_Logon.fields.setValue"LastLogonAttempt", cstr(Now())

		temp_FailedLogons = cint("0" & RS_Logon.fields.getValue("FailedLogons")) + 1
		RS_Logon.fields.setValue"FailedLogons", (temp_FailedLogons)

		'Response.Write "SiteAccLockDef : " & SiteAccountLockoutDefault & "  FailedLogongs : " & RS_Logon.fields.getValue("FailedLogons")

		if (RS_Logon.fields.getValue("FailedLogons") >= SiteAccountLockoutDefault) then
			RS_Logon.fields.setValue"AccountLocked", True
			RS_Logon.updateRecord
			' Reload page with mode=3 (account locked)
			Response.Redirect "Login_Guts_Frame.asp?mode=3"
		else
			RS_Logon.updateRecord
			' Reload page with mode=2
			Response.Redirect "Login_Guts_Frame.asp?mode=2"	
		end if
		
	end if
	
	RS_Logon.Close
		
end Sub

</SCRIPT>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=RS_Logon style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject=\qTables\q,TCDBObjectName=\qUsers\q,TCControlID_Unmatched=\qRS_Logon\q,TCPPConn=\qConnection1\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qUsers\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Recordset.ASP"-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initRS_Logon()
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
	RS_Logon.setRecordSource(rsTmp);
	if (thisPage.getState('pb_RS_Logon') != null)
		RS_Logon.setBookmark(thisPage.getState('pb_RS_Logon'));
}
function _RS_Logon_ctor()
{
	CreateRecordset('RS_Logon', _initRS_Logon, null);
}
function _RS_Logon_dtor()
{
	RS_Logon._preserveState();
	thisPage.setState('pb_RS_Logon', RS_Logon.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=RS_Site style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12197">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject=\qTables\q,TCDBObjectName=\qSiteInfo\q,TCControlID_Unmatched=\qRS_Site\q,TCPPConn=\qConnection1\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qSiteInfo\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=0,GCParameters=(Rows=0))">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE="JavaScript" RUNAT="server">
function _initRS_Site()
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
	RS_Site.setRecordSource(rsTmp);
	if (thisPage.getState('pb_RS_Site') != null)
		RS_Site.setBookmark(thisPage.getState('pb_RS_Site'));
}
function _RS_Site_ctor()
{
	CreateRecordset('RS_Site', _initRS_Site, null);
}
function _RS_Site_dtor()
{
	RS_Site._preserveState();
	thisPage.setState('pb_RS_Site', RS_Site.getBookmark());
}
</SCRIPT>

<!--METADATA TYPE="DesignerControl" endspan-->
<%  if (Request.QueryString("mode") = 1) then
	 Page_Title = "<p>Community Site Administration Login</p>"
	elseif (Request.QueryString("mode") = 2) then
	 Page_Title = "<p>Login Attempt Failed.</p>"
	elseif (Request.QueryString("mode") = 3) then
	 Page_Title = "<p>Login Attempt Failed.  Your account is LOCKED.<br>Contact a System Administrator for Help.</p>"
	end if
 		
Response.Write "<p><strong><big><big><font face=verdana>" & Page_Title & "</font></big></big></strong></p>"
%>
<p>
<font face=verdana>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=lbl_UserName 
	style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 56px" width=56>
	<PARAM NAME="_ExtentX" VALUE="1482">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lbl_UserName">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Username:">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Label.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlbl_UserName()
{
	lbl_UserName.setCaption('Username:');
}
function _lbl_UserName_ctor()
{
	CreateLabel('lbl_UserName', _initlbl_UserName, null);
}
</script>
<% lbl_UserName.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
&nbsp;
 <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" id=txtbox_Username style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtbox_Username">
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
function _inittxtbox_Username()
{
	txtbox_Username.setStyle(TXT_TEXTBOX);
	txtbox_Username.setMaxLength(20);
	txtbox_Username.setColumnCount(20);
}
function _txtbox_Username_ctor()
{
	CreateTextbox('txtbox_Username', _inittxtbox_Username, null);
}
</script>
<% txtbox_Username.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</p><p>
<font face=verdana>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=lbl_Password 
	style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 54px" width=54>
	<PARAM NAME="_ExtentX" VALUE="1429">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="lbl_Password">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="Password:">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initlbl_Password()
{
	lbl_Password.setCaption('Password:');
}
function _lbl_Password_ctor()
{
	CreateLabel('lbl_Password', _initlbl_Password, null);
}
</script>
<% lbl_Password.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</font>
&nbsp;
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" id=txtbox_password style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="3175">
	<PARAM NAME="_ExtentY" VALUE="503">
	<PARAM NAME="id" VALUE="txtbox_password">
	<PARAM NAME="ControlType" VALUE="2">
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
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxtbox_password()
{
	txtbox_password.setStyle(TXT_PASSWORD);
	txtbox_password.setMaxLength(20);
	txtbox_password.setColumnCount(20);
}
function _txtbox_password_ctor()
{
	CreateTextbox('txtbox_password', _inittxtbox_password, null);
}
</script>
<% txtbox_password.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</p><p>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btn_Login 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 54px" width=54>
	<PARAM NAME="_ExtentX" VALUE="1429">
	<PARAM NAME="_ExtentY" VALUE="714">
	<PARAM NAME="id" VALUE="btn_Login">
	<PARAM NAME="Caption" VALUE="Login">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtn_Login()
{
	btn_Login.value = 'Login';
	btn_Login.setStyle(0);
}
function _btn_Login_ctor()
{
	CreateButton('btn_Login', _initbtn_Login, null);
}
</script>
<% btn_Login.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</p>

<!-- #include File="../copyright.asp" -->
</BODY>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</HTML>
