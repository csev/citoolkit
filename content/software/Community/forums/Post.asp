<%@ Language=VBScript %>
<!-- #include File="../mapvar.asp" -->
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<body>
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

Sub btn_Post_onclick()
	
	quote = chr(39)
	
	sql = "INSERT INTO messages (MessageTitle, MessageBody, UserID, MessageDate, ConferenceID, " & _
	"ParentMessageID) VALUES (" & quote & txt_MessageTitle.value & quote & "," & quote & txt_MessageBody.value & quote & "," & _
	 quote & txt_UserID.value & quote & "," & quote & txt_MessageDate.value & quote & "," & _
	 quote & txt_ConferenceID.value & quote & "," & quote & txt_ParentMessageID.value & quote & ")"
	
	'Response.Write sql
	
	Set conn = Server.CreateObject ("ADODB.Command")
	
	conn.CommandText = sql
	conn.ActiveConnection =  Session("Connection1_ConnectionString")
	'conn.CommandType = adCmdText
	conn.Execute Rec_Affect
	
	if Rec_Affect = 1 then
		Response.Write "Your message has been posted!"
	end if

	Response.Redirect "Post_Success.htm"

End Sub

Sub btn_Cancel_Post_onclick()
	Response.Redirect "Default_Display_Message.asp"
End Sub

</SCRIPT>


<%		

	txt_UserID.value = Request.QueryString("UserID")
	txt_ConferenceID.value = Request.QueryString("ConferenceID")
	txt_ParentMessageID.value = Request.QueryString("ParentMessageID")
	txt_MessageDate.value = FormatDateTime(Now, 2)
	txt_MessageTitle.value = Request.QueryString("MessageTitle")

	
if Request.QueryString("PostType") = "N" then
	Response.Write "<strong>New Message &nbsp;</strong>" 
	txt_MessageTitle.disabled = 0
else
	Response.Write "<strong>Message Reply &nbsp;</strong>"
	txt_MessageTitle.disabled = 1
end if

%>	
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btn_Post 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 47px" width=47>
	<PARAM NAME="_ExtentX" VALUE="1143">
	<PARAM NAME="_ExtentY" VALUE="699">
	<PARAM NAME="id" VALUE="btn_Post">
	<PARAM NAME="Caption" VALUE="Post">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtn_Post()
{
	btn_Post.value = 'Post';
	btn_Post.setStyle(0);
}
function _btn_Post_ctor()
{
	CreateButton('btn_Post', _initbtn_Post, null);
}
</script>
<% btn_Post.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=27 id=btn_Cancel_Post 
	style="HEIGHT: 27px; LEFT: 0px; TOP: 0px; WIDTH: 64px" width=64>
	<PARAM NAME="_ExtentX" VALUE="1566">
	<PARAM NAME="_ExtentY" VALUE="699">
	<PARAM NAME="id" VALUE="btn_Cancel_Post">
	<PARAM NAME="Caption" VALUE="Cancel">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initbtn_Cancel_Post()
{
	btn_Cancel_Post.value = 'Cancel';
	btn_Cancel_Post.setStyle(0);
}
function _btn_Cancel_Post_ctor()
{
	CreateButton('btn_Cancel_Post', _initbtn_Cancel_Post, null);
}
</script>
<% btn_Cancel_Post.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" id=txt_MessageDate style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="2963">
	<PARAM NAME="_ExtentY" VALUE="508">
	<PARAM NAME="id" VALUE="txt_MessageDate">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="0">
	<PARAM NAME="MaxChars" VALUE="20">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/TextBox.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxt_MessageDate()
{
	txt_MessageDate.setStyle(TXT_TEXTBOX);
	txt_MessageDate.hide();
	txt_MessageDate.setMaxLength(20);
	txt_MessageDate.setColumnCount(20);
}
function _txt_MessageDate_ctor()
{
	CreateTextbox('txt_MessageDate', _inittxt_MessageDate, null);
}
</script>
<% txt_MessageDate.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" id=txt_ConferenceID style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="2963">
	<PARAM NAME="_ExtentY" VALUE="508">
	<PARAM NAME="id" VALUE="txt_ConferenceID">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="0">
	<PARAM NAME="MaxChars" VALUE="20">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxt_ConferenceID()
{
	txt_ConferenceID.setStyle(TXT_TEXTBOX);
	txt_ConferenceID.hide();
	txt_ConferenceID.setMaxLength(20);
	txt_ConferenceID.setColumnCount(20);
}
function _txt_ConferenceID_ctor()
{
	CreateTextbox('txt_ConferenceID', _inittxt_ConferenceID, null);
}
</script>
<% txt_ConferenceID.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" id=txt_ParentMessageID 
	style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="_ExtentX" VALUE="2963">
	<PARAM NAME="_ExtentY" VALUE="508">
	<PARAM NAME="id" VALUE="txt_ParentMessageID">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="0">
	<PARAM NAME="MaxChars" VALUE="20">
	<PARAM NAME="DisplayWidth" VALUE="20">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxt_ParentMessageID()
{
	txt_ParentMessageID.setStyle(TXT_TEXTBOX);
	txt_ParentMessageID.hide();
	txt_ParentMessageID.setMaxLength(20);
	txt_ParentMessageID.setColumnCount(20);
}
function _txt_ParentMessageID_ctor()
{
	CreateTextbox('txt_ParentMessageID', _inittxt_ParentMessageID, null);
}
</script>
<% txt_ParentMessageID.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=19 id=txt_UserID 
	style="HEIGHT: 19px; LEFT: 0px; TOP: 0px; WIDTH: 210px" width=210>
	<PARAM NAME="_ExtentX" VALUE="4445">
	<PARAM NAME="_ExtentY" VALUE="508">
	<PARAM NAME="id" VALUE="txt_UserID">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="3">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="0">
	<PARAM NAME="MaxChars" VALUE="20">
	<PARAM NAME="DisplayWidth" VALUE="30">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxt_UserID()
{
	txt_UserID.setStyle(TXT_TEXTBOX);
	txt_UserID.hide();
	txt_UserID.setMaxLength(20);
	txt_UserID.setColumnCount(30);
}
function _txt_UserID_ctor()
{
	CreateTextbox('txt_UserID', _inittxt_UserID, null);
}
</script>
<% txt_UserID.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<br><br>Subject:<br>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=24 id=txt_MessageTitle 
	style="HEIGHT: 24px; LEFT: 10px; TOP: 199px; WIDTH: 350px" width=350>
	<PARAM NAME="_ExtentX" VALUE="7408">
	<PARAM NAME="_ExtentY" VALUE="508">
	<PARAM NAME="id" VALUE="txt_MessageTitle">
	<PARAM NAME="ControlType" VALUE="0">
	<PARAM NAME="Lines" VALUE="1">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="50">
	<PARAM NAME="DisplayWidth" VALUE="50">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxt_MessageTitle()
{
	txt_MessageTitle.setStyle(TXT_TEXTBOX);
	txt_MessageTitle.setMaxLength(50);
	txt_MessageTitle.setColumnCount(50);
}
function _txt_MessageTitle_ctor()
{
	CreateTextbox('txt_MessageTitle', _inittxt_MessageTitle, null);
}
</script>
<% txt_MessageTitle.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<br><br>Message:<br>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E469-DC5F-11D0-9846-0000F8027CA0" height=160 id=txt_MessageBody 
	style="HEIGHT: 160px; LEFT: 10px; TOP: 223px; WIDTH: 280px" width=280>
	<PARAM NAME="_ExtentX" VALUE="5927">
	<PARAM NAME="_ExtentY" VALUE="3387">
	<PARAM NAME="id" VALUE="txt_MessageBody">
	<PARAM NAME="ControlType" VALUE="1">
	<PARAM NAME="Lines" VALUE="10">
	<PARAM NAME="DataSource" VALUE="">
	<PARAM NAME="DataField" VALUE="">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="MaxChars" VALUE="500">
	<PARAM NAME="DisplayWidth" VALUE="40">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _inittxt_MessageBody()
{
	txt_MessageBody.setStyle(TXT_TEXTAREA);
	txt_MessageBody.setRowCount(10);
	txt_MessageBody.setColumnCount(40);
}
function _txt_MessageBody_ctor()
{
	CreateTextbox('txt_MessageBody', _inittxt_MessageBody, null);
}
</script>
<% txt_MessageBody.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>

</HTML>
