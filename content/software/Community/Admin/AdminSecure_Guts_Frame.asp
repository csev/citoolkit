<%@ Language=VBScript%>
<!-- #include File="../mapvar.asp" -->
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="../_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
<!--#include file="../ADOVBS.INC"-->
<html>

<head>
<title>Admin-Main Menu</title>
</head>

<body vlink=White alink=White link=White>
<SCRIPT ID=serverEventHandlersVBS LANGUAGE=vbscript RUNAT=Server>

Sub Help_OnClick()

Response.Redirect "../docs/default.htm"

end sub
</script>


<table border="1" width="350" cellpadding=3 cellspacing=0>
  <tr>
	<td width=180 bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="../index.asp" target="_top" vlink="White">Home</a></strong></big></font></center></td>
    <td><strong><font face="Verdana"><center>Administration</center></font></strong>
  </tr>
</table>
<P>
<%
	UserID = Session("UserID")
	If UserID <> "" then
		Sql = "Select *  from AdminFunctions where AdminFunctID IN (Select AdminFunctID from AdminFunctRoles where RoleID IN (Select RoleID from UserRoles where UserID=" & cstr(UserID) & "))"
		'Response.Write sql

		Set conn = Server.CreateObject("ADODB.Connection")
		conn.Open Application("Connection1_ConnectionString")
		Set RecordSet1 = Server.CreateObject("ADODB.Recordset")
		RecordSet1.Open sql, conn, , , adCmdText
		    
		If Recordset1.EOF then
			Response.Write "You have not been granted any admin priveleges."
		End If

		Do While NOT (RecordSet1.EOF)    
		    	Response.Write "<table width=190 bgcolor=#008080 border=""0"" width=""100%""><tr>" 
		    	Response.Write "<td><font face=""Verdana"" color=""#FFFFFF""><a HREF=""http:" & Recordset1.fields("Url") & """ ><strong> "  & Recordset1.fields("AdminFunctName") & "</strong></a> <br></font></td>"
		    	Response.Write "</table>"
			RecordSet1.MoveNext
		Loop

		RecordSet1.Close
		set Recordset1 = nothing
		Conn.Close
		Set conn = nothing
%>
<br><br>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B6FC3A14-F837-11D0-9CC8-006008058731" height=33 id=Help style="HEIGHT: 33px; LEFT: 10px; TOP: 37px; WIDTH: 55px" 
	width=55>
	<PARAM NAME="_ExtentX" VALUE="1164">
	<PARAM NAME="_ExtentY" VALUE="699">
	<PARAM NAME="id" VALUE="Help">
	<PARAM NAME="Caption" VALUE="Help">
	<PARAM NAME="Image" VALUE="">
	<PARAM NAME="AltText" VALUE="">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="Platform" VALUE="256">
	<PARAM NAME="LocalPath" VALUE="../"></OBJECT>
-->
<!--#INCLUDE FILE="../_ScriptLibrary/Button.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initHelp()
{
	Help.value = 'Help';
	Help.setStyle(0);
}
function _Help_ctor()
{
	CreateButton('Help', _initHelp, null);
}
</script>
<% Help.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
<%		
	Else
		Response.Write "You must login to use this function"
	End If

%>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:8CC35CD6-E98B-11D0-B218-00A0C92764F5" id=PageObject1 style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="4233">
	<PARAM NAME="ExtentY" VALUE="1508">
	<PARAM NAME="State" VALUE="(ObjectName_Unmatched=\qAdminSecure_Guts_Frame\q,NavigateMethods=(Rows=0),ExecuteMethods=(Rows=0),Properties=(Rows=0),References=(Rows=0))"></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=SERVER>
/* VIPM PAGE DESCRIPTION
<DSC NAME="AdminSecure_Guts_Frame">
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
if (typeof AdminSecure_Guts_Frame_onbeforeserverevent == 'function' || typeof AdminSecure_Guts_Frame_onbeforeserverevent == 'unknown')
	thisPage.advise('onbeforeserverevent', 'AdminSecure_Guts_Frame_onbeforeserverevent()');

AdminSecure_Guts_Frame = thisPage;
AdminSecure_Guts_Frame.location = "../Admin/AdminSecure_Guts_Frame.asp";
AdminSecure_Guts_Frame.navigate = new Object;
AdminSecure_Guts_Frame.navigate.show = Function('thisPage.invokeMethod("", "show", this.show.arguments);');
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

AdminSecure_Guts_Frame = thisPage;
AdminSecure_Guts_Frame.location = "../Admin/AdminSecure_Guts_Frame.asp";
AdminSecure_Guts_Frame.navigate = new Object;
AdminSecure_Guts_Frame.navigate.show = Function('return;');

	thisPage._objEventManager.adviseDefaultHandler('AdminSecure_Guts_Frame','onenter');
	thisPage._objEventManager.adviseDefaultHandler('AdminSecure_Guts_Frame','onexit');
	thisPage._objEventManager.adviseDefaultHandler('AdminSecure_Guts_Frame','onshow');
	thisPage.registerVTable(thisPage.navigate, PAGE_NAVIGATE);
}

function _PO_dtor()
{
if (thisPage._redirect == '')
	_PO_OutputClientCode();
}

</SCRIPT>


<!--METADATA TYPE="DesignerControl" endspan-->
<!-- #include File="../copyright.asp" -->
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</html>
