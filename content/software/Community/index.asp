<%@ Language=VBScript %>
<% ' *** map variables to user session *** %>
<!-- #include File="mapvar.asp" -->
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--#include file="_ScriptLibrary/pm.asp"-->
<% if StartPageProcessing() Then Response.End() %>
<FORM name=thisForm METHOD=post>
<html>
	<head>
		<title>Library Home Page</title>
	</head>
<body vLink=White aLink=White Link=White>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:9CF5D7C2-EC10-11D0-9862-0000F8027CA0" id=Recordset1 
style="LEFT: 0px; TOP: 0px">
	<PARAM NAME="ExtentX" VALUE="12198">
	<PARAM NAME="ExtentY" VALUE="2090">
	<PARAM NAME="State" VALUE="(TCConn=\qConnection1\q,TCDBObject=\qTables\q,TCDBObjectName=\qSiteInfo\q,TCControlID_Unmatched=\qRecordset1\q,TCPPConn=\qConnection1\q,RCDBObject=\qRCDBObject\q,TCPPDBObject=\qTables\q,TCPPDBObjectName=\qSiteInfo\q,TCCursorType=\q3\s-\sStatic\q,TCCursorLocation=\q3\s-\sUse\sclient-side\scursors\q,TCLockType=\q3\s-\sOptimistic\q,TCCacheSize_Unmatched=\q10\q,TCCommTimeout_Unmatched=\q10\q,CCPrepared=0,CCAllRecords=1,TCNRecords_Unmatched=\q10\q,TCODBCSyntax_Unmatched=\q\q,TCHTargetPlatform=\q\q,TCHTargetBrowser_Unmatched=\qServer\s(ASP)\q,TCTargetPlatform=\qInherit\sfrom\spage\q,RCCache=\qRCBookPage\q,CCOpen=1,GCParameters=(Rows=0))">
	
 </OBJECT>
-->
<!--#INCLUDE FILE="_ScriptLibrary/Recordset.ASP"-->
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
	cmdTmp.CommandText = '`SiteInfo`';
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
<table>
<tr>
<td width=180 bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="http:calendar/calendar.asp" target="_top" vlink="White">Calendar</a></strong></big></font></center></td>
<td rowspan="4"><FONT></FONT><TABLE CellPadding=2 CellSpacing=2 Cols=2>
<TR>
	<TD><FONT face=Verdana>Organization Name </FONT></td>
	<TD><FONT face=Verdana>
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT 
                        classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" 
                        height=17 id=Label1 
                        style="HEIGHT: 17px; LEFT: 10px; TOP: 113px; WIDTH: 49px" 
                        width=49>
	<PARAM NAME="_ExtentX" VALUE="1296">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="Label1">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="OrgName">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="">
	
                         </OBJECT>
-->
<!--#INCLUDE FILE="_ScriptLibrary/Label.ASP"-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel1()
{
	Label1.setDataSource(Recordset1);
	Label1.setDataField('OrgName');
}
function _Label1_ctor()
{
	CreateLabel('Label1', _initLabel1, null);
}
</script>
<% Label1.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD>
</TR>
<TR>
	<TD><FONT face=Verdana>Address </FONT></td>
	<TD><FONT face=Verdana>
                        <!--METADATA TYPE="DesignerControl" startspan
<OBJECT 
                        classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" 
                        height=17 id=Label2 
                        style="HEIGHT: 17px; LEFT: 10px; TOP: 130px; WIDTH: 51px" 
                        width=51>
	<PARAM NAME="_ExtentX" VALUE="1349">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="Label2">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="OrgAddr1">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE="">
	
                         
                        </OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel2()
{
	Label2.setDataSource(Recordset1);
	Label2.setDataField('OrgAddr1');
}
function _Label2_ctor()
{
	CreateLabel('Label2', _initLabel2, null);
}
</script>
<% Label2.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD>
</TR>
<TR>
	<TD><FONT face=Verdana> </FONT></td>
	<TD><FONT face=Verdana>
  <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=Label3 style="HEIGHT: 17px; LEFT: 10px; TOP: 147px; WIDTH: 51px" 
	width=51>
	<PARAM NAME="_ExtentX" VALUE="1349">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="Label3">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="OrgAddr2">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE=""></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel3()
{
	Label3.setDataSource(Recordset1);
	Label3.setDataField('OrgAddr2');
}
function _Label3_ctor()
{
	CreateLabel('Label3', _initLabel3, null);
}
</script>
<% Label3.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD>
<TR>
	<TD><FONT face=Verdana> </FONT></td>
	<TD><FONT face=Verdana>
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=Label4 style="HEIGHT: 17px; LEFT: 10px; TOP: 164px; WIDTH: 41px" 
	width=41>
	<PARAM NAME="_ExtentX" VALUE="1085">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="Label4">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="OrgCity">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE=""></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel4()
{
	Label4.setDataSource(Recordset1);
	Label4.setDataField('OrgCity');
}
function _Label4_ctor()
{
	CreateLabel('Label4', _initLabel4, null);
}
</script>
<% Label4.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
&nbsp;
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=Label5 style="HEIGHT: 17px; LEFT: 10px; TOP: 181px; WIDTH: 48px" 
	width=48>
	<PARAM NAME="_ExtentX" VALUE="1270">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="Label5">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="OrgState">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE=""></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel5()
{
	Label5.setDataSource(Recordset1);
	Label5.setDataField('OrgState');
}
function _Label5_ctor()
{
	CreateLabel('Label5', _initLabel5, null);
}
</script>
<% Label5.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
&nbsp;
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=Label6 style="HEIGHT: 17px; LEFT: 10px; TOP: 198px; WIDTH: 36px" 
	width=36>
	<PARAM NAME="_ExtentX" VALUE="953">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="Label6">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="OrgZip">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE=""></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel6()
{
	Label6.setDataSource(Recordset1);
	Label6.setDataField('OrgZip');
}
function _Label6_ctor()
{
	CreateLabel('Label6', _initLabel6, null);
}
</script>
<% Label6.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</FONT></TD>
</tr>
<TR>
	<TD><FONT face=Verdana>Phone </FONT></td>
	<TD><FONT face=Verdana>
 <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=Label7 style="HEIGHT: 17px; LEFT: 10px; TOP: 215px; WIDTH: 52px" 
	width=52>
	<PARAM NAME="_ExtentX" VALUE="1376">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="Label7">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="OrgPhone">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE=""></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel7()
{
	Label7.setDataSource(Recordset1);
	Label7.setDataField('OrgPhone');
}
function _Label7_ctor()
{
	CreateLabel('Label7', _initLabel7, null);
}
</script>
<% Label7.display %>

<!--METADATA TYPE="DesignerControl" endspan-->                       
</FONT></TD>
</tr>
<TR>
	<TD><FONT face=Verdana>Fax </FONT></td>
	<TD><FONT face=Verdana>
   <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=Label8 style="HEIGHT: 17px; LEFT: 10px; TOP: 232px; WIDTH: 40px" 
	width=40>
	<PARAM NAME="_ExtentX" VALUE="1058">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="Label8">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="OrgFax">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE=""></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel8()
{
	Label8.setDataSource(Recordset1);
	Label8.setDataField('OrgFax');
}
function _Label8_ctor()
{
	CreateLabel('Label8', _initLabel8, null);
}
</script>
<% Label8.display %>

<!--METADATA TYPE="DesignerControl" endspan-->                     
</FONT></TD>
</tr>
<TR>
	<TD><FONT face=Verdana>Contact Name </FONT></td>
	<TD><FONT face=Verdana>
   <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=Label9 style="HEIGHT: 17px; LEFT: 10px; TOP: 249px; WIDTH: 87px" 
	width=87>
	<PARAM NAME="_ExtentX" VALUE="2302">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="Label9">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="OrgContactName">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE=""></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel9()
{
	Label9.setDataSource(Recordset1);
	Label9.setDataField('OrgContactName');
}
function _Label9_ctor()
{
	CreateLabel('Label9', _initLabel9, null);
}
</script>
<% Label9.display %>

<!--METADATA TYPE="DesignerControl" endspan-->                     
</FONT></TD>
</tr>
<TR>
	<TD><FONT face=Verdana>Contact Email </FONT></td>
	<TD><FONT face=Verdana>
 <!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=Label10 style="HEIGHT: 17px; LEFT: 10px; TOP: 266px; WIDTH: 84px" 
	width=84>
	<PARAM NAME="_ExtentX" VALUE="2223">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="Label10">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="OrgContactEmail">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE=""></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel10()
{
	Label10.setDataSource(Recordset1);
	Label10.setDataField('OrgContactEmail');
}
function _Label10_ctor()
{
	CreateLabel('Label10', _initLabel10, null);
}
</script>
<% Label10.display %>

<!--METADATA TYPE="DesignerControl" endspan-->                       
</FONT></TD>
</tr>
<TR>
	<TD><FONT face=Verdana>HomePage </FONT></td>
	<TD><FONT face=Verdana>                        
<a href="
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=Label12 style="HEIGHT: 17px; LEFT: 10px; TOP: 283px; WIDTH: 92px" 
	width=92>
	<PARAM NAME="_ExtentX" VALUE="2434">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="Label12">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="OrgHomePageURL">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE=""></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel12()
{
	Label12.setDataSource(Recordset1);
	Label12.setDataField('OrgHomePageURL');
}
function _Label12_ctor()
{
	CreateLabel('Label12', _initLabel12, null);
}
</script>
<% Label12.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
">
<!--METADATA TYPE="DesignerControl" startspan
<OBJECT classid="clsid:B5F0E460-DC5F-11D0-9846-0000F8027CA0" height=17 id=Label12 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 92px" 
	width=92>
	<PARAM NAME="_ExtentX" VALUE="2434">
	<PARAM NAME="_ExtentY" VALUE="450">
	<PARAM NAME="id" VALUE="Label12">
	<PARAM NAME="DataSource" VALUE="Recordset1">
	<PARAM NAME="DataField" VALUE="OrgHomePageURL">
	<PARAM NAME="FontFace" VALUE="">
	<PARAM NAME="FontSize" VALUE="">
	<PARAM NAME="FontColor" VALUE="">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="Visible" VALUE="-1">
	<PARAM NAME="FormatAsHTML" VALUE="0">
	<PARAM NAME="Platform" VALUE="0">
	<PARAM NAME="LocalPath" VALUE=""></OBJECT>
-->
<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function _initLabel12()
{
	Label12.setDataSource(Recordset1);
	Label12.setDataField('OrgHomePageURL');
}
function _Label12_ctor()
{
	CreateLabel('Label12', _initLabel12, null);
}
</script>
<% Label12.display %>

<!--METADATA TYPE="DesignerControl" endspan-->
</a>
</FONT></TD>
</TR>
</TABLE>
</td>
</tr>
<tr>
<td width=180 bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="http:forums/default.htm" target="_top" vlink="White">Forums</a></strong></big></font></center></td>
<td rowspan="4"></td>
</tr>
<tr>
<td width=180 bgcolor="#008080"><center><font face="Verdana"><big><strong><A href="http:Admin/Login_Guts_Frame.asp?mode=1" target="_top" vlink="White">Administration</a></strong></big></font></center></td>
<td rowspan="4"></td>
<td rowspan="4"></td>
</tr>
<tr>
<td></td>
<td rowspan="4"></td>
</tr>
</table>
<%	' *** If there is a URL to the logo, show it ***
	if Recordset1.fields.getValue("OrgLogoURL") <> "" then
%>
<img SRC="<%=Recordset1.fields.getValue("OrgLogoURL")%>" ALIGN=RIGHT>
<%
	end if
%>
<!-- #include File="copyright.asp" -->
</body>
<% ' VI 6.0 Scripting Object Model Enabled %>
<% EndPageProcessing() %>
</FORM>
</html>
