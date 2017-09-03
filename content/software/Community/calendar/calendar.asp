<%@ Language=VBScript %>
<!-- #include file="../adovbs.inc" -->
<!-- #include File="../mapvar.asp" -->
<% thisYear = 0 
	Response.Buffer=true
	Response.Expires=0
%>
<HTML>
<HEAD>
</HEAD>
<%

	Set conn = Server.CreateObject("ADODB.Connection")
	Set rsSiteInfo = Server.CreateObject("ADODB.Recordset")
	
	
	' Open the default calendar
	conn.Open Session("Connection1_ConnectionString")
		
	' Open the SiteInformation Recordset
	rsSiteInfo.Open "SiteInfo", conn, adOpenForwardOnly, adLockReadOnly

	GetDateInfo curMonth, curDay, curYear, DateStr

	' Get the view
	CalendarView =  Request.QueryString("CalendarView") & ""
	
	' If no view is passed, then check for it in session
	If CalendarView = "" and session("CalendarView") <> "" then
		CalendarView = Session("CalendarView")
	End If
	
	' Response.Write Request.QueryString("CalendarView")
	' Response.Write CalendarView
	
	' Determine what to set the view to...
	Select Case CalendarView
		Case "Month"
			Session("CalendarView") = "Month"
		Case "BigMonth"
			Session("CalendarView") = "BigMonth"
		Case "List"
			Session("CalendarView") = "List"
		Case Else
			if rsSiteInfo("CalDefaultMode") <> "" then
				Session("CalendarView") = rsSiteInfo("CalDefaultMode")
				CalendarView = rsSiteInfo("CalDefaultMode")
			else
				Session("CalendarView") = "Month"
				CalendarView = "Month"
			end if
			
	End Select
	
	' Response.Write CalendarView
	
	' Render the correct view

	Select Case CalendarView
		Case "Month"
			If Request.QueryString("FrameType") = "Left" Then
				Response.Write PrintMonth(False)
			End If
			
			If Request.QueryString("FrameType") = "Right" Then
				ShowList(Conn)
			End If
			
			If Request.QueryString("FrameType") = "" Then
				%>
				<frameset cols="250,*">
				<frame name="contents" target="right" 
<%
Response.Write "src=""calendar.asp?CalendarView=Month&FrameType=Left&Date=" & server.urlencode(DateStr) & """>"
%>
				<frame name="right" 
<%
Response.Write "src=""calendar.asp?CalendarView=Month&FrameType=Right&Date=" & server.urlencode(DateStr) & """>"
%>
				<noframes>
				<body>
				<p>This page uses frames, but your browser doesn't support them.
				</body>
				</noframes>
				</frameset>
				<%
			End If
		
		Case "BigMonth"
			Response.Write PrintMonth(True)
		Case "List"
			ShowList(Conn)

	End Select
	rsSiteInfo.Close
	Set rsSiteInfo = Nothing
	conn.Close
	Set conn=Nothing
	

	%>
<!-- #include File="../copyright.asp" -->
</HTML>
</BODY>
<Script Language=vbscript runat=server>

Function MakeASPURL(Parameters)
   GetDateInfo thisMonth, thisDay, thisYear, Datestr
   Temp = "calendar.asp?" & Parameters
   Call AddURLParm(Temp,"FrameType","")
   Call AddURLParm(Temp,"CalendarView","")
   Call AddURLParm(Temp,"Detail","")
   Call AddURLParm(Temp,"RecStart","")
   Call AddURLParm(Temp,"Date",DateStr)
   MakeASPURL = Temp
End Function

Sub AddURLParm(Temp,QString,Value)

   TmpVal = Value
   If Len(Trim(TmpVal)) < 1 Then
	TmpVal = Request.QueryString(QString)
   End If
   If TmpVal <> "" Then
      If InStr(1,Temp,QString&"=") = 0 Then  ' Dont add a parameter twice
        If Right(Temp,1) <> "?" Then Temp = Temp & "&"
        Temp = Temp & QString  & "=" & server.urlencode(TmpVal)
      End If
   End If
End Sub

Function PrintMonth(BigCal)
	Set cal = new eventCalendar
	
	Cal.DSN = Session("Connection1_ConnectionString")
	
	If BigCal then
		Cal.TableAttributes = " border=1 cellspacing=0 cellpadding=2 width=100% bgcolor=#ffffff"
	Else
		Cal.TableAttributes = " border=0 cellspacing=0 cellpadding=2"
	End If
	
	GetDateInfo curMonth, curDay, curYear, DateStr
	
	Cal.CurrentDate = DateStr

	If BigCal then
		Cal.DayLinkURL = "calendar.asp?CalendarView=List"
'		Cal.DayLinkTarget = "_top"
		Cal.DayLinkTarget = ""
	Else
		Cal.DayLinkURL = "calendar.asp?CalendarView=Month&FrameType=Right"
		Cal.DayLinkTarget = "right"
	end if

        ' Response.Write request.QueryString("action")

	If request.QueryString("action") = "nextmonth" then
		Cal.MoveMonth(1)
	End If
	If request.QueryString("action") = "prevmonth" then
		Cal.MoveMonth(-1)
	End If

	Cal.nextURL = MakeASPURL("action=nextmonth&Date=" & server.urlencode(Cal.CurrentDate))
	Cal.prevURL = MakeASPURL("action=prevmonth&Date=" & server.urlencode(Cal.CurrentDate))
	

	If BigCal then
		PrintMonth = "<body><table border=1 cellspacing=0 cellpadding=8 width=100%  bgcolor=#008080 ><tr><td>"
	Else
		PrintMonth = "<body><table border=1 cellspacing=0 cellpadding=2><tr><td>"
	End If

	If BigCal then
		Cal.CellHeight = "70"
		Cal.CellWidth = "14%"
	end if
	PrintMonth = PrintMonth & Cal.DrawCalendar(BigCal)

	PrintMonth = PrintMonth & "</td></tr></table>"

	If BigCal Then
		PrintMonth = PrintMonth & vbCR & ShowViews(Cal.CurrentDate,True,False,True,"","") & ShowHome()
	Else
		PrintMonth = PrintMonth & vbCR & ShowViews(Cal.CurrentDate,False,True,True,"","parent.") & ShowHome() & "<BR>" & VBCRLF
  	End if

End Function

Function ShowViews(CurrentDate,Month,BigMonth,List,ViewOpt,Target)

	views = 0

	ShowViews =  "<select name=diffview onChange=""" & Target & "window.location=this.options[this.selectedIndex].value;this.selectedIndex=0;"">" & vbCrLf
	ShowViews = ShowViews &  "<option value="""">View</option>"

	If rsSiteInfo("CalAllowListView") = true AND List = True then
		ShowViews = ShowViews & "<option value=""calendar.asp?CalendarView=List&Date=" & server.urlencode( CurrentDate) & """>List</option>" & vbCrLf
		views = views + 1
	End If

	If rsSiteInfo("CalAllowBigMonthView") = true AND BigMonth=True then
		ShowViews = ShowViews & "<option value=""calendar.asp?CalendarView=BigMonth&Date=" & server.urlencode( CurrentDate) & """>Big Month</option>" & vbCrLf
		views = views + 1
	End If

	If rsSiteInfo("CalAllowMonthView") = true AND Month=True then
		ShowViews = ShowViews & "<option value=""calendar.asp?CalendarView=Month&Date=" & server.urlencode( CurrentDate) & """>Month</option>" & vbCrLf
		views = views + 1
	End If

	If Len(Trim(ViewOpt)) > 0 Then
		ShowViews = ShowViews & ViewOpt
		views = views + 1
  	End If

	ShowViews = ShowViews & "</select>" & vbCrLf

        If Views < 1 Then
		ShowViews = ""
	End If
End Function

Function ShowHome()

	ShowHome = ""

	whereto = "../index.asp"
	If rsSiteInfo("HomeURL") <> "" then
		whereto = rsSiteInfo("HomeURL")	
	End If

	If lcase(whereto) <> "none" then
		ShowHome = ShowHome & "<input type=button value=""Home"" onClick=""top.window.location='" & whereto & "';"" id=button1 name=button1>"
	End If

End Function

Function ShowList(Conn)

	Set rsEvents = Server.CreateObject("ADODB.Recordset")

        GetDateInfo thisMonth, thisDay, thisYear, Datestr

	SQL = "SELECT Events.* FROM Events WHERE MONTH(EventDate) = "& cstr(thisMonth) 
	SQL = SQL & " AND YEAR(EventDate) = " & cstr(thisYear)
	
	If thisDay <> 0 then
		SQL =  "SELECT Events.* FROM Events WHERE EventDate >= #"& cstr(thisMonth) & "/" & cstr(thisDay) & "/" & cstr(thisYear) & "#"
	end if

        SQL = SQL & " AND (EventPublishDate < #"& Now()&"# OR EventPublishDate is Null)"

	If Request("lstEventType") <> "" then
		SQL = SQL & " AND EventTypeID=" & cstr(Request("lstEventType"))
	End If
	
	If Request("lstAudience") <> "" then
		SQL = SQL & " AND AudienceID=" & cstr(Request("lstAudience"))
	End If
	
	If rsSiteInfo("CalApproval") = True Then
		SQL = SQL & " AND ApproveUserID > 0 "
	End If
		
	If trim(Request("keyword")) <> "" then
		SQL =  SQL & " AND (EventName like '%" & trim(Request("keyword")) & "%' or EventDescription like '%" & trim(Request("keyword")) & "%')"
	End If
	
	SQL = SQL & " ORDER BY EventDate, EventTimeStart"

	rsEvents.Open SQL, Conn, adOpenStatic

	CalLength = Session("Detail")

	If Request.QueryString("Detail") = "Long" then
		CalLength = "Long"
		Session("Detail") = "Long"
	End If
	
	If Request.QueryString("Detail") = "Short" then
		CalLength = "Short"
		Session("Detail") = "Short"
	End If

	If CalLength = "" then
		CalLength = "Short"
		Session("Detail") = "Short"
	End If

	Response.Write "<body>"	

	If Request.QueryString("FrameType") <> "Right" Then
		if Trim(rsSiteInfo("OrgLogoURL")) <> "" then
			 Response.Write "<img SRC=" &chr(34) &  rsSiteInfo("OrgLogoURL") & chr(34) & " ALIGN=RIGHT>"  & vbCRLF
		End If
		if Trim(rsSiteInfo("OrgName")) <> "" Then
			Response.Write "<h3><font face=""helvetica"">"
			Response.Write rsSiteInfo("OrgName")
			Response.Write "</font></h3>"  & vbCrLF
		End If
	End If

	' Figure out where we are supposed to start

	RecCount = rsEvents.RecordCount
	RecCurrent = 1
	
	RecStart =1
	RecDir = ""	
	if CInt(Request.QueryString("RecStart")) > 0 Then
 		RecStart = CInt(Request.QueryString("RecStart")) + 1
		RecDir = Request.QueryString("Direction")
	End If
	
	Response.Write "<font face=""helvetica"" size=""-1"">"

    	GetDateInfo curMonth, curDay, curYear, Datestr
  
	If Request.QueryString("FrameType") = "Right" Then
		Response.Write DisplayNextPrev (RecStart,RecCount,curYear,curMonth)
	Else 
 		Response.Write "<form action=calendar.asp method=post>"	& vbCrLF
	    	Response.Write "<input type=submit value=""Go:"">" & vbCrLF
		
		Call SelectDay(curMonth,curDay,curYear)

		If rsSiteInfo("CalEventType") = True then
			Response.Write "Event Type "
			Set rsEventType = Server.CreateObject("ADODB.Recordset")
			rsEventType.Open "EventType", conn, adOpenForwardOnly, adLockReadOnly
			Response.Write DropDown("lstEventType", "EventTypeID", "EventTypeName", rsEventType)
			rsEventType.Close
			Set rsEventType = Nothing
		End If

		If rsSiteInfo("CalAudience") = True then	
			Response.Write "Audience "
			Set rsAudience = Server.CreateObject("ADODB.Recordset")
			rsAudience.Open "Audience", conn, adOpenForwardOnly, adLockReadOnly
			Response.Write DropDown("lstAudience", "AudienceID", "AudienceName", rsAudience)
			rsAudience.Close
			Set rsEventType = Nothing
		End If 	
	
		If CalLength = "Short" then
			ViewOpt = "<option value=""" & MakeASPURL("Detail=Long") & """>More Detail</option>"
		End If
		If CalLength = "Long" then
			ViewOpt = "<option value=""" & MakeASPURL("Detail=Short") & """>Less Detail</option>"
		End If

		Response.Write vbCrLf & "Search:" & vbCrLf
		Response.Write "<input type=text name=keyword>"

		Response.Write ShowViews(DateStr,True,True,False,ViewOpt,"")

		Response.Write DisplayNextPrev (RecStart,RecCount,curYear,curMonth)

		Response.Write("</form>")
	End If  'End of Frametype <> Right

	
	Response.Write vbCrLF & "<hr>" 
	
	If RecDir = "Prev" then
		RecStart = RecStart - 10
		If RecStart < 1 then RecStart = 1
	End If
		
	Do while not RsEvents.eof
	
		If (RecCurrent < RecStart or RecCurrent > RecStart + 9)  Then
			' Dont print a record		
		Else
			Response.Write vbCrLf
			If CalLength = "Long" then
				Response.Write("<Table border=0 width=300 cellpadding=2><tr><td colspan=2 bgcolor=#008080>" & vbcrlf)
				Response.Write("<font face=Helvetica color=White size=+1><b>" & formatDateTime(rsEvents("EventDate"),1) & "</b></font><br></td></tr>"& vbcrlf) 
			Else
				Response.Write("<Table border=0 width=350 cellpadding=2><tr><td colspan=6 bgcolor=#008080>"& vbcrlf)
				Response.Write("<font face=Helvetica color=White size=+1><b>" & formatDateTime(rsEvents("EventDate"),1) & "</b></font><br></td></tr></table>"& vbcrlf) 
			End If
			Response.Write vbCrLf

			DetailURL = "<input type=button onclick=""window.location.href='"
		
			If CalLength = "Short" then
				DetailURL = DetailURL & MakeASPURL("Detail=Long&RecStart=" & (RecCurrent-1))
				DetailURL = DetailURL &  "'"" value='More Detail'>" & vbCrLF
			End If
			If CalLength = "Long" then
				DetailURL = DetailURL &  MakeASPURL("Detail=Short&RecStart=" & (RecCurrent-1))
				DetailURL = DetailURL &  "'"" value='Less Detail'>" & vbCrLF
			End If

			' We don't want the individual Less/More in the basic list view
			If Request.QueryString("FrameType") <> "Right" Then
				DetailURL = ""
			End If

			DisplayEvent rsSiteInfo, rsEvents, CalLength, DetailURL
			Response.Write vbCrLf
			' Response.Write " " & RecCurrent & "<br>"
			
			Response.Write("</table><br>")
			
		End If
		RecCurrent = RecCurrent + 1
		rsEvents.Movenext
	Loop
	
	Response.Write DisplayNextPrev (RecStart,RecCount,curYear,curMonth)
	
	
	Response.Write "<BR>"
	If Request.QueryString("FrameType") <> "Right" Then
		Response.Write ShowHome()	
	End If

	rsEvents.Close
	Set rsEvents = Nothing
	
End Function

Function DisplayNextPrev(RecStart,RecCount,curYear,curMonth)



	If RecStart > 9 then
		ptmpStr = "<input type=button onclick=""window.location.href='"
		ptmpStr = ptmpStr & "calendar.asp?RecStart="&Cstr(RecStart-11) & "&Direction=Prev"
		If Request.QueryString("FrameType") = "Right" Then
 			ptmpStr = ptmpStr & "&CalendarView=Month&FrameType=Right"
		End If
		ptmpStr = ptmpStr & "'" & chr(34)
		ptmpStr = ptmpStr & "Value=" & chr(34) & "Previous 10 Events" & chr(34) & ">"
	End If	
		
	If RecStart + 10 > RecCount then
		'
	Else
		NextCount = RecCount - RecStart - 9
		If NextCount > 10 then NextCount = 10
		ntmpStr = "<input type=button onclick=""window.location.href='"
		ntmpStr = ntmpStr & "calendar.asp?RecStart="&Cstr(RecStart+9) 
		If Request.QueryString("FrameType") = "Right" Then
 			ntmpStr = ntmpStr & "&CalendarView=Month&FrameType=Right"
		End If
		ntmpStr = ntmpStr & "'" & chr(34)
		ntmpStr = ntmpStr & "Value=" & chr(34) & "Next "&cstr(NextCount) &" Events" & chr(34) & ">"
	End If 

	'Response.Write ptmpstr & ntmpstr
	DisplayNextPrev = ptmpstr & ntmpstr

End Function

Sub SelectDay(curMonth,curDay,curYear)

	Response.Write("<Select Name=Month>")
	For cnt = 1 to 12
		Response.Write("<option value="&cstr(cnt))
		If cnt = curMonth then
			Response.write " selected "
		End If
		Response.Write(">" &cstr(monthname(cnt,True))&"</option>" & vbcrlf)
	Next	
	Response.Write("</select>")

	' Show the day
	Response.Write("<Select Name=Day>")
	Response.Write ("<option></option>")
	For cnt = 1 to 31
		Response.Write("<option value="&cstr(cnt))
		If cnt = curDay then
			Response.write " selected "
		End If
		Response.Write(">" &cstr(cnt)&"</option>" & vbcrlf)
	Next
	Response.Write("</select>")

	' Show the year + 5
	Response.Write("<Select Name=Year>")
	For cnt = -1 to 5
		Response.Write("<option value="&cstr(cnt+Year(date())))
		If cnt+Year(Date()) = cint(curYear) then
			'Response.Write cstr(cnt+Year(Date())) &"--"& cstr(cint(curYear))
			Response.write " selected "
		End If
		Response.Write(">" &cstr(cnt+Year(date()))&"</option>" & vbcrlf)
	Next
	Response.Write("</select>")
End Sub



function DropDown(name, keyField, displayfield, rs, addBlank)

' //\\ Generic function to create a listbox from a recordset //\\

	Dim tmpText
	
	tmpText = "<select name=" & chr(34) & name & chr(34) & ">" & vbcrlf
	
	if addBlank = true then
		tmpText = tmpText & "<option value=" & chr(34) &  chr(34) & ">&nbsp;" & "</option>" & vbcrlf
	End If
	
	do until rs.EOF
		tmpText = tmpText & "<option value=" & chr(34) & rs(keyField) & chr(34) & ">" & vbcrlf
		tmpText = tmpText & rs(displayField) &  "</option>" & vbcrlf
		rs.movenext
	loop
	
	tmpText = tmpText & "</select>" & vbcrlf
	
	DropDown = tmpText

end function


Sub GetDateInfo(thisMonth, thisDay, thisYear, DateStr)

'  First we check today's date, then the query string, but the highest
'  Priority (checked last) is the drop-downs from the form
'  It might be a good idea to check to see if RecStart is set and use
'  That for the date - It would require a database lookup

	thisMonth = Cint(Month(Now()))
	thisYear = Cint(Year(Now()))
        thisDay = Cint(Day(Now()))

	If Request.QueryString("Date") <> "" Then
		On Error Resume Next
  		DateVal = CDate(Request.QueryString("Date"))
		If Err.Number = 0 Then
			thisMonth = Cint(Month(DateVal))
			thisYear = Cint(Year(DateVal))
	        	thisDay = Cint(Day(DateVal))
		End If
		On Error Goto 0 
	End If

	If Request("Month") <> "" then
		thisMonth = Cint(Request("Month"))
	End If
	If Request("Day") <> "" then
		thisDay = Cint(Request("Day"))
	End If
	If Request("Year") <> "" then
		thisYear = Cint(Request("Year"))
	End If

' Clean up Invalid Dates
	
	datestr = thisMonth & "/" & thisDay & "/" & ThisYear
	While Not IsDate(datestr) and thisDay > 1 
           thisDay = thisDay - 1
           datestr = thisMonth & "/" & thisDay & "/" & ThisYear
        Wend
End Sub

Dim DisplayFieldLength ' Sorry - Global Variable

Sub DisplayEvent(rsSiteInfo, rsEvents,Length,DetailURL)

    	DisplayFieldLength = Length
    
	Response.Write DisplayField("Event: ",cstr(rsEvents("EventName")))
	Response.Write DisplayField("Details: ",rsEvents("EventDescription"))
	Response.Write DisplayField("Start&nbsp;Time: ",rsEvents("EventTimeStart"))
	Response.Write DisplayField("End&nbsp;Time: ",rsEvents("EventTimeEnd"))
	If Length = "Long" Then
		Response.Write DisplayField("Location:",rsEvents("Location"))
		Response.Write DisplayField("Contact:",rsEvents("Contact"))
		Response.Write DisplayField("Phone:",rsEvents("ContactPhone"))
		Response.Write DisplayField("Email:",rsEvents("ContactEmail"))
		If rsEvents("AdmissionCharge") = vbTrue Then
			Response.Write DisplayField("Admission:",rsEvents("AdmissionAmount"))
		End If
		Response.Write DisplayField("Directions:",rsEvents("TextDirections"))
		Response.Write DisplayURL("Map:",rsEvents("URLMap"),"Directions")
                Response.Write DisplayURL("Additional Info:",rsEvents("URLInfo"),rsEvents("URLName"))

		If rsEvents("URLPhoto") <> "" Then
			DescStr = "<img src=""" & rsEvents("URLPhoto") & """ WIDTH=100>"
			DescStr = "<a href=""" & rsEvents("URLPhoto") & """ target=_new>" & DescStr & "</a>"
			Response.Write DisplayField("Photo:",DescStr)
		End If
		
		Set rs = server.CreateObject("ADODB.Recordset")
	
		If rsSiteInfo("CalEventType") = True then	
			rs.Open "SELECT * FROM EventType WHERE EventTypeID="&cstr(0+rsEvents("EventTypeID")),Conn
			If not rs.EOF then
				If rs("EventTypeImageURL") <> "" then
					Response.Write DisplayField("Event Type:",rs("EventTypeName") & "<br> <img src=" & chr(34) & rs("EventTypeImageURL") & chr(34) & ">")
				Else
					Response.Write DisplayField("Event Type:",rs("EventTypeName"))
				End If
	
			End If
			rs.Close
		End If
	
		If rsSiteInfo("CalAudience") = True then	
			rs.Open "SELECT * FROM Audience WHERE AudienceID="&cstr(rsEvents("AudienceID")),Conn
			tmpStr = ""
			Do while not rs.EOF
				tmpStr = tmpStr & rs("AudienceName") & "<br>" & vbcrlf
				If rs("AudienceImageURL") <> "" then
					tmpStr = tmpStr & "<img src=" & chr(34) & rs("AudienceImageURL") & chr(34) & "><br>" & vbcrlf
				End If
				rs.movenext
			Loop
			Response.Write DisplayField("Audience:",tmpStr)
			rs.Close
		End If
	
		set rs=nothing
                If rsEvents("UserField1") <> "" AND rsSiteInfo("UserField1Label") <> "" then
			Response.Write DisplayField(rsSiteInfo("UserField1Label"),rsEvents("UserField1"))
		End If
                If rsEvents("UserField2") <> "" AND rsSiteInfo("UserField2Label") <> "" then
			Response.Write DisplayField(rsSiteInfo("UserField2Label"),rsEvents("UserField2"))
		End If
                If rsEvents("UserField3") <> "" AND rsSiteInfo("UserField3Label") <> "" then
			Response.Write DisplayField(rsSiteInfo("UserField3Label"),rsEvents("UserField3"))
		End If
		
	End If

	If DisplayFieldLength = "Long" then
		Response.Write "<tr><td width=85 valign=top>"
	End If
	If DetailURL <> "" Then
		Response.Write(DetailURL)
	End if
	If DisplayFieldLength = "Long" then
		Response.Write "</td></tr>"
	End If


End Sub

Function DisplayField(FieldTitle,FieldValue)
	DisplayField = ""
	If Trim(FieldValue) <> "" then
		DisplayField = "<font face=helvetica size=-1><b>" & FieldTitle & "</b></font></td><td width=215><font face=helvetica size=-1>" & FieldValue & "</font>&nbsp;" 

		If DisplayFieldLength = "Long" then
			DisplayField = "<tr><td width=85 valign=top>" & DisplayField & "</td></tr>"
		End If
 		DisplayField = DisplayField & vbcrlf
	End If
End Function

Function DisplayURL(FieldTitle,FieldValue,FieldHref)
	DisplayURL = ""
	If Trim(FieldValue) <> "" then
		DisplayURL = "<font face=helvetica size=-1><b>" & FieldTitle & "</b></font></td><td width=215><font face=helvetica size=-1><a href=""" & FieldValue & """ target=""_new"">" & FieldHref & "</a></font>&nbsp;" 

		If DisplayFieldLength = "Long" then
			DisplayURL = "<tr><td width=85 valign=top>" & DisplayURL & "</td></tr>"
		End If
 		DisplayURL = DisplayURL & vbcrlf
	End If
End Function

Class EventCalendar

Private mvarCurrentDate 
Private mvarCellAttributes 
Private mvarRowAttributes 
Private mvarTableAttributes 
Private mvarDSN 
Private mvarFont 
Private mvarnextURL 
Private mvarprevURL 
Private mvarDayLinkURL 
Private mvarDayLinkTarget 
Private mvarHeaderAttributes 
Private mvarCellHeight 
Private mvarCellWidth 

Public Property Let CellWidth(vData)
    mvarCellWidth = vData
End Property

Public Property Set CellWidth( vData )
    Set mvarCellWidth = vData
End Property

Public Property Get CellWidth() 
    If IsObject(mvarCellWidth) Then
        Set CellWidth = mvarCellWidth
    Else
        CellWidth = mvarCellWidth
    End If
End Property

Public Property Let CellHeight( vData )
    mvarCellHeight = vData
End Property

Public Property Set CellHeight( vData )
    Set mvarCellHeight = vData
End Property

Public Property Get CellHeight() 
    If IsObject(mvarCellHeight) Then
        Set CellHeight = mvarCellHeight
    Else
        CellHeight = mvarCellHeight
    End If
End Property

Public Property Let HeaderAttributes( vData )
    mvarHeaderAttributes = vData
End Property

Public Property Set HeaderAttributes( vData )
    Set mvarHeaderAttributes = vData
End Property

Public Property Get HeaderAttributes() 
    If IsObject(mvarHeaderAttributes) Then
        Set HeaderAttributes = mvarHeaderAttributes
    Else
        HeaderAttributes = mvarHeaderAttributes
    End If
End Property

Public Property Let DayLinkTarget( vData )
    mvarDayLinkTarget = vData
End Property

Public Property Set DayLinkTarget( vData )
    Set mvarDayLinkTarget = vData
End Property

Public Property Get DayLinkTarget() 
    If IsObject(mvarDayLinkTarget) Then
        Set DayLinkTarget = mvarDayLinkTarget
    Else
        DayLinkTarget = mvarDayLinkTarget
    End If
End Property

Public Property Let DayLinkURL( vData )
    mvarDayLinkURL = vData
End Property

Public Property Set DayLinkURL( vData )
    Set mvarDayLinkURL = vData
End Property

Public Property Get DayLinkURL() 
    If IsObject(mvarDayLinkURL) Then
        Set DayLinkURL = mvarDayLinkURL
    Else
        DayLinkURL = mvarDayLinkURL
    End If
End Property

Public Property Let prevURL( vData )
    mvarprevURL = vData
End Property

Public Property Set prevURL( vData )
    Set mvarprevURL = vData
End Property

Public Property Get prevURL() 
    If IsObject(mvarprevURL) Then
        Set prevURL = mvarprevURL
    Else
        prevURL = mvarprevURL
    End If
End Property

Public Property Let nextURL( vData )
    mvarnextURL = vData
End Property

Public Property Set nextURL( vData )
    Set mvarnextURL = vData
End Property

Public Property Get nextURL() 
    If IsObject(mvarnextURL) Then
        Set nextURL = mvarnextURL
    Else
        nextURL = mvarnextURL
    End If
End Property

Public Sub MoveMonth(Direction )
    Dim dtCurrent
  
    dtCurrent = mvarCurrentDate
' Response.Write "Old "
' Response.write mvarCurrentDate
    mvarCurrentDate = CStr(DateSerial(Year(dtCurrent), Month(dtCurrent) + CInt(Direction), Day(dtCurrent)))
' Response.Write " New "
' Response.write mvarCurrentDate
End Sub

Public Property Let Font( vData )
    mvarFont = vData
End Property

Public Property Set Font( vData )
    Set mvarFont = vData
End Property

Public Property Get Font() 
    If IsObject(mvarFont) Then
        Set Font = mvarFont
    Else
        Font = mvarFont
    End If
End Property

Public Property Let DSN( vData )
    mvarDSN = vData
End Property

Public Property Get DSN() 
    DSN = mvarDSN
End Property

Public Property Let TableAttributes( vData )
    mvarTableAttributes = vData
End Property

Public Property Get TableAttributes() 
    TableAttributes = mvarTableAttributes
End Property

Public Property Let RowAttributes( vData )
    mvarRowAttributes = vData
End Property

Public Property Get RowAttributes() 
    RowAttributes = mvarRowAttributes
End Property

Public Property Let CellAttributes( vData )
    mvarCellAttributes = vData
End Property

Public Property Get CellAttributes() 
    CellAttributes = mvarCellAttributes
End Property

Public Property Let CurrentDate( vData )
    mvarCurrentDate = vData
End Property

Public Property Get CurrentDate() 
    CurrentDate = mvarCurrentDate
End Property

' Private Functions
Private Function PrintTable() 

    PrintTable = "<table " & CStr(mvarTableAttributes) & " >" & vbCrLf

End Function

Private Function PrintRow( moreRowAttributes ) 

    PrintRow = "<tr " & CStr(mvarRowAttributes) & " " & CStr(moreRowAttributes) & " >" & vbCrLf
    
End Function

Private Function PrintCell(CellContents ,  moreCellAttributes ,  fontAttributes ) 

    PrintCell = "<td " & CStr(mvarCellAttributes) & " " & CStr(moreCellAttributes) & " >"
    PrintCell = PrintCell & "<Font Face=" & Chr(34) & mvarFont & Chr(34) & " " & fontAttributes & " > " & CellContents & "</font></td>" & vbCrLf
    
End Function

Private Function PrintRowEnd() 

    PrintRowEnd = "</tr>"

End Function

Private Function PrintTableEnd() 

    PrintTableEnd = "</table>"

End Function

Private Sub Class_Initialize()

    mvarFont = "Verdana"
    mvarCellWidth = "14%"
    mvarCellHeight = "14%"
    
End Sub


' Public Functions (aka Methods)
Public Function DrawCalendar(BigCal)

	Dim strCal 
    Dim intWeeks
    Dim intDays 
    Dim intDay 
    Dim intMonth
    Dim intYear
    Dim dtCurrent 
    Dim dtFirstDay 
    Dim dtLastDay
    Dim dtLastDayLastMonth
    Dim strCell 
    Dim intStartingDay 
    Dim intStartingMonth 
    Dim intLastDay 
    Dim intLastDayLastMonth 
    Dim intWeekDayFirst 
    Dim intLoopDay 
    Dim intLoopMonth 
    Dim objContext 
    Dim strTmpCell 
    Dim objServer 
    Dim rsEvents 
    Dim Conn 
    Dim numevents 
    

	Set Conn = Server.CreateObject("ADODB.Connection")
	Set rsEvents = Server.CreateObject("ADODB.RecordSet")
	       
    dtCurrent = mvarCurrentDate
    dtFirstDay = DateSerial(Year(dtCurrent), Month(dtCurrent), 1)
    dtLastDay = DateSerial(Year(dtCurrent), Month(dtCurrent) + 1, 0)
    dtLastDayLastMonth = DateSerial(Year(dtCurrent), Month(dtCurrent), 0)
    
    intDay = Day(dtCurrent)
    intMonth = Month(dtCurrent)
    intYear = Year(dtCurrent)
    
    intWeekDayFirst = Weekday(dtFirstDay)
    intLastDayLastMonth = Day(dtLastDayLastMonth)
    If intWeekDayFirst > 1 Then
        intStartingDay = intLastDayLastMonth - (intWeekDayFirst - 2)
        intStartingMonth = Month(DateSerial(Year(dtCurrent), Month(dtCurrent), 0))
    Else
        intStartingDay = 1
        intStartingMonth = intMonth
    End If
    intLastDay = Day(dtLastDay)
        
    ' Print the table tag
    strCal = PrintTable()
    ' Print the row tag
    strCal = strCal & PrintRow("")
    If BigCal = True Then
        strCell = "<center><a href=" & Chr(34) & mvarprevURL & Chr(34) & "><img src=prev.gif border=0></a></center>"
    Else
        strCell = "<a href=" & Chr(34) & mvarprevURL & Chr(34) & "><img src=prev.gif border=0></a>"
    End If
    'Print the Prev Month Link
    strCal = strCal & PrintCell(strCell,"","")
    ' Print the name of the current Month
    If BigCal = True Then
        strCal = strCal & PrintCell("<center>" & MonthName(Month(dtCurrent)) & " " & cstr(Year(dtCurrent)) & "</center>", " colspan=5 halign=center valign=center ","")
    Else
        strCal = strCal & PrintCell(MonthName(Month(dtCurrent)) & " " & cstr(Year(dtCurrent)), " colspan=5 halign=center valign=center ", "")
    End If
    
    If BigCal = True Then
        strCell = "<center><a href=" & Chr(34) & mvarnextURL & Chr(34) & "><img src=next.gif border=0></a></center>"
    Else
        strCell = "<a href=" & Chr(34) & mvarnextURL & Chr(34) & "><img src=next.gif border=0></a>"
    End If
    ' Print the Next Month link
    strCal = strCal & PrintCell(strCell,"","")
    
    strCal = strCal & PrintRowEnd()
    
    strCal = strCal & PrintRow("")
    
    For intDays = 1 To 7
        If BigCal = True Then
            strCell = WeekdayName(intDays, False)
            strCal = strCal & PrintCell(strCell, " width=" & mvarCellWidth & " height=" & mvarCellHeight & " ", "")
        Else
            strCell = Left(WeekdayName(intDays), 1)
            strCal = strCal & PrintCell(strCell,"","")
        End If
        
    Next
    
    strCal = strCal & PrintRowEnd
    
    If BigCal <> True Then
        strCal = strCal & PrintRow(" height=1 ")
        strCal = strCal & PrintCell("<center><hr><center>", " colspan=7 align=center valign=center height=1 ","")
        strCal = strCal & PrintRowEnd()
    End If
    
    intLoopDay = intStartingDay
    intLoopMonth = intStartingMonth
    
    If mvarDSN <> "" Then
        Conn.Open mvarDSN
    End If
    ' Loop through the weeks
    For intWeeks = 1 To 6
        strCal = strCal & PrintRow("")
    
        ' Loop through the days
        For intDays = 1 To 7
            
            strCell = CStr(intLoopDay)
            
            ' Bold the current month
            If intLoopMonth <> intMonth Then
                If BigCal = True Then
                    strCell = strCell
                    '& "<br>&nbsp;<br>&nbsp;"
                End If
                strCal = strCal & PrintCell(strCell, " valign=top ", " COLOR=Gray ")
            Else
				if instr(1,mvarDayLinkURL,"?") then 
					sepStr = "&"
				else
					sepStr = "?"
				end if
                strTmpCell = "<a href=" & Chr(34) & mvarDayLinkURL & sepStr & "Date="
                strTmpCell = strTmpCell & Server.urlencode(CStr(intMonth) & "/" & CStr(intLoopDay) & "/" & CStr(intYear)) & Chr(34)
                If mvarDayLinkTarget <> "" Then
                    strTmpCell = strTmpCell & " target=" & Chr(34) & CStr(mvarDayLinkTarget) & Chr(34)
                End If
                
                
                If mvarDSN <> "" Then
                    
                    Set rsEvents = Conn.Execute("SELECT * FROM EVENTS WHERE EventDate = #" & CStr(intMonth) & "/" & CStr(intLoopDay) & "/" & CStr(intYear) & "# and ApproveDate is not null AND (EventPublishDate < date() or EventPublishDate is Null)")
                    If Not rsEvents.EOF Then
                        If BigCal = True Then
                            strTmpCell = strTmpCell & ">" & "<b>" & strCell & "</b>" & "</a>"
                            numevents = 1
                            Do Until rsEvents.EOF Or numevents = 5
                                strTmpCell = strTmpCell & "<br><font size=-2>" & rsEvents("EventName") & "</font>"
                                numevents = numevents + 1
                                rsEvents.MoveNext
                            Loop
                        Else

                            strTmpCell = strTmpCell & ">" & "<b>" & strCell & "</b>" & "</a>"
                        End If
                    Else
                        strTmpCell = strTmpCell & ">" & strCell & "</a>"
                       If BigCal = True Then
                          strTmpCell = strTmpCell '& "<br>&nbsp;<br>&nbsp;<br>&nbsp;"
                        End If
                    End If
                    rsEvents.Close
          
                End If
                        
                If BigCal = True Then
                    strCal = strCal & PrintCell(strTmpCell, " width=" & mvarCellWidth & " height=" & mvarCellHeight & "  valign=top ","")
                Else
                    strCal = strCal & PrintCell(strTmpCell,"","")
                End If
            End If
            
            intLoopDay = intLoopDay + 1
            If (intLoopMonth <> intMonth) And intLoopDay > intLastDayLastMonth Then
                intLoopDay = 1
                intLoopMonth = intMonth
            End If
            
            If (intLoopMonth = intMonth) And intLoopDay > intLastDay Then
                intLoopDay = 1
                intLoopMonth = intLoopMonth + 1
            End If
            
        Next
        
        strCal = strCal & PrintRowEnd
        If intLoopMonth > intMonth Then Exit For
    Next
    Set rsEvents = Nothing
    Conn.Close
    Set Conn = Nothing
    strCal = strCal & PrintTableEnd
    
    DrawCalendar = strCal
    
End Function

End Class

</script>
