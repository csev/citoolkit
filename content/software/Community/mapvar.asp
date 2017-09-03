<%

	Response.Buffer = true

	if Session("Connection1_ConnectionString") = "" then 
		Response.Write "Please choose a library"
		set fs = server.CreateObject("scripting.filesystemobject")
		toplevel = fs.FileExists(server.MapPath("global.asa"))
		set fs = nothing
		if  toplevel then
			Response.redirect "start/expired.htm"
		else
			Response.Redirect "../start/expired.htm"
		end if
	End If

	Application("Connection1_ConnectionString") = Session("Connection1_ConnectionString")
	Application("Connection1_ConnectionTimeout") = Session("Connection1_ConnectionTimeout")
	Application("Connection1_CommandTimeout") = Session("Connection1_CommandTimeout")
	Application("Connection1_CursorLocation") = Session("Connection1_CursorLocation")
	Application("Connection1_RuntimeUserName") = ""
	Application("Connection1_RuntimePassword") = ""

%>