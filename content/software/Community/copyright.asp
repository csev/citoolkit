<%
	set Connx = server.CreateObject("adodb.connection")
	Connx.Open Session("Connection1_ConnectionString")
	
	Set rsc = Connx.Execute("SELECT * FROM SiteInfo")
	
	If rsc("OrgIncludeCopyright") = True then
		Set fs = server.CreateObject("Scripting.FileSystemObject")
		strPath =  rsc("OrgIncludeCopyrightFile")
		' Response.Write strPath
		if fs.FileExists(strPath) then
			Const ForReading = 1
			set ts = fs.OpenTextFile(strPath,ForReading)
			Response.Write ts.readall
			set ts=nothing
		end if
		set fs = nothing
	End If
	
	rsc.close
	Set rsc=Nothing
	Connx.Close
	Set Connx = Nothing

%>