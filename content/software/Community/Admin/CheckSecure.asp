<%

If (Session("userID")& "") = ""  then
	Response.Redirect "login.asp?mode=1"
End If

%>
