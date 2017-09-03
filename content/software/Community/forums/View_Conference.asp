<!-- #include File="../mapvar.asp" -->
<html>

<head>
<title>Library Conferences</title>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
</head>

<frameset rows="65,40%,*">
  <frame name="header" scrolling="no" target="main" src="Title_Conference.asp?ConferenceTitle=<%= Request.QueryString("ConferenceTitle")%>&ConferenceID=<%= Request.QueryString("ConferenceID")%>">
  <frame name="listmsg" scrolling="auto" src="Browse_Conference_Messages.asp?ConferenceID=<%= Request.QueryString("ConferenceID")%>&ConferenceModerated=<%= Request.QueryString("ConferenceModerated")%>">
  <frame name="message" scrolling="auto" src="Default_Display_Message.asp?ConferenceTitle=<%= Request.QueryString("ConferenceTitle")%>&ConferenceID=<%= Request.QueryString("ConferenceID")%>">
  <noframes>
  <body>
  <p>This page uses frames, but your browser doesn't support them.</p>
  </body>
  </noframes>
</frameset>
</html>
