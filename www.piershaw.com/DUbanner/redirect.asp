<%@LANGUAGE="VBSCRIPT"%>
<% Response.Buffer = "true" %>
<!--#include file="../Connections/connDUportal.asp" -->
<%

if(Request.QueryString("id") <> "") then spRedirect__varID = Request.QueryString("id")

%>
<%

set spRedirect = Server.CreateObject("ADODB.Command")
spRedirect.ActiveConnection = MM_connDUportal_STRING
spRedirect.CommandText = "UPDATE BANNERS   SET B_CLICKED_TOTAL = B_CLICKED_TOTAL + 1, B_CLICKED_DATE = Date()          WHERE B_ID = " + Replace(spRedirect__varID, "'", "''") + "  "
spRedirect.CommandType = 1
spRedirect.CommandTimeout = 0
spRedirect.Prepared = true
spRedirect.Execute()
%>
<html>
<head>
<title>Redirect</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<% Response.Redirect (Request.QueryString("url")) %>
<body text="#000000">
</body>
</html>
