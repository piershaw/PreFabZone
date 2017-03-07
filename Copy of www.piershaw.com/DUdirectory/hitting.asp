<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/connDUdirectory.asp" -->
<%

if(Request.QueryString("id") <> "") then cmdHit__id = Request.QueryString("id")

%>
<%

set cmdHit = Server.CreateObject("ADODB.Command")
cmdHit.ActiveConnection = MM_connDUdirectory_STRING
cmdHit.CommandText = "UPDATE LINKS  SET NO_HITS = NO_HITS + 1  WHERE LINK_ID = " + Replace(cmdHit__id, "'", "''") + ""
cmdHit.CommandType = 1
cmdHit.CommandTimeout = 0
cmdHit.Prepared = true
cmdHit.Execute()
%>
<html>
<head>
<title>DUdirectory</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<% Response.Redirect (Request.QueryString("url")) %>
</body>
</html>
