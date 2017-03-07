<%@LANGUAGE="VBSCRIPT"%>

<!--#include file="../Connections/connDUportal.asp" -->
<%

if(Request.QueryString("rate_value") <> "") then cmdRating__varRATE = Request.QueryString("rate_value")

if(Request.QueryString("id") <> "") then cmdRating__varID = Request.QueryString("id")

if(Request.QueryString("catid") <> "") then cmdRating__catID = Request.QueryString("catid")

%>
<%
set cmdRating = Server.CreateObject("ADODB.Command")
cmdRating.ActiveConnection = MM_connDUportal_STRING
cmdRating.CommandText = "UPDATE LINKS  SET LINK_RATE = LINK_RATE + " + Replace(cmdRating__varRATE, "'", "''") + ", NO_RATES = NO_RATES + 1  WHERE LINK_ID = " + Replace(cmdRating__varID, "'", "''") + ""
cmdRating.CommandType = 1
cmdRating.CommandTimeout = 0
cmdRating.Prepared = true
cmdRating.Execute()
%>


<%
Response.Redirect "dirCat.asp?id=" & cmdRating__catid
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
</body>
</html>
