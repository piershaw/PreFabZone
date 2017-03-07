<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/connDUportal.asp" -->

<% If Request.QueryString("Submit") = "Delete" Then %>
<%
if(Request.QueryString("LINK_ID") <> "") then spLinkDeleting__varID = Request.QueryString("LINK_ID")
%>
<%

set spLinkDeleting = Server.CreateObject("ADODB.Command")
spLinkDeleting.ActiveConnection = MM_connDUportal_STRING
spLinkDeleting.CommandText = "DELETE FROM LINKS  WHERE LINK_ID IN (" + Replace(spLinkDeleting__varID, "'", "''") + ")"
spLinkDeleting.CommandType = 1
spLinkDeleting.CommandTimeout = 0
spLinkDeleting.Prepared = true
spLinkDeleting.Execute()

Response.Redirect("whatsnew.asp")

%>
<% End If %>



<% If Request.QueryString("Submit") = "Approve" Then %>
<%

if(Request.QueryString("LINK_ID") <> "") then spLinkApproving__varID = Request.QueryString("LINK_ID")

%>
<%

set spLinkApproving = Server.CreateObject("ADODB.Command")
spLinkApproving.ActiveConnection = MM_connDUportal_STRING
spLinkApproving.CommandText = "UPDATE LINKS  SET LINK_APPROVED = True WHERE LINK_ID IN (" + Replace(spLinkApproving__varID, "'", "''") + ") "
spLinkApproving.CommandType = 1
spLinkApproving.CommandTimeout = 0
spLinkApproving.Prepared = true
spLinkApproving.Execute()
Response.Redirect("whatsnew.asp")
%>
<% End If %>

