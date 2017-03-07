<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/connDUportal.asp" -->





<% If Request.QueryString("Submit") = "Delete" Then %>
<%
if(Request.QueryString("NEWS_ID") <> "") then spNewsDeleting__varID = Request.QueryString("NEWS_ID")
%>
<%

set spNewsDeleting = Server.CreateObject("ADODB.Command")
spNewsDeleting.ActiveConnection = MM_connDUportal_STRING
spNewsDeleting.CommandText = "DELETE FROM NEWS  WHERE NEWS_ID IN (" + Replace(spNewsDeleting__varID, "'", "''") + ")"
spNewsDeleting.CommandType = 1
spNewsDeleting.CommandTimeout = 0
spNewsDeleting.Prepared = true
spNewsDeleting.Execute()
Response.Redirect("whatsnew.asp")
%>
<% End If %>




<% If Request.QueryString("Submit") = "Approve" Then %>
<%

if(Request.QueryString("NEWS_ID") <> "") then spNewsApproving__varID = Request.QueryString("NEWS_ID")

%>
<%

set spNewsApproving = Server.CreateObject("ADODB.Command")
spNewsApproving.ActiveConnection = MM_connDUportal_STRING
spNewsApproving.CommandText = "UPDATE NEWS  SET NEWS_APPROVED = True WHERE NEWS_ID IN (" + Replace(spNewsApproving__varID, "'", "''") + ") "
spNewsApproving.CommandType = 1
spNewsApproving.CommandTimeout = 0
spNewsApproving.Prepared = true
spNewsApproving.Execute()
Response.Redirect("whatsnew.asp")
%>
<% End If %>















<% If Request.QueryString("Submit") = "Approve" Then %>
<%

if(Request.QueryString("NEWS_ID") <> "") then spNewsApproving__varID = Request.QueryString("NEWS_ID")

%>
<%

set spNewsApproving = Server.CreateObject("ADODB.Command")
spNewsApproving.ActiveConnection = MM_connDUportal_STRING
spNewsApproving.CommandText = "UPDATE NEWS  SET NEWS_APPROVED = True WHERE NEWS_ID IN (" + Replace(spNewsApproving__varID, "'", "''") + ") "
spNewsApproving.CommandType = 1
spNewsApproving.CommandTimeout = 0
spNewsApproving.Prepared = true
spNewsApproving.Execute()
Response.Redirect("whatsnew.asp")
%>
<% End If %>
