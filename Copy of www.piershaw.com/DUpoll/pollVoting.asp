<%@LANGUAGE="VBSCRIPT"%>

<!--#include file="../Connections/connDUportal.asp" -->
<%

if(Request.QueryString("ANS_ID") <> "") then cmdUpdateAns__varANS = Request.QueryString("ANS_ID")

%>
<%

if(Request.QueryString("QUEST_ID") <> "") then cmdUpdateQuest__varQUEST = Request.QueryString("QUEST_ID")

%>
<% If Request.Cookies(cmdUpdateQuest__varQUEST) <> "" Then %>
<%
Response.Redirect "pollHaveVoted.asp?id=" & cmdUpdateQuest__varQUEST
%>
<% Else %>
<%
set cmdUpdateAns = Server.CreateObject("ADODB.Command")
cmdUpdateAns.ActiveConnection = MM_connDUportal_STRING
cmdUpdateAns.CommandText = "UPDATE ANSWERS  SET VOTES = VOTES +1, LAST_VOTE = date() WHERE ANS_ID = " + Replace(cmdUpdateAns__varANS, "'", "''") + " "
cmdUpdateAns.CommandType = 1
cmdUpdateAns.CommandTimeout = 0
cmdUpdateAns.Prepared = true
cmdUpdateAns.Execute()
%>

<%
set cmdUpdateQuest = Server.CreateObject("ADODB.Command")
cmdUpdateQuest.ActiveConnection = MM_connDUportal_STRING
cmdUpdateQuest.CommandText = "UPDATE QUESTIONS  SET TOTAL_VOTES = TOTAL_VOTES +1  WHERE QUEST_ID = " + Replace(cmdUpdateQuest__varQUEST, "'", "''") + " "
cmdUpdateQuest.CommandType = 1
cmdUpdateQuest.CommandTimeout = 0
cmdUpdateQuest.Prepared = true
cmdUpdateQuest.Execute()
%>

<%
Response.Cookies(cmdUpdateQuest__varQUEST) = (cmdUpdateAns__varANS)
Response.Cookies(cmdUpdateQuest__varQUEST).Expires = Date + 1000
%>

<%
Response.Redirect "pollResult.asp?id=" & cmdUpdateQuest__varQUEST
%>

<% End If %>