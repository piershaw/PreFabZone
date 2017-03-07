
<!--#include file="../Connections/connDUportal.asp" -->
<%

if(Request.QueryString("ID") <> "") then cmdActivate__id = Request.QueryString("ID")

%>
<%

set cmdDeactivate = Server.CreateObject("ADODB.Command")
cmdDeactivate.ActiveConnection = MM_connDUportal_STRING
cmdDeactivate.CommandText = "UPDATE QUESTIONS  SET QUEST_ACTIVE = False  WHERE QUEST_ACTIVE = True"
cmdDeactivate.CommandType = 1
cmdDeactivate.CommandTimeout = 0
cmdDeactivate.Prepared = true
cmdDeactivate.Execute()

%>
<%

set cmdActivate = Server.CreateObject("ADODB.Command")
cmdActivate.ActiveConnection = MM_connDUportal_STRING
cmdActivate.CommandText = "UPDATE QUESTIONS  SET QUEST_ACTIVE  = True WHERE QUEST_ID = " + Replace(cmdActivate__id, "'", "''") + ""
cmdActivate.CommandType = 1
cmdActivate.CommandTimeout = 0
cmdActivate.Prepared = true
cmdActivate.Execute()

%>
<%
Response.Redirect "polls.asp"
%>