
<!--#include file="../Connections/connDUportal.asp" -->

<% If Request.QueryString("SUBMIT") = "POST" then 'submit = post ----> codes for posting a new message %>
<%
if(Request.QueryString("FOR_ID") <> "") then cmdPost__for_id = Request.QueryString("FOR_ID")

if(Request.QueryString("MSG_AUTHOR") <> "") then cmdPost__msg_author = Request.QueryString("MSG_AUTHOR")

if(Request.QueryString("MSG_SUBJECT") <> "") then cmdPost__msg_subject = Request.QueryString("MSG_SUBJECT")

if(Request.QueryString("MSG_BODY") <> "") then cmdPost__msg_body = Request.QueryString("MSG_BODY")
%>
<% 'codes to insert a new topic into MESSAGES table
set cmdPost = Server.CreateObject("ADODB.Command")
cmdPost.ActiveConnection = MM_connDUportal_STRING
cmdPost.CommandText = "INSERT INTO MESSAGES (FOR_ID, MSG_AUTHOR, MSG_SUBJECT, MSG_BODY) VALUES (" + Replace(cmdPost__for_id, "'", "''") + ", '" + Replace(cmdPost__msg_author, "'", "''") + "',  '" + Replace(cmdPost__msg_subject, "'", "''") + "', '" + Replace(cmdPost__msg_body, "'", "''") + "') "
cmdPost.CommandType = 1
cmdPost.CommandTimeout = 0
cmdPost.Prepared = true
cmdPost.Execute()
%>

<% ' codes to insert LAST_POST and increase 1 to TOPIC_COUNT  in the FORUMS table
set cmdPostCount = Server.CreateObject("ADODB.Command")
cmdPostCount.ActiveConnection = MM_connDUportal_STRING
cmdPostCount.CommandText = "UPDATE FORUMS  SET FOR_TOPIC_COUNT = FOR_TOPIC_COUNT + 1, FOR_LAST_POST = NOW() WHERE FOR_ID = " + Replace(cmdPost__for_id, "'", "''") + " "
cmdPostCount.CommandType = 1
cmdPostCount.CommandTimeout = 0
cmdPostCount.Prepared = true
cmdPostCount.Execute()
%>
<% Response.Redirect "messages.asp?for_id=" & cmdPost__for_id %>

<% End If 'submit = post ----> codes for posting a new message %>

















<% If Request.QueryString("SUBMIT") = "REPLY" then 'submit = post ----> codes for posting a new message %>
<%
if(Request.QueryString("REP_AUTHOR") <> "") then cmdReply__rep_author = Request.QueryString("REP_AUTHOR")

if(Request.QueryString("REP_BODY") <> "") then cmdReply__rep_body = Request.QueryString("REP_BODY")

if(Request.QueryString("MSG_ID") <> "") then cmdReply__msg_id = Request.QueryString("MSG_ID")

if(Request.QueryString("FOR_ID") <> "") then cmdReply__for_id = Request.QueryString("FOR_ID")
%>
<% 'codes to insert a new reply into REPLIES table
set cmdReply = Server.CreateObject("ADODB.Command")
cmdReply.ActiveConnection = MM_connDUportal_STRING
cmdReply.CommandText = "INSERT INTO REPLIES (MSG_ID, REP_AUTHOR, REP_BODY) VALUES (" + Replace(cmdReply__msg_id, "'", "''") + ", '" + Replace(cmdReply__rep_author, "'", "''") + "',  '" + Replace(cmdReply__rep_body, "'", "''") + "') "
cmdReply.CommandType = 1
cmdReply.CommandTimeout = 0
cmdReply.Prepared = true
cmdReply.Execute()
%>

<% ' codes to insert FOR_LAST_POST and increase 1 to FOR_REP_COUNT  in the FORUMS table
set cmdForRepCount = Server.CreateObject("ADODB.Command")
cmdForRepCount.ActiveConnection = MM_connDUportal_STRING
cmdForRepCount.CommandText = "UPDATE FORUMS  SET FOR_REPLY_COUNT = FOR_REPLY_COUNT + 1, FOR_LAST_POST = NOW() WHERE FOR_ID = " + Replace(cmdReply__for_id, "'", "''") + " "
cmdForRepCount.CommandType = 1
cmdForRepCount.CommandTimeout = 0
cmdForRepCount.Prepared = true
cmdForRepCount.Execute()
%>

<% ' codes to insert MSG_LAST_POST and increase 1 to MSG_REPLY_COUNT  in the FORUMS table
set cmdMsgRepCount = Server.CreateObject("ADODB.Command")
cmdMsgRepCount.ActiveConnection = MM_connDUportal_STRING
cmdMsgRepCount.CommandText = "UPDATE MESSAGES SET MSG_REPLY_COUNT = MSG_REPLY_COUNT + 1, MSG_LAST_POST = NOW() WHERE MSG_ID = " + Replace(cmdReply__msg_id, "'", "''") + " "
cmdMsgRepCount.CommandType = 1
cmdMsgRepCount.CommandTimeout = 0
cmdMsgRepCount.Prepared = true
cmdMsgRepCount.Execute()
%>


<% Response.Redirect "msgDetail.asp?msg_id=" & cmdReply__msg_id & "&for_id=" & cmdReply__for_id %>

<% End If 'submit = post ----> codes for posting a new reply %>
