<!--#include file="../Connections/connDUportal.asp" -->
<%
'****************************************************************************************
'**  Copyright Notice                                                               
'**  Copyright 2003 DUware All Rights Reserved.                                
'**  This program is free software; you can modify (at your own risk) any part of it 
'**  under the terms of the License that accompanies this software and use it both 
'**  privately and commercially.
'**  All copyright notices must remain in tacked in the scripts and the 
'**  outputted HTML.
'**  You may use parts of this program in your own private work, but you may NOT
'**  redistribute, repackage, or sell the whole or any part of this program even 
'**  if it is modified or reverse engineered in whole or in part without express 
'**  permission from the author.
'**  You may not pass the whole or any part of this application off as your own work.
'**  All links to DUware and powered by logo's must remain unchanged and in place
'**  and must remain visible when the pages are viewed unless permission is first granted
'**  by the copyright holder.
'**  This program is distributed in the hope that it will be useful,
'**  but WITHOUT ANY WARRANTY; without even the implied warranty of
'**  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE OR ANY OTHER 
'**  WARRANTIES WHETHER EXPRESSED OR IMPLIED.
'**  No official support is available for this program but you may post support questions at: -
'**  http://www.duware.com/support
'****************************************************************************************
%>


<% If Request.Cookies(Request.QueryString("DAT_PARENT")) = "" Then %>


<%
set cmdUpdateAns = Server.CreateObject("ADODB.Command")
cmdUpdateAns.ActiveConnection = MM_connDUportal_STRING
cmdUpdateAns.CommandText = "UPDATE DATAS  SET DAT_COUNT = DAT_COUNT +1 WHERE DAT_ID = " & Request.QueryString("DAT_ID")
cmdUpdateAns.CommandType = 1
cmdUpdateAns.CommandTimeout = 0
cmdUpdateAns.Prepared = true
cmdUpdateAns.Execute()

set cmdUpdateQuest = Server.CreateObject("ADODB.Command")
cmdUpdateQuest.ActiveConnection = MM_connDUportal_STRING
cmdUpdateQuest.CommandText = "UPDATE DATAS  SET DAT_COUNT = = DAT_COUNT + 1, DAT_LAST = date() WHERE DAT_ID = " & Request.QueryString("DAT_PARENT")
cmdUpdateQuest.CommandType = 1
cmdUpdateQuest.CommandTimeout = 0
cmdUpdateQuest.Prepared = true
cmdUpdateQuest.Execute()

Response.Cookies(Request.QueryString("DAT_PARENT")) = Request.QueryString("DAT_ID")
Response.Cookies(Request.QueryString("DAT_PARENT")).Expires = Date + 1000

Response.Redirect("../home/poll.asp?action=vote") & "&iData=" & Request.QueryString("DAT_PARENT") & "&iCat=" & Request.QueryString("DAT_PTYPE") & "&iChannel=" & Request.QueryString("CHA_ID") & "&nChannel=" &  Request.QueryString("CHA_NAME")
%>

<% 
Else
Response.Redirect("../home/poll.asp?action=voted") & "&iData=" & Request.QueryString("DAT_PARENT") & "&iCat=" & Request.QueryString("DAT_PTYPE") & "&iChannel=" & Request.QueryString("CHA_ID") & "&nChannel=" &  Request.QueryString("CHA_NAME")
%>

<% End If %>

