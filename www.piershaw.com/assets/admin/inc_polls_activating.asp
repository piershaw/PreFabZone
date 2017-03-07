<!--#include file="../Connections/connDUportal.asp" -->
<!--#include file="inc_restriction.asp" -->

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
<%

if(Request.QueryString("ID") <> "") then cmdActivate__id = Request.QueryString("ID")

%>
<%

set cmdDeactivate = Server.CreateObject("ADODB.Command")
cmdDeactivate.ActiveConnection = MM_connDUportal_STRING
cmdDeactivate.CommandText = "UPDATE DATAS  SET DAT_ACTIVE = 0  WHERE DAT_ACTIVE = 1"
cmdDeactivate.CommandType = 1
cmdDeactivate.CommandTimeout = 0
cmdDeactivate.Prepared = true
cmdDeactivate.Execute()

%>
<%

set cmdActivate = Server.CreateObject("ADODB.Command")
cmdActivate.ActiveConnection = MM_connDUportal_STRING
cmdActivate.CommandText = "UPDATE DATAS  SET DAT_ACTIVE  = 1 WHERE DAT_ID = " + Replace(cmdActivate__id, "'", "''") + ""
cmdActivate.CommandType = 1
cmdActivate.CommandTimeout = 0
cmdActivate.Prepared = true
cmdActivate.Execute()

%>
<%
Response.Redirect "polls.asp"
%>