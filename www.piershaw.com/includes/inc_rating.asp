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
<%

if(Request.QueryString("iRate") <> "") then cmdRating__iRate = Request.QueryString("iRate")

if(Request.QueryString("iData") <> "") then cmdRating__iData = Request.QueryString("iData")

if(Request.QueryString("iCat") <> "") then cmdRating__iCat = Request.QueryString("iCat")

if(Request.QueryString("iChannel") <> "") then cmdRating__iChannel = Request.QueryString("iChannel")

if(Request.QueryString("nChannel") <> "") then cmdRating__nChannel = Request.QueryString("nChannel")

%>
<%
set cmdRating = Server.CreateObject("ADODB.Command")
cmdRating.ActiveConnection = MM_connDUportal_STRING
cmdRating.CommandText = "UPDATE DATAS  SET DAT_RATED = DAT_RATED + " + Replace(cmdRating__iRate, "'", "''") + ", DAT_RATES = DAT_RATES + 1  WHERE DAT_ID = " + Replace(cmdRating__iData, "'", "''") + ""
cmdRating.CommandType = 1
cmdRating.CommandTimeout = 0
cmdRating.Prepared = true
cmdRating.Execute()
%>
<%
Response.Cookies(cmdRating__iData) = cmdRating__iRate
Response.Cookies(cmdRating__iData).Expires = Date + 1000
Response.Redirect "../home/detail.asp?iData=" & cmdRating__iData & "&iCat=" & cmdRating__iCat & "&iChannel=" & cmdRating__iChannel & "&nChannel=" & cmdRating__nChannel
%>



