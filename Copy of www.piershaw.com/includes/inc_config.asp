<!--#include file="../Connections/connDUportal.asp" -->
<% Response.Buffer = True %>
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
Dim rsConfig
Dim rsConfig_numRows

Set rsConfig = Server.CreateObject("ADODB.Recordset")
rsConfig.ActiveConnection = MM_connDUportal_STRING
rsConfig.Source = "SELECT * FROM CONFIGURATION"
rsConfig.CursorType = 0
rsConfig.CursorLocation = 2
rsConfig.LockType = 1
rsConfig.Open()

rsConfig_numRows = 0
%>
<%
Dim strPageTitle
Dim strPageSize
strEmail = rsConfig.Fields.Item("CON_ADMIN_EMAIL").Value
strPageTitle = rsConfig.Fields.Item("CON_TITLE").Value
strPageSize = rsConfig.Fields.Item("CON_PAGE_SIZE").Value
strLeftSize = rsConfig.Fields.Item("CON_LEFT_SIZE").Value
strRightSize = rsConfig.Fields.Item("CON_RIGHT_SIZE").Value
%>

<% 
myPaypalID = rsConfig.Fields.Item("CON_PAYPAL_ID").Value
myPaypalCurrency = rsConfig.Fields.Item("CON_PAYPAL_CURRENCY").Value
myPaypalCurrencySign = rsConfig.Fields.Item("CON_PAYPAL_CURRENCY_SIGN").Value
myReturnURL = rsConfig.Fields.Item("CON_PAYPAL_RETURN_SUCCESS").Value
myCancelURL = rsConfig.Fields.Item("CON_PAYPAL_RETURN_CANCEL").Value
%>


<%
Function TrimBody(str)

  	Str = (Replace(str, "<", "&lt;"))
  	Str = (Replace(str, ">", "&gt;")) 
  	Str = (Replace(str, vbCrlf, "<br>"))
	Str = (Replace(str, "[", "<"))
	Str = (Replace(str, "]", ">"))
	
	dim re, sOut
	set re = New RegExp
	re.global = true
	re.ignorecase = true
	re.pattern = "((mailto\:|(news|(ht|f)tp(s?))\://){1}\S+)"
	sOut = re.replace( Str, "<A HREF=""$1"" TARGET=""_new"">$1</A>")
	set re = Nothing
  	
   TrimBody = sOut
   
End Function
%>

<%
rsConfig.Close()
Set rsConfig = Nothing
%>
