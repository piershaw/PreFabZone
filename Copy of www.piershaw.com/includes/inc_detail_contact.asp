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

<% If Request.Form("Sent") <> "" Then
Dim objCDO
Set objCDO = Server.CreateObject("CDONTS.NewMail")
objCDO.From = Request.Form("EMAIL")
objCDO.To = Request.Form("POSTER_EMAIL")
objCDO.cc = Request.Form("EMAIL")
objCDO.Subject = Request.Form("SUBJECT")
objCDO.Body = "Hello, " & Request.Form("POSTER_NAME") & vbnewline & "!" & vbnewline & vbnewline & "Your classified ad posted at " & strPageTitle & " has been replied. Below is the detail message." & vbnewline & vbnewline & "From: " & Request.Form("NAME") & vbnewline & "Email Address: " & Request.Form("EMAIL")  & vbnewline & "Dated: " & now() & vbnewline & "Subject: " & Request.Form("SUBJECT") & vbnewline & "Message: " & Request.Form("MESSAGE")
objCDO.Send()
Set objCDO = Nothing
End If
%>

<%
Dim rsContact__MMColParam
rsContact__MMColParam = "0"
if (Request.QueryString("iData") <> "") then rsContact__MMColParam = Request.QueryString("iData")
%>
<%
set rsContact = Server.CreateObject("ADODB.Recordset")
rsContact.ActiveConnection = MM_connDUportal_STRING
rsContact.Source = "SELECT *  FROM DATAS, CATEGORIES, CHANNELS, USERS  WHERE DAT_ID = " + Replace(rsContact__MMColParam, "'", "''") + " AND DAT_CATEGORY = CAT_ID AND CAT_CHANNEL = CHA_ID AND DAT_USER = U_ID"
rsContact.CursorType = 0
rsContact.CursorLocation = 2
rsContact.LockType = 3
rsContact.Open()
rsContact_numRows = 0
%>
<%
Dim rsReplier__MMColParam
rsReplier__MMColParam = "1"
If (Request.Cookies("DUportalUser") <> "") Then 
  rsReplier__MMColParam = Request.Cookies("DUportalUser")
End If
%>
<%
Dim rsReplier
Dim rsReplier_numRows

Set rsReplier = Server.CreateObject("ADODB.Recordset")
rsReplier.ActiveConnection = MM_connDUportal_STRING
rsReplier.Source = "SELECT * FROM USERS WHERE U_ID = '" + Replace(rsReplier__MMColParam, "'", "''") + "'"
rsReplier.CursorType = 0
rsReplier.CursorLocation = 2
rsReplier.LockType = 1
rsReplier.Open()

rsReplier_numRows = 0
%>
 <link href="../assets/DUportal.css" rel="stylesheet" type="text/css">
<% If Not rsContact.EOF Or Not rsContact.BOF Then %>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#003399">
              <tr> 
                <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../assets/bg_header.gif">
                    <tr> 
                      <td width="10"><img src="../assets/header_end_left.gif"></td>
                      <td align="left" valign="middle" class="textBoldColor"><a href="default.asp">HOME 
                        </a> &raquo; <%=UCASE(rsContact.Fields.Item("CHA_NAME").Value)%> &raquo; <%=UCASE(rsContact.Fields.Item("CAT_NAME").Value)%> &raquo; CONTACT</td>
                      <td width="28" align="right" valign="middle"><img src="../assets/header_end_right.gif"></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> <form action="" method="post" name="CONTACT" id="CONTACT">
                <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif"></td>
                <td align="center" valign="top" class="bgTable">
                    <table border="0" cellpadding="2" cellspacing="2" class="textBold">
                      <% If Not rsReplier.EOF Or Not rsReplier.BOF Then %>
                      <tr> 
                        <td align="left" valign="middle">Name: 
                        </td>
                        <td><input name="NAME" type="text" class="form" id="NAME" value="<%=(rsReplier.Fields.Item("U_FIRST").Value)%>&nbsp;<%=(rsReplier.Fields.Item("U_LAST").Value)%>" size="50" maxlength="50"></td>
                      </tr>
                      <tr> 
                        <td align="left" valign="middle">Email 
                          Address: </td>
                        <td><input name="EMAIL" type="text" class="form" id="EMAIL" value="<%=(rsReplier.Fields.Item("U_EMAIL").Value)%>" size="50" maxlength="100"></td>
                      </tr>
                      <% End If ' end Not rsReplier.EOF Or NOT rsReplier.BOF %>
                      <% If rsReplier.EOF And rsReplier.BOF Then %>
                      <tr> 
                        <td align="left" valign="middle">Name: 
                        </td>
                        <td><input name="NAME" type="text" class="form" id="NAME" value="" size="50" maxlength="60"></td>
                      </tr>
                      <tr> 
                        <td align="left" valign="middle">Email 
                          Address: </td>
                        <td><input name="EMAIL" type="text" class="form" id="EMAIL" value="" size="50" maxlength="100"></td>
                      </tr>
                      <% End If ' end rsReplier.EOF And rsReplier.BOF %>
                      <tr> 
                        <td align="left" valign="middle">Subject: </td>
                        <td><input name="SUBJECT" type="text" class="form" id="SUBJECT" value="Re: <%=(rsContact.Fields.Item("DAT_NAME").Value)%>" size="50" maxlength="150"></td>
                      </tr>
                      <tr> 
                        <td align="left" valign="top">Message: </td>
                        <td><textarea name="MESSAGE" cols="50" rows="20" class="form" id="MESSAGE"></textarea></td>
                      </tr>
                      <tr> 
                        <td><input name="POSTER_EMAIL" type="hidden" id="POSTER_EMAIL" value="<%=(rsContact.Fields.Item("U_EMAIL").Value)%>"> 
                          <input name="POSTER_NAME" type="hidden" id="POSTER_NAME" value="<%=(rsContact.Fields.Item("U_FIRST").Value)%>&nbsp;<%=(rsContact.Fields.Item("U_FIRST").Value)%>"></td>
                        <td><input name="Send" type="submit" class="button" id="Send" onClick="MM_validateForm('0','','R','0','','RisEmail','1','','R','1','','NisEmail','SUBJECT','','R','MESSAGE','','R');return document.MM_returnValue" value="Send"> 
                        </td>
                      </tr>
                    </table>
                 </td>
                <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif"></td>
				 </form>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td align="center" valign="top" background="../assets/bg_header_bottom.gif"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="7" align="left" valign="top"><img src="../assets/_spacer.gif" width="1" height="1"></td>
  </tr>
</table>
<% End If ' end Not rsContact.EOF Or NOT rsContact.BOF %>
<%
rsContact.Close()
Set rsContact = Nothing
%>
<%
rsReplier.Close()
Set rsReplier = Nothing
%>
