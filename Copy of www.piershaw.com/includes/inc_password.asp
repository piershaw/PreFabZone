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
<% If Request.Form("Register") <> "" Then %>
<%
Dim rsPassword
Dim rsPassword_numRows

Set rsPassword = Server.CreateObject("ADODB.Recordset")
rsPassword.ActiveConnection = MM_connDUportal_STRING
rsPassword.Source = "SELECT * FROM USERS WHERE U_EMAIL = '" & Request.Form("email") & "'"
rsPassword.CursorType = 0
rsPassword.CursorLocation = 2
rsPassword.LockType = 1
rsPassword.Open()

rsPassword_numRows = 0
If rsPassword.EOF Or rsPassword.BOF Then 
response.redirect "../home/password.asp?result=Your account was not found! Please check your entry."

Else

Do While NOT rsPassword.EOF

	Dim objCDO
	Set objCDO = Server.CreateObject("CDONTS.NewMail")
	objCDO.From = strEmail
	objCDO.To = rsPassword.Fields.Item("U_EMAIL").Value
	objCDO.Subject = "Your password"
	
	objCDO.Body = "Thank you for using " & strPageTitle & ". Below is your account ID and Password. " & vbnewline & "UserName: " & rsPassword.Fields.Item("U_ID").Value & vbnewline & "Password: " & rsPassword.Fields.Item("U_PASSWORD").Value & vbnewline
	
	objCDO.Send()
	Set objCDO = Nothing

rsPassword.MoveNext
loop

response.redirect "../home/password.asp?result=Your password has been sent to " & Request.Form("email")

End If

rsPassword.Close()
%>
<% End If %>


<link href="../assets/DUportal.css" rel="stylesheet" type="text/css"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#003399">
              <tr> 
                <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../assets/bg_header.gif">
                    <tr> 
                      <td width="10"><img src="../assets/header_end_left.gif"></td>
                      <td align="left" valign="middle" class="textBoldColor">RETRIEVE 
                        LOST PASSWORD</td>
                      <td width="28" align="right" valign="middle"><img src="../assets/header_end_right.gif"></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif"></td>
				<form name="register" action ="" method="post">
                  <td align="left" valign="top" class="bgTable"> <table align="center" cellpadding="3" cellspacing="3">
                      <tr valign="baseline">
                        <td colspan="2" align="left" valign="middle" nowrap class="textRed">
						<%= Request.QueryString("result") %>
						</td>
                      </tr>
                      <tr valign="baseline"> 
                        <td colspan="2" align="left" valign="middle" nowrap class="textBold">Please 
                          provide the email address you used to register:</td>
                      </tr>
                      <tr valign="baseline"> 
                        <td align="right" nowrap class="textBold">EMAIL ADDRESS:</td>
                        <td> <input name="EMAIL" type="text" class="form" id="EMAIL" value="" size="40" maxlength="60"> 
                        </td>
                      </tr>
                      <tr valign="baseline"> 
                        <td nowrap align="right">&nbsp;</td>
                        <td> <input name="Register" type="submit" class="button" id="Register" onClick="MM_validateForm('U_EMAIL','','RisEmail');return document.MM_returnValue" value="Submit"> 
                        </td>
                      </tr>
                    </table>
                    
                  </td>
                </form>
                <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif"></td>
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
