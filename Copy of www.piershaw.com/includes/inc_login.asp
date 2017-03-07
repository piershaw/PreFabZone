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
Dim rsLogged
Dim rsLogged_numRows

Set rsLogged = Server.CreateObject("ADODB.Recordset")
rsLogged.ActiveConnection = MM_connDUportal_STRING
rsLogged.Source = "SELECT * FROM USERS WHERE U_ID = '" & Session("MM_Username") & "' OR U_ID = '" & Request.Cookies("DUportalUser") & "'"
rsLogged.CursorType = 0
rsLogged.CursorLocation = 2
rsLogged.LockType = 1
rsLogged.Open()

rsLogged_numRows = 0
%>
 <link href="../assets/DUportal.css" rel="stylesheet" type="text/css"> 
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#003399">
              <tr> 
                <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../assets/bg_header.gif">
                    <tr> 
                      <td width="10"><img src="../assets/header_end_left.gif"></td>
                      <td align="left" valign="middle" class="textBoldColor"> 
                        <% If Not rsLogged.EOF Or Not rsLogged.BOF Then %>
                        USER 
                        <% End If ' end Not rsLogged.EOF Or NOT rsLogged.BOF %> <% If rsLogged.EOF And rsLogged.BOF Then %>
                        LOGIN 
                        <% End If ' end rsLogged.EOF And rsLogged.BOF %> </td>
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
                <form name="login" method="post" action="../includes/inc_logging.asp">
                  <td align="left" valign="top" class="bgTable"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td align="left" valign="top"> <% If rsLogged.EOF And rsLogged.BOF Then %>
                          <table width="100%" border="0" cellpadding="2" cellspacing="2" class="bgTable">
                            <tr align="left" valign="middle"> 
                              <td class="textBold">User ID:</td>
                              <td><input name="id" type="text" class="form" id="id" size="15" maxlength="100"></td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td class="textBold">Password: </td>
                              <td><input name="password" type="password" class="form" id="password" size="15" maxlength="20"></td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td class="textBold">&nbsp; </td>
                              <td><input name="Submit" type="submit" class="button" onClick="MM_validateForm('id','','R','password','','R');return document.MM_returnValue" value="Login"></td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td colspan="2" class="textColor"> &#8226; <a href="../home/register.asp">New 
                                User Registration</a></td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td colspan="2" class="textColor"> &#8226; <a href="../home/password.asp">Retrieve 
                                Lost Password</a></td>
                            </tr>
                          </table>
                          <% End If ' end rsLogged.EOF And rsLogged.BOF %> </td>
                      </tr>
                      <tr>
                        <td align="left" valign="top"> <% If Not rsLogged.EOF Or Not rsLogged.BOF Then %>
                          <table width="100%" border="0" cellspacing="2" cellpadding="2">
                            <tr> 
                              <td align="left" valign="middle" class="textBoldColor">Welcome 
                                back, <%=(rsLogged.Fields.Item("U_FIRST").Value)%>!</td>
                            </tr>
                            <tr> 
                              <td align="left" valign="middle" class="textColor">&#8226; 
                                <a href="../home/profile.asp">My Profile</a></td>
                            </tr>
                            <tr> 
                              <td align="left" valign="middle" class="textColor">&#8226; 
                                <a href="../home/portal.asp">My Portal</a></td>
                            </tr>
                           
                            <tr> 
                              <td align="left" valign="middle" class="textColor">&#8226; 
                                <a href="../includes/inc_logout.asp">Log Out</a></td>
                            </tr>
                          </table>
                          <% End If ' end Not rsLogged.EOF Or NOT rsLogged.BOF %> </td>
                      </tr>
                    </table> </td>
                </form>
                <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif"></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td align="center" valign="top" background="../assets/bg_header_bottom.gif"><table border="0" cellpadding="0" cellspacing="0" class="bgTable" >
              <tr> 
                <td><img src="../assets/header_bottom.gif"></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="5" align="left" valign="top"><img src="../assets/_spacer.gif" width="1" height="1"></td>
  </tr>
</table>

<%
rsLogged.Close()
Set rsLogged = Nothing
%>
