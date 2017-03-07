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
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_connDUportal_STRING
  MM_editTable = "CONFIGURATION"
  MM_editColumn = "CON_ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "config.asp"
  MM_fieldsStr  = "CON_PAGE_SIZE|value|CON_TITLE|value|CON_ADMIN_EMAIL|value|CON_LEFT_SIZE|value|CON_RIGHT_SIZE|value|CON_PAYPAL_ID|value|CON_PAYPAL_CURRENCY|value|CON_PAYPAL_CURRENCY_SIGN|value|CON_PAYPAL_RETURN_SUCCESS|value|CON_PAYPAL_RETURN_CANCEL|value"
  MM_columnsStr = "CON_PAGE_SIZE|',none,''|CON_TITLE|',none,''|CON_ADMIN_EMAIL|',none,''|CON_LEFT_SIZE|',none,''|CON_RIGHT_SIZE|',none,''|CON_PAYPAL_ID|',none,''|CON_PAYPAL_CURRENCY|',none,''|CON_PAYPAL_CURRENCY_SIGN|',none,''|CON_PAYPAL_RETURN_SUCCESS|',none,''|CON_PAYPAL_RETURN_CANCEL|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim rsSiteConfig
Dim rsSiteConfig_numRows

Set rsSiteConfig = Server.CreateObject("ADODB.Recordset")
rsSiteConfig.ActiveConnection = MM_connDUportal_STRING
rsSiteConfig.Source = "SELECT * FROM CONFIGURATION"
rsSiteConfig.CursorType = 0
rsSiteConfig.CursorLocation = 2
rsSiteConfig.LockType = 1
rsSiteConfig.Open()

rsSiteConfig_numRows = 0
%>
 <link rel="stylesheet" href="../css/default.css" type="text/css">
<script language="JavaScript">
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
      } else if (test!='R') {
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (val<min || max<val) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
<link href="../assets/DUportal.css" rel="stylesheet" type="text/css"> 
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  
  <tr>
    <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#003399">
              <tr> 
                <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../assets/bg_header.gif">
                    <tr> 
                      <td width="10"><img src="../assets/header_end_left.gif"></td>
                      <td align="left" valign="middle" class="textBoldColor">CONFIGUARION</td>
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
                <td align="left" valign="top" class="bgTable">&nbsp; <form method="post" action="<%=MM_editAction%>" name="form1">
                    <table align="center" cellpadding="3" cellspacing="3" class="textBold">
                      <tr valign="baseline"> 
                        <td nowrap align="right">PAGE SIZE:</td>
                        <td class="textRed"> <input name="CON_PAGE_SIZE" type="text" class="form" value="<%=(rsSiteConfig.Fields.Item("CON_PAGE_SIZE").Value)%>" size="10">
                          (Use 100% for stretched full page) </td>
                      </tr>
                      <tr valign="baseline"> 
                        <td nowrap align="right">SITE TITLE:</td>
                        <td> <input name="CON_TITLE" type="text" class="form" value="<%=(rsSiteConfig.Fields.Item("CON_TITLE").Value)%>" size="32">
                        </td>
                      </tr>
                      <tr valign="baseline"> 
                        <td nowrap align="right">ADMIN EMAIL:</td>
                        <td> <input name="CON_ADMIN_EMAIL" type="text" class="form" value="<%=(rsSiteConfig.Fields.Item("CON_ADMIN_EMAIL").Value)%>" size="50">
                        </td>
                      </tr>
                      <tr valign="baseline"> 
                        <td nowrap align="right">LEFT SIZE:</td>
                        <td> <input name="CON_LEFT_SIZE" type="text" class="form" value="<%=(rsSiteConfig.Fields.Item("CON_LEFT_SIZE").Value)%>" size="10">
                        </td>
                      </tr>
                      <tr valign="baseline"> 
                        <td nowrap align="right">RIGHT SIZE:</td>
                        <td> <input name="CON_RIGHT_SIZE" type="text" class="form" value="<%=(rsSiteConfig.Fields.Item("CON_RIGHT_SIZE").Value)%>" size="10">
                        </td>
                      </tr>
                      <tr valign="baseline"> 
                        <td nowrap align="right">PAYPAL ID:</td>
                        <td> <input name="CON_PAYPAL_ID" type="text" class="form" value="<%=(rsSiteConfig.Fields.Item("CON_PAYPAL_ID").Value)%>" size="50">
                        </td>
                      </tr>
                      <tr valign="baseline"> 
                        <td nowrap align="right"> CURRENCY:</td>
                        <td> <input name="CON_PAYPAL_CURRENCY" type="text" class="form" value="<%=(rsSiteConfig.Fields.Item("CON_PAYPAL_CURRENCY").Value)%>" size="15">
                        </td>
                      </tr>
                      <tr valign="baseline"> 
                        <td nowrap align="right"> CURRENCY SIGN:</td>
                        <td> <input name="CON_PAYPAL_CURRENCY_SIGN" type="text" class="form" value="<%=(rsSiteConfig.Fields.Item("CON_PAYPAL_CURRENCY_SIGN").Value)%>" size="5">
                        </td>
                      </tr>
                      <tr valign="baseline"> 
                        <td nowrap align="right">PAYPAL RETURN SUCCESS:</td>
                        <td> <input name="CON_PAYPAL_RETURN_SUCCESS" type="text" class="form" value="<%=(rsSiteConfig.Fields.Item("CON_PAYPAL_RETURN_SUCCESS").Value)%>" size="60">
                        </td>
                      </tr>
                      <tr valign="baseline"> 
                        <td nowrap align="right">PAYPAL RETURN CANCEL:</td>
                        <td> <input name="CON_PAYPAL_RETURN_CANCEL" type="text" class="form" value="<%=(rsSiteConfig.Fields.Item("CON_PAYPAL_RETURN_CANCEL").Value)%>" size="60">
                        </td>
                      </tr>
                      <tr valign="baseline"> 
                        <td nowrap align="right">&nbsp;</td>
                        <td> <input type="submit" class="button" onClick="MM_validateForm('CON_PAGE_SIZE','','R','CON_TITLE','','R','CON_ADMIN_EMAIL','','RisEmail','CON_LEFT_SIZE','','R','CON_RIGHT_SIZE','','R','CON_PAYPAL_ID','','RisEmail','CON_PAYPAL_CURRENCY','','R','CON_PAYPAL_CURRENCY_SIGN','','R','CON_PAYPAL_RETURN_SUCCESS','','R','CON_PAYPAL_RETURN_CANCEL','','R');return document.MM_returnValue" value="Save Changes"> </td>
                      </tr>
                    </table>
                    <input type="hidden" name="MM_update" value="form1">
                    <input type="hidden" name="MM_recordId" value="<%= rsSiteConfig.Fields.Item("CON_ID").Value %>">
                  </form>
                  <p>&nbsp;</p></td>
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

<%
rsSiteConfig.Close()
Set rsSiteConfig = Nothing
%>
