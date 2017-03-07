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
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="default.asp"
MM_grantAccess=false
If Session("MM_AdminUser") = "admin" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
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
  MM_editTable = "USERS"
  MM_editColumn = "U_ID"
  MM_recordId = "'" + Request.Form("MM_recordId") + "'"
  MM_editRedirectUrl = "users.asp"
  MM_fieldsStr  = "U_ID|value|U_PASSWORD|value|U_ACCESS|value|U_FIRST|value|U_LAST|value|U_EMAIL|value|U_DATED|value|U_ADDRESS|value|U_COUNTRY|value"
  MM_columnsStr = "U_ID|',none,''|U_PASSWORD|',none,''|U_ACCESS|',none,''|U_FIRST|',none,''|U_LAST|',none,''|U_EMAIL|',none,''|U_DATED|',none,NULL|U_ADDRESS|',none,''|U_COUNTRY|',none,''"

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
Dim rsUserEdit__MMColParam
rsUserEdit__MMColParam = "1"
If (Request.QueryString("iUser") <> "") Then 
  rsUserEdit__MMColParam = Request.QueryString("iUser")
End If
%>
<%
Dim rsUserEdit
Dim rsUserEdit_numRows

Set rsUserEdit = Server.CreateObject("ADODB.Recordset")
rsUserEdit.ActiveConnection = MM_connDUportal_STRING
rsUserEdit.Source = "SELECT * FROM USERS WHERE U_ID = '" + Replace(rsUserEdit__MMColParam, "'", "''") + "'"
rsUserEdit.CursorType = 0
rsUserEdit.CursorLocation = 2
rsUserEdit.LockType = 1
rsUserEdit.Open()

rsUserEdit_numRows = 0
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
                      <td align="left" valign="middle" class="textBoldColor">EDIT 
                        USER </td>
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
                <td align="left" valign="top" class="bgTable"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="bgTable">
                          <tr> 
						  <form method="post" action="<%=MM_editAction%>" name="form1">
                            <td align="left" valign="top">
                                <table align="center" cellpadding="3" cellspacing="3" class="textBold">
                                  <tr valign="baseline"> 
                                    <td nowrap align="right">USER ID:</td>
                                    <td> <input name="U_ID" type="text" class="form" value="<%=(rsUserEdit.Fields.Item("U_ID").Value)%>" size="20">
                                    </td>
                                  </tr>
                                  <tr valign="baseline"> 
                                    <td nowrap align="right">PASSWORD:</td>
                                    <td> <input name="U_PASSWORD" type="text" class="form" value="<%=(rsUserEdit.Fields.Item("U_PASSWORD").Value)%>" size="20"></td>
                                  </tr>
                                  <tr valign="baseline"> 
                                    <td nowrap align="right">ACCESS LEVEL:</td>
                                    <td> <input name="U_ACCESS" type="text" class="form" value="<%=(rsUserEdit.Fields.Item("U_ACCESS").Value)%>" size="10">
                                    </td>
                                  </tr>
                                  <tr valign="baseline"> 
                                    <td nowrap align="right">FIRST NAME:</td>
                                    <td> <input name="U_FIRST" type="text" class="form" value="<%=(rsUserEdit.Fields.Item("U_FIRST").Value)%>" size="40">
                                    </td>
                                  </tr>
                                  <tr valign="baseline"> 
                                    <td nowrap align="right">LAST NAME:</td>
                                    <td> <input name="U_LAST" type="text" class="form" value="<%=(rsUserEdit.Fields.Item("U_LAST").Value)%>" size="40">
                                    </td>
                                  </tr>
                                  <tr valign="baseline"> 
                                    <td nowrap align="right">EMAIL:</td>
                                    <td> <input name="U_EMAIL" type="text" class="form" value="<%=(rsUserEdit.Fields.Item("U_EMAIL").Value)%>" size="50">
                                    </td>
                                  </tr>
                                  <tr valign="baseline"> 
                                    <td nowrap align="right">DATED:</td>
                                    <td> <input name="U_DATED" type="text" class="form" value="<%=(rsUserEdit.Fields.Item("U_DATED").Value)%>" size="15">
                                    </td>
                                  </tr>
                                  <tr valign="baseline"> 
                                    <td nowrap align="right">HOME ADDRESS:</td>
                                    <td> <input name="U_ADDRESS" type="text" class="form" value="<%=(rsUserEdit.Fields.Item("U_ADDRESS").Value)%>" size="50">
                                    </td>
                                  </tr>
                                  <tr valign="baseline"> 
                                    <td nowrap align="right">COUNTRY:</td>
                                    <td> <input name="U_COUNTRY" type="text" class="form" value="<%=(rsUserEdit.Fields.Item("U_COUNTRY").Value)%>" size="30">
                                    </td>
                                  </tr>
                                  <tr valign="baseline"> 
                                    <td nowrap align="right">&nbsp;</td>
                                    <td> <input name="Save" type="submit" class="button" id="Save" onClick="MM_validateForm('U_ID','','R','U_PASSWORD','','R','U_ACCESS','','R','U_FIRST','','R','U_LAST','','R','U_EMAIL','','RisEmail','U_DATED','','R','U_COUNTRY','','R');return document.MM_returnValue" value="Save Changes"> 
                                    </td>
                                  </tr>
                                </table>
                                <input type="hidden" name="MM_update" value="form1">
                                <input type="hidden" name="MM_recordId" value="<%= rsUserEdit.Fields.Item("U_ID").Value %>">
                              
                             </td>
							 </form>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
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
rsUserEdit.Close()
Set rsUserEdit = Nothing
%>
