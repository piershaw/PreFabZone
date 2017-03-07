<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/connDUportal.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="default.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
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

MM_editAction = CStr(Request("URL"))
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

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_connDUportal_STRING
  MM_editTable = "USERS"
  MM_editColumn = "U_ID"
  MM_recordId = "'" + Request.Form("MM_recordId") + "'"
  MM_editRedirectUrl = "myInfoUpdated.asp"
  MM_fieldsStr  = "U_PASSWORD|value|U_FIRST|value|U_LAST|value|U_EMAIL|value"
  MM_columnsStr = "U_PASSWORD|',none,''|U_FIRST|',none,''|U_LAST|',none,''|U_EMAIL|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(i+1) = CStr(Request.Form(MM_fields(i)))
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
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    FormVal = MM_fields(i+1)
    MM_typeArray = Split(MM_columns(i+1),",")
    Delim = MM_typeArray(0)
    If (Delim = "none") Then Delim = ""
    AltVal = MM_typeArray(1)
    If (AltVal = "none") Then AltVal = ""
    EmptyVal = MM_typeArray(2)
    If (EmptyVal = "none") Then EmptyVal = ""
    If (FormVal = "") Then
      FormVal = EmptyVal
    Else
      If (AltVal <> "") Then
        FormVal = AltVal
      ElseIf (Delim = "'") Then  ' escape quotes
        FormVal = "'" & Replace(FormVal,"'","''") & "'"
      Else
        FormVal = Delim + FormVal + Delim
      End If
    End If
    If (i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(i) & " = " & FormVal
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
Dim rsMyInfo__MMColParam
rsMyInfo__MMColParam = "1"
if (Session("MM_Username") <> "") then rsMyInfo__MMColParam = Session("MM_Username")
%>
<%
set rsMyInfo = Server.CreateObject("ADODB.Recordset")
rsMyInfo.ActiveConnection = MM_connDUportal_STRING
rsMyInfo.Source = "SELECT * FROM USERS WHERE U_ID = '" + Replace(rsMyInfo__MMColParam, "'", "''") + "'"
rsMyInfo.CursorType = 0
rsMyInfo.CursorLocation = 2
rsMyInfo.LockType = 3
rsMyInfo.Open()
rsMyInfo_numRows = 0
%>
<% Response.Buffer = "true" %>
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>	
function DoTrimProperly(str, nNamedFormat, properly, pointed, points)
  dim strRet
  strRet = Server.HTMLEncode(str)
  strRet = replace(strRet, vbcrlf,"")
  strRet = replace(strRet, vbtab,"")
  If (LEN(strRet) > nNamedFormat) Then
    strRet = LEFT(strRet, nNamedFormat)			
    If (properly = 1) Then					
      Dim TempArray								
      TempArray = split(strRet, " ")	
      Dim n
      strRet = ""
      for n = 0 to Ubound(TempArray) - 1
        strRet = strRet & " " & TempArray(n)
      next
    End If
    If (pointed = 1) Then
      strRet = strRet & points
    End If
  End If
  DoTrimProperly = strRet
End Function
</SCRIPT>
<html>
<head>
<title>DUportal</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/default.css" type="text/css">
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
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
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr valign="top"> 
    <td align="left" class = "bg_banner" height="62" valign="middle"> 
      <!--#include file="../includes/inc_header.asp" -->
    </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr valign="middle"> 
    <td align="left" class = "bg_navigator" height="20"> 
      <!--#include file="../includes/inc_navigator.asp" -->
    </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr> 
    <td align="left" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr align="left" valign="top"> 
          <td width="200"> 
            <!--#include file="inc_left.asp" -->
          </td>
          <td bgcolor="#000000" width="1"><img src="../assets/verticalBar.gif" width="1" height="5"></td>
          <td> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td align="left" valign="top" class = "bg_login" height="30"> 
                  <!--#include file="../includes/inc_login.asp" -->
                </td>
              </tr>
              <tr> 
                <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="top"> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td align="left" valign="middle" height="20" class = "bg_navigator"><font face="Verdana, Arial, Helvetica, sans-serif" size="3"><b><font size="1">&nbsp;UPDATE 
                        MY INFO</font></b></font></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                    </tr>
                    <tr> 
                      <td align="center" valign="top">&nbsp; 
                        <form method="POST" action="<%=MM_editAction%>" name="form1">
                          <table align="center" cellpadding="5" cellspacing="3">
                            <tr valign="baseline"> 
                              <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">User 
                                ID:</font></b></td>
                              <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#009999"><%=(rsMyInfo.Fields.Item("U_ID").Value)%></font></b></font></td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Password:</font></b></td>
                              <td> 
                                <input type="password" name="U_PASSWORD" size="10" maxlength="10" value="<%=(rsMyInfo.Fields.Item("U_PASSWORD").Value)%>">
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">First 
                                Name:</font></b></td>
                              <td> 
                                <input type="text" name="U_FIRST" value="<%=(rsMyInfo.Fields.Item("U_FIRST").Value)%>" size="40" maxlength="40">
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Last 
                                Name:</font></b></td>
                              <td> 
                                <input type="text" name="U_LAST" value="<%=(rsMyInfo.Fields.Item("U_LAST").Value)%>" size="40" maxlength="40">
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Email:</font></b></td>
                              <td> 
                                <input type="text" name="U_EMAIL" value="<%=(rsMyInfo.Fields.Item("U_EMAIL").Value)%>" size="50" maxlength="50">
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right">&nbsp;</td>
                              <td> 
                                <input type="submit" value="Update" onClick="MM_validateForm('U_FIRST','','R','U_LAST','','R','U_EMAIL','','RisEmail','U_PASSWORD','','R');return document.MM_returnValue">
                              </td>
                            </tr>
                          </table>
                          <input type="hidden" name="MM_update" value="true">
                          <input type="hidden" name="MM_recordId" value="<%= rsMyInfo.Fields.Item("U_ID").Value %>">
                        </form>
                        <p>&nbsp;</p>
                      </td>
                    </tr>
                    <tr> 
                      <td align="left" valign="top">&nbsp;</td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td height="40"> 
      <!--#include file="../includes/inc_footer.asp" -->
    </td>
  </tr>
</table>
</body>
</html>
<%
rsMyInfo.Close()
%>

