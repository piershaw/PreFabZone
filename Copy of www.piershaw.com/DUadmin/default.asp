<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/connDUportal.asp" -->
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Request.QueryString
MM_valUsername=CStr(Request.Form("id"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization="ACCESS"
  MM_redirectLoginSuccess="whatsnew.asp"
  MM_redirectLoginFailed="default.asp"
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_connDUportal_STRING
  MM_rsUser.Source = "SELECT U_ID, U_PASSWORD"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM USERS WHERE U_ID='" & MM_valUsername &"' AND U_PASSWORD='" & CStr(Request.Form("pass")) & "'"
  MM_rsUser.CursorType = 0
  MM_rsUser.CursorLocation = 2
  MM_rsUser.LockType = 3
  MM_rsUser.Open
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And true Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<html>
<head>
<title>DUportal</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#009999" text="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr> 
    <td align="center" valign="middle"> 
      <form name="LOGIN" method="post" action="<%=MM_LoginAction%>">
        <table width="400" border="0" cellspacing="5" cellpadding="5">
          <tr align="left" valign="middle"> 
            <td colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="5" color="#FF0000">DU<font color="#00FF00">portal</font></font><font face="Verdana, Arial, Helvetica, sans-serif" size="5" color="#00FF00"> 
              <font color="#FFFFFF">Login:</font></font></td>
          </tr>
          <tr align="left" valign="middle"> 
            <td align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">Admin</font></b></td>
            <td> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
              <input type="text" name="id" size="10" maxlength="10" value = "demo_admin">
              </font></b></td>
          </tr>
          <tr align="left" valign="middle"> 
            <td align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">Password</font></b></td>
            <td> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
              <input type="password" name="pass" size="10" maxlength="10" value = "password">
              </font></b></td>
          </tr>
          <tr align="left" valign="middle"> 
            <td align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font></b></td>
            <td> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
              <input type="submit" name="Submit" value="Login">
              </font></b></td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table>
</body>
</html>

