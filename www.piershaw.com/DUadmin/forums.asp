<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../Connections/connDUportal.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="admin"
MM_authFailedURL="../DUhome/default.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (false Or CStr(Session("MM_UserAuthorization"))="") Or _
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
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_connDUportal_STRING
  MM_editTable = "MESSAGES"
  MM_editColumn = "MSG_ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "forums.asp"

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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then

  MM_editConnection = MM_connDUportal_STRING
  MM_editTable = "FORUMS"
  MM_editRedirectUrl = ""
  MM_fieldsStr  = "FOR_NAME|value|FOR_MODERATOR|value|FOR_DESCRIPTION|value"
  MM_columnsStr = "FOR_NAME|',none,''|FOR_MODERATOR|',none,''|FOR_DESCRIPTION|',none,''"

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
' *** Update Record: set variables

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_connDUportal_STRING
  MM_editTable = "MESSAGES"
  MM_editColumn = "MSG_ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "forums.asp"
  MM_fieldsStr  = "FOR_ID|value|MSG_DATE|value|MSG_AUTHOR|value|MSG_SUBJECT|value|MSG_BODY|value"
  MM_columnsStr = "FOR_ID|none,none,NULL|MSG_DATE|',none,NULL|MSG_AUTHOR|',none,''|MSG_SUBJECT|',none,''|MSG_BODY|',none,''"

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
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the delete
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
' *** Insert Record: construct a sql insert statement and execute it

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
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
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End if
    MM_tableValues = MM_tableValues & MM_columns(i)
    MM_dbValues = MM_dbValues & FormVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
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
Dim rsAdminMsg__MMColParam
rsAdminMsg__MMColParam = "1"
if (Request.QueryString("MSG_ID") <> "") then rsAdminMsg__MMColParam = Request.QueryString("MSG_ID")
%>
<%
set rsAdminMsg = Server.CreateObject("ADODB.Recordset")
rsAdminMsg.ActiveConnection = MM_connDUportal_STRING
rsAdminMsg.Source = "SELECT * FROM MESSAGES WHERE MSG_ID = " + Replace(rsAdminMsg__MMColParam, "'", "''") + ""
rsAdminMsg.CursorType = 0
rsAdminMsg.CursorLocation = 2
rsAdminMsg.LockType = 3
rsAdminMsg.Open()
rsAdminMsg_numRows = 0
%>
<%
Dim rsAdminEditMsg__MMColParam
rsAdminEditMsg__MMColParam = "1"
if (Request.QueryString("ID") <> "") then rsAdminEditMsg__MMColParam = Request.QueryString("ID")
%>
<%
set rsAdminEditMsg = Server.CreateObject("ADODB.Recordset")
rsAdminEditMsg.ActiveConnection = MM_connDUportal_STRING
rsAdminEditMsg.Source = "SELECT * FROM MESSAGES WHERE MSG_ID = " + Replace(rsAdminEditMsg__MMColParam, "'", "''") + ""
rsAdminEditMsg.CursorType = 0
rsAdminEditMsg.CursorLocation = 2
rsAdminEditMsg.LockType = 3
rsAdminEditMsg.Open()
rsAdminEditMsg_numRows = 0
%>
<html>
<head>
<title>DUportal</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/default.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="top" height="60" bgcolor="#009999"><img src="../assets/DUportalAdmin.gif" width="268" height="60"></td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr> 
    <td align="left" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="200" align="left" valign="top"> 
            <!--#include file="inc_left.asp" -->
          </td>
          <td width="1" bgcolor="#000000"><img src="../assets/verticalBar.gif" width="1" height="5"></td>
          <td align="left" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td align="left" valign="middle" bgcolor="#00CC99" height="20" colspan="2"> 
                  <div class = "login">&nbsp;<b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="whatsnew.asp">HOME</a> 
                    | <a href="users.asp">USERS</a> | <a href="banners.asp">BANNERS</a> 
                    | <a href="links.asp">LINKS</a> | <a href="forums.asp">FORUMS</a> 
                    | <a href="news.asp">NEWS</a> | <a href="polls.asp">POLLS</a></font></b></div>
                </td>
              </tr>
              <tr> 
                <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="middle" height="20" bgcolor="#CCCCCC"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp;MANAGING 
                  FORUMS</font></b></td>
                <td align="right" valign="middle" height="20" bgcolor="#CCCCCC"> 
                  <font face="Verdana, Arial, Helvetica, sans-serif"> <font size="1"> 
                  &nbsp; </font> </font> </td>
              </tr>
              <tr> 
                <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="top" colspan="2"> 
                  <table width="100%" border="0" cellspacing="5" cellpadding="5">
                    <tr> 
                      <td align="left" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#FF0000">Delete 
                        a message: </font></b></font><br>
                        <form name="form1" method="get" action="forums.asp">
                          <font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2">ENTER 
                          MESSAGE ID:</font></b></font> 
                          <input type="text" name="MSG_ID" size="10">
                          <input type="submit" name="Submit" value="Find">
                        </form>
                      </td>
                    </tr>
                    <% If Not rsAdminMsg.EOF Or Not rsAdminMsg.BOF Then %>
                    <tr> 
                      <td align="left" valign="top"> 
                        <form name="form2" method="POST" action="<%=MM_editAction%>">
                          <table width="100%" border="0" cellspacing="5" cellpadding="5">
                            <tr align="left" valign="middle"> 
                              <td align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">MSG 
                                ID</font></td>
                              <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#009999"><%=(rsAdminMsg.Fields.Item("MSG_ID").Value)%></font></td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">FOR 
                                ID</font></td>
                              <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#009999"><%=(rsAdminMsg.Fields.Item("FOR_ID").Value)%></font></td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">DATED</font></td>
                              <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#009999"><%=(rsAdminMsg.Fields.Item("MSG_DATE").Value)%></font></td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">AUTHOR</font></td>
                              <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#009999"><%=(rsAdminMsg.Fields.Item("MSG_AUTHOR").Value)%></font></td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">SUBJECT</font></td>
                              <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#009999"><%=(rsAdminMsg.Fields.Item("MSG_SUBJECT").Value)%></font></td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">BODY</font></td>
                              <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#009999"><%=(rsAdminMsg.Fields.Item("MSG_BODY").Value)%></font></td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td align="right">&nbsp;</td>
                              <td> 
                                <input type="submit" name="Submit2" value="Delete">
                              </td>
                            </tr>
                          </table>
                          <input type="hidden" name="MM_delete" value="true">
                          <input type="hidden" name="MM_recordId" value="<%= rsAdminMsg.Fields.Item("MSG_ID").Value %>">
                        </form>
                      </td>
                    </tr>
                    <% End If ' end Not rsAdminMsg.EOF Or NOT rsAdminMsg.BOF %>
                  </table>
                  <p>&nbsp;</p>
                </td>
              </tr>
              <tr bgcolor="#000000"> 
                <td align="left" valign="top" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="top" colspan="2"> 
                  <table width="100%" border="0" cellspacing="5" cellpadding="5">
                    <tr> 
                      <td align="left" valign="top"><font size="2" color="#FF0000"><b><font face="Verdana, Arial, Helvetica, sans-serif">Add 
                        a new forum:</font></b></font><br>
                        <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#FF0000"> 
                        <form method="post" action="<%=MM_editAction%>" name="form3">
                          <table align="center" width="100%" cellpadding="5" cellspacing="5">
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">FOR_NAME:</font></td>
                              <td> 
                                <input type="text" name="FOR_NAME" value="" size="32">
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">FOR_MODERATOR:</font></td>
                              <td> 
                                <input type="text" name="FOR_MODERATOR" value="demo_admin" size="32">
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">FOR_DESCRIPTION:</font></td>
                              <td> 
                                <textarea name="FOR_DESCRIPTION" cols="45" rows="2"></textarea>
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right">&nbsp;</td>
                              <td> 
                                <input type="submit" value="Insert Record" name="submit">
                              </td>
                            </tr>
                          </table>
                          <input type="hidden" name="MM_insert" value="true">
                        </form>
                        </font></b></font> </td>
                    </tr>
                  </table>
                  <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#FF0000"> 
                  <p>&nbsp;</p>
                  </font></b></font></td>
              </tr>
              <tr bgcolor="#000000"> 
                <td align="left" valign="top" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="top" colspan="2"> 
                  <table width="100%" border="0" cellspacing="5" cellpadding="5">
                    <tr> 
                      <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FF0000"><b>Edit 
                        a message:<br>
                        </b></font> 
                        <form name="form4" method="get" action="forums.asp">
                          <b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">ENRER 
                          MESSAGE ID: </font> </b> 
                          <input type="text" name="ID" size="5">
                          <input type="submit" name="Submit3" value="Find">
                        </form>
                        <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FF0000"><b> 
                        </b></font></td>
                    </tr>
                    <% If Not rsAdminEditMsg.EOF Or Not rsAdminEditMsg.BOF Then %>
                    <tr> 
                      <td>&nbsp; 
                        <form method="post" action="<%=MM_editAction%>" name="form5">
                          <table align="center" cellpadding="5" cellspacing="5" width="100%">
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">FOR_ID:</font></td>
                              <td> 
                                <input type="text" name="FOR_ID" value="<%=(rsAdminEditMsg.Fields.Item("FOR_ID").Value)%>" size="32">
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">MSG_DATE:</font></td>
                              <td> 
                                <input type="text" name="MSG_DATE" value="<%=(rsAdminEditMsg.Fields.Item("MSG_DATE").Value)%>" size="32">
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">MSG_AUTHOR:</font></td>
                              <td> 
                                <input type="text" name="MSG_AUTHOR" value="<%=(rsAdminEditMsg.Fields.Item("MSG_AUTHOR").Value)%>" size="32">
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">MSG_SUBJECT:</font></td>
                              <td> 
                                <input type="text" name="MSG_SUBJECT" value="<%=(rsAdminEditMsg.Fields.Item("MSG_SUBJECT").Value)%>" size="32">
                              </td>
                            </tr>
                            <tr> 
                              <td nowrap align="right" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">MSG_BODY:</font></td>
                              <td valign="baseline"> 
                                <textarea name="MSG_BODY" cols="50" rows="5"><%=(rsAdminEditMsg.Fields.Item("MSG_BODY").Value)%></textarea>
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right">&nbsp;</td>
                              <td> 
                                <input type="submit" value="Update Message">
                              </td>
                            </tr>
                          </table>
                          <input type="hidden" name="MM_update" value="true">
                          <input type="hidden" name="MM_recordId" value="<%= rsAdminEditMsg.Fields.Item("MSG_ID").Value %>">
                        </form>
                        <p>&nbsp;</p>
                      </td>
                    </tr>
                    <% End If ' end Not rsAdminEditMsg.EOF Or NOT rsAdminEditMsg.BOF %>
                  </table>
                </td>
              </tr>
              <tr> 
                <td align="left" valign="top" colspan="2">&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
<%
rsAdminMsg.Close()
%>
<%
rsAdminEditMsg.Close()
%>

