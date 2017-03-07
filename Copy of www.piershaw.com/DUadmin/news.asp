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
  MM_editTable = "NEWS"
  MM_editColumn = "NEWS_ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "news.asp"

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
  MM_editTable = "NEWS_TYPES"
  MM_editRedirectUrl = "news.asp"
  MM_fieldsStr  = "TYPE_NAME|value"
  MM_columnsStr = "TYPE_NAME|',none,''"

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
  MM_editTable = "NEWS"
  MM_editColumn = "NEWS_ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "news.asp"
  MM_fieldsStr  = "NEWS_TYPE|value|NEWS_TITLE|value|NEWS_DESC|value|NEWS_URL|value|NEWS_DATE|value|NEWS_SOURCE|value|NEWS_ADDER|value|select|value"
  MM_columnsStr = "NEWS_TYPE|none,none,NULL|NEWS_TITLE|',none,''|NEWS_DESC|',none,''|NEWS_URL|',none,''|NEWS_DATE|',none,NULL|NEWS_SOURCE|',none,''|NEWS_ADDER|',none,''|NEWS_APPROVED|',none,''"

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
Dim rsAdminNewDelete__MMColParam
rsAdminNewDelete__MMColParam = "1"
if (Request.QueryString("ID") <> "") then rsAdminNewDelete__MMColParam = Request.QueryString("ID")
%>
<%
set rsAdminNewDelete = Server.CreateObject("ADODB.Recordset")
rsAdminNewDelete.ActiveConnection = MM_connDUportal_STRING
rsAdminNewDelete.Source = "SELECT * FROM NEWS WHERE NEWS_ID = " + Replace(rsAdminNewDelete__MMColParam, "'", "''") + ""
rsAdminNewDelete.CursorType = 0
rsAdminNewDelete.CursorLocation = 2
rsAdminNewDelete.LockType = 3
rsAdminNewDelete.Open()
rsAdminNewDelete_numRows = 0
%>
<%
Dim rsAdminNewsEdit__MMColParam
rsAdminNewsEdit__MMColParam = "1"
if (Request.QueryString("news") <> "") then rsAdminNewsEdit__MMColParam = Request.QueryString("news")
%>
<%
set rsAdminNewsEdit = Server.CreateObject("ADODB.Recordset")
rsAdminNewsEdit.ActiveConnection = MM_connDUportal_STRING
rsAdminNewsEdit.Source = "SELECT * FROM NEWS WHERE NEWS_ID = " + Replace(rsAdminNewsEdit__MMColParam, "'", "''") + ""
rsAdminNewsEdit.CursorType = 0
rsAdminNewsEdit.CursorLocation = 2
rsAdminNewsEdit.LockType = 3
rsAdminNewsEdit.Open()
rsAdminNewsEdit_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = 5
Dim Repeat1__index
Repeat1__index = 0
rsAdminBanners_numRows = rsAdminBanners_numRows + Repeat1__numRows
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
                  NEWS </font></b></td>
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
                      <td><font color="#FF0000" face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2">Delete 
                        a news:</font></b></font><font face="Verdana, Arial, Helvetica, sans-serif"><br>
                        </font> 
                        <form name="form1" method="get" action="news.asp">
                          <font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2">ENTER 
                          NEWS ID: 
                          <input type="text" name="ID" size="5">
                          </font></b> <font size="2"> 
                          <input type="submit" name="Submit" value="Submit">
                          </font> </font> 
                        </form>
                      </td>
                    </tr>
                    <tr> 
                      <% If Not rsAdminNewDelete.EOF Or Not rsAdminNewDelete.BOF Then %>
                      <td> 
                        <form name="form2" method="POST" action="<%=MM_editAction%>">
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Title:</b> 
                          <%=(rsAdminNewDelete.Fields.Item("NEWS_TITLE").Value)%><br>
                          <br>
                          <b>Description:</b> <%=(rsAdminNewDelete.Fields.Item("NEWS_DESC").Value)%><br>
                          <input type="submit" name="Submit2" value="Delete">
                          </font> <font face="Verdana, Arial, Helvetica, sans-serif"> 
                          <input type="hidden" name="MM_delete" value="true">
                          <input type="hidden" name="MM_recordId" value="<%= rsAdminNewDelete.Fields.Item("NEWS_ID").Value %>">
                          </font> 
                        </form>
                      </td>
                      <% End If ' end Not rsAdminNewDelete.EOF Or NOT rsAdminNewDelete.BOF %>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr bgcolor="#000000"> 
                <td align="left" valign="top" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="top" colspan="2"> 
                  <table width="100%" border="0" cellspacing="5" cellpadding="5">
                    <tr> 
                      <td><font color="#FF0000"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Add 
                        new news type:</font></b></font><br>
                        <form method="post" action="<%=MM_editAction%>" name="form3">
                          <table align="center" width="100%" cellpadding="5" cellspacing="5">
                            <tr valign="baseline"> 
                              <td nowrap align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">TYPE_NAME:</font></b></td>
                              <td> 
                                <input type="text" name="TYPE_NAME" value="" size="50" maxlength="50">
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right">&nbsp;</td>
                              <td> 
                                <input type="submit" value="Insert Record">
                              </td>
                            </tr>
                          </table>
                          <input type="hidden" name="MM_insert" value="true">
                        </form>
                        <p>&nbsp;</p>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr bgcolor="#000000"> 
                <td align="left" valign="top" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="top" colspan="2"> 
                  <table width="100%" border="0" cellspacing="5" cellpadding="5">
                    <tr> 
                      <td><font color="#FF0000"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Edit 
                        a news:</font></b></font><br>
                        <form name="form1" method="get" action="news.asp">
                          <b>ENTER NEWS ID: 
                          <input type="text" name="news" size="5">
                          </b> 
                          <input type="submit" name="Submit3" value="Submit">
                        </form>
                      </td>
                    </tr>
                    <% If Not rsAdminNewsEdit.EOF Or Not rsAdminNewsEdit.BOF Then %>
                    <tr> 
                      <td>&nbsp; 
                        <form method="POST" action="<%=MM_editAction%>" name="form4">
                          <table align="center" cellpadding="5" cellspacing="5" width="100%">
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">NEWS_TYPE:</font></td>
                              <td> 
                                <input type="text" name="NEWS_TYPE" value="<%=(rsAdminNewsEdit.Fields.Item("NEWS_TYPE").Value)%>" size="32">
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">NEWS_TITLE:</font></td>
                              <td> 
                                <input type="text" name="NEWS_TITLE" value="<%=(rsAdminNewsEdit.Fields.Item("NEWS_TITLE").Value)%>" size="32">
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">NEWS_DESC:</font></td>
                              <td> 
                                <textarea name="NEWS_DESC" cols="32" rows="3"><%=(rsAdminNewsEdit.Fields.Item("NEWS_DESC").Value)%></textarea>
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">NEWS_URL:</font></td>
                              <td> 
                                <input type="text" name="NEWS_URL" value="<%=(rsAdminNewsEdit.Fields.Item("NEWS_URL").Value)%>" size="32">
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">NEWS_DATE:</font></td>
                              <td> 
                                <input type="text" name="NEWS_DATE" value="<%=(rsAdminNewsEdit.Fields.Item("NEWS_DATE").Value)%>" size="32">
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">NEWS_SOURCE:</font></td>
                              <td> 
                                <input type="text" name="NEWS_SOURCE" value="<%=(rsAdminNewsEdit.Fields.Item("NEWS_SOURCE").Value)%>" size="32">
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">NEWS_ADDER:</font></td>
                              <td> 
                                <input type="text" name="NEWS_ADDER" value="<%=(rsAdminNewsEdit.Fields.Item("NEWS_ADDER").Value)%>" size="32">
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">NEWS_APPROVED:</font></td>
                              <td> 
                                <select name="select">
                                  <option value="1" selected>Yes</option>
                                  <option value="0">No</option>
                                  <%
While (NOT rsAdminNewDelete.EOF)
%>
                                  <option value="<%=(rsAdminNewDelete.Fields.Item("NEWS_APPROVED").Value)%>" <%if (CStr(rsAdminNewDelete.Fields.Item("NEWS_APPROVED").Value) = CStr(rsAdminNewsEdit.Fields.Item("NEWS_APPROVED").Value)) then Response.Write("SELECTED") : Response.Write("")%>><%=(rsAdminNewDelete.Fields.Item("NEWS_APPROVED").Value)%></option>
                                  <%
  rsAdminNewDelete.MoveNext()
Wend
If (rsAdminNewDelete.CursorType > 0) Then
  rsAdminNewDelete.MoveFirst
Else
  rsAdminNewDelete.Requery
End If
%>
                                </select>
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right">&nbsp;</td>
                              <td> 
                                <input type="submit" value="Update News">
                              </td>
                            </tr>
                          </table>
                          <input type="hidden" name="MM_update" value="true">
                          <input type="hidden" name="MM_recordId" value="<%= rsAdminNewsEdit.Fields.Item("NEWS_ID").Value %>">
                        </form>
                        <p>&nbsp;</p>
                      </td>
                    </tr>
                    <% End If ' end Not rsAdminNewsEdit.EOF Or NOT rsAdminNewsEdit.BOF %>
                  </table>
                </td>
              </tr>
              <tr> 
                <td align="left" valign="top" colspan="2">&nbsp;</td>
              </tr>
              <tr> 
                <td align="left" valign="top" colspan="2">&nbsp;</td>
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
<p>&nbsp;</p>
</body>
</html>
<%
rsAdminNewDelete.Close()
%>
<%
rsAdminNewsEdit.Close()
%>
