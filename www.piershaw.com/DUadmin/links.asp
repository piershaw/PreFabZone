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
' *** Edit Operations: declare variables

MM_editActionCat = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editActionCat = MM_editActionCat & "?" & Request.QueryString
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
  MM_editTable = "LINKS"
  MM_editColumn = "LINK_ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "links.asp"

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
  MM_editTable = "LINK_SUBS"
  MM_editRedirectUrl = "links.asp"
  MM_fieldsStr  = "CAT_ID|value|SUB_NAME|value"
  MM_columnsStr = "CAT_ID|none,none,NULL|SUB_NAME|',none,''"

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
  MM_editTable = "LINKS"
  MM_editColumn = "LINK_ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "links.asp"
  MM_fieldsStr  = "CAT_ID|value|SUB_ID|value|LINK_NAME|value|LINK_URL|value|LINK_DESC|value"
  MM_columnsStr = "CAT_ID|none,none,NULL|SUB_ID|none,none,NULL|LINK_NAME|',none,''|LINK_URL|',none,''|LINK_DESC|',none,''"

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
' *** Insert Record: set variables

If (CStr(Request("MM_insertCat")) <> "") Then

  MM_editConnection = MM_connDUportal_STRING
  MM_editTable = "LINK_CATS"
  MM_editRedirectUrl = "links.asp"
  MM_fieldsStr  = "CAT_NAME|value"
  MM_columnsStr = "CAT_NAME|',none,''"

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
' *** Insert Record: construct a sql insert statement and execute it

If (CStr(Request("MM_insertCat")) <> "") Then

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
Dim rsAdminLink__MMColParam
rsAdminLink__MMColParam = "1"
if (Request.QueryString("ID") <> "") then rsAdminLink__MMColParam = Request.QueryString("ID")
%>
<%
set rsAdminLink = Server.CreateObject("ADODB.Recordset")
rsAdminLink.ActiveConnection = MM_connDUportal_STRING
rsAdminLink.Source = "SELECT * FROM LINKS WHERE LINK_ID = " + Replace(rsAdminLink__MMColParam, "'", "''") + ""
rsAdminLink.CursorType = 0
rsAdminLink.CursorLocation = 2
rsAdminLink.LockType = 3
rsAdminLink.Open()
rsAdminLink_numRows = 0
%>
<%
set rsCat = Server.CreateObject("ADODB.Recordset")
rsCat.ActiveConnection = MM_connDUportal_STRING
rsCat.Source = "SELECT * FROM LINK_CATS ORDER BY CAT_ID ASC"
rsCat.CursorType = 0
rsCat.CursorLocation = 2
rsCat.LockType = 3
rsCat.Open()
rsCat_numRows = 0
%>
<%
Dim rsAdminLinkEdit__MMColParam
rsAdminLinkEdit__MMColParam = "1"
if (Request.QueryString("link") <> "") then rsAdminLinkEdit__MMColParam = Request.QueryString("link")
%>
<%
set rsAdminLinkEdit = Server.CreateObject("ADODB.Recordset")
rsAdminLinkEdit.ActiveConnection = MM_connDUportal_STRING
rsAdminLinkEdit.Source = "SELECT * FROM LINKS WHERE LINK_ID = " + Replace(rsAdminLinkEdit__MMColParam, "'", "''") + ""
rsAdminLinkEdit.CursorType = 0
rsAdminLinkEdit.CursorLocation = 2
rsAdminLinkEdit.LockType = 3
rsAdminLinkEdit.Open()
rsAdminLinkEdit_numRows = 0
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
    <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
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
                <td align="left" valign="middle" height="20" bgcolor="#CCCCCC"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp; 
                  MANAGING LINKS DIRECTORY</font></b></td>
                <td align="right" valign="middle" height="20" bgcolor="#CCCCCC"> 
                  <font face="Verdana, Arial, Helvetica, sans-serif"> <font size="1"> 
                  &nbsp; </font> </font> </td>
              </tr>
              <tr> 
                <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr valign="top" align="left"> 
                <td colspan="2"> 
                  <table border="0" cellspacing="3" cellpadding="3" width="100%">
                    <tr> 
                      <td align="left" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#FF0000">Delete 
                        a link:</font></b></font></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="top"> 
                        <form name="LINK" method="get" action="links.asp">
                          <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>ENTER 
                          LINK ID </b></font> 
                          <input type="text" name="ID" size="10">
                          <input type="submit" name="Submit" value="Find This Link">
                        </form>
                      </td>
                    </tr>
                    <% If Not rsAdminLink.EOF Or Not rsAdminLink.BOF Then %>
                    <tr> 
                      <td align="left" valign="top"> 
                        <form name="DELETE" method="POST" action="<%=MM_editAction%>">
                          <table width="100%" border="0" cellspacing="5" cellpadding="5">
                            <tr> 
                              <td valign="middle" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">ID</font></b></td>
                              <td align="left" valign="middle"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#009999"><%=(rsAdminLink.Fields.Item("LINK_ID").Value)%></font></b></td>
                            </tr>
                            <tr> 
                              <td valign="middle" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">CAT 
                                ID</font></b></td>
                              <td align="left" valign="middle"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#009999"><%=(rsAdminLink.Fields.Item("CAT_ID").Value)%></font></b></td>
                            </tr>
                            <tr> 
                              <td valign="middle" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">SUB 
                                ID</font></b></td>
                              <td align="left" valign="middle"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#009999"><%=(rsAdminLink.Fields.Item("SUB_ID").Value)%></font></b></td>
                            </tr>
                            <tr> 
                              <td valign="middle" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">NAME</font></b></td>
                              <td align="left" valign="middle"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#009999"><%=(rsAdminLink.Fields.Item("LINK_NAME").Value)%></font></b></td>
                            </tr>
                            <tr> 
                              <td valign="middle" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">URL</font></b></td>
                              <td align="left" valign="middle"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#009999"><%=(rsAdminLink.Fields.Item("LINK_URL").Value)%></font></b></td>
                            </tr>
                            <tr> 
                              <td valign="middle" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">DESCRIPTION</font></b></td>
                              <td align="left" valign="middle"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#009999"><%=(rsAdminLink.Fields.Item("LINK_DESC").Value)%></font></b></td>
                            </tr>
                            <tr> 
                              <td valign="middle" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">DATED</font></b></td>
                              <td align="left" valign="middle"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#009999"><%=(rsAdminLink.Fields.Item("LINK_DATE").Value)%></font></b></td>
                            </tr>
                            <tr> 
                              <td valign="middle" align="right"> 
                                <input type="hidden" name="hiddenField" value="<%=(rsAdminLink.Fields.Item("LINK_ID").Value)%>">
                              </td>
                              <td align="left" valign="middle"> 
                                <input type="submit" name="Submit2" value="Delete">
                              </td>
                            </tr>
                          </table>
                          <input type="hidden" name="MM_delete" value="true">
                          <input type="hidden" name="MM_recordId" value="<%= rsAdminLink.Fields.Item("LINK_ID").Value %>">
                        </form>
                      </td>
                    </tr>
                    <% End If ' end Not rsAdminLink.EOF Or NOT rsAdminLink.BOF %>
                  </table>
                </td>
              </tr>
              <tr> 
                <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="top"> 
                  <form method="POST" action="<%=MM_editAction%>" name="form1">
                    <table align="center" cellpadding="5" cellspacing="5" width="100%">
                      <tr valign="baseline" align="left"> 
                        <td nowrap colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FF0000"><b>Insert 
                          a new category:</b></font></td>
                      </tr>
                      <tr valign="baseline"> 
                        <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
                          CAT NAME:</font></b></td>
                        <td> 
                          <input type="text" name="CAT_NAME" size="45" maxlength="45">
                        </td>
                      </tr>
                      <tr valign="baseline"> 
                        <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font></b></td>
                        <td> 
                          <input type="submit" value="Insert Category">
                        </td>
                      </tr>
                    </table>
                    <input type="hidden" name="MM_insertCat" value="true">
                  </form>
                </td>
              </tr>
              <tr> 
                <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="top" colspan="2"> 
                  <form method="post" action="<%=MM_editAction%>" name="form2">
                    <table align="center" cellpadding="5" cellspacing="5" width="100%">
                      <tr valign="baseline"> 
                        <td nowrap align="left" colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FF0000"><b>Insert 
                          a new sub-category:</b></font></td>
                      </tr>
                      <tr valign="baseline"> 
                        <td nowrap align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">CAT_ID:</font></b></td>
                        <td> 
                          <select name="CAT_ID">
                            <%
While (NOT rsCat.EOF)
%>
                            <option value="<%=(rsCat.Fields.Item("CAT_ID").Value)%>" <%if (CStr(rsCat.Fields.Item("CAT_ID").Value) = CStr(rsCat.Fields.Item("CAT_NAME").Value)) then Response.Write("SELECTED") : Response.Write("")%> ><%=(rsCat.Fields.Item("CAT_NAME").Value)%></option>
                            <%
  rsCat.MoveNext()
Wend
If (rsCat.CursorType > 0) Then
  rsCat.MoveFirst
Else
  rsCat.Requery
End If
%>
                          </select>
                        </td>
                      </tr>
                      <tr valign="baseline"> 
                        <td nowrap align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">SUB 
                          NAME:</font></b></td>
                        <td> 
                          <input type="text" name="SUB_NAME" value="" size="32">
                        </td>
                      </tr>
                      <tr valign="baseline"> 
                        <td nowrap align="right">&nbsp;</td>
                        <td> 
                          <input type="submit" value="Insert Sub Category">
                        </td>
                      </tr>
                    </table>
                    <input type="hidden" name="MM_insert" value="true">
                  </form>
                  <p>&nbsp;</p>
                  <img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr bgcolor="#000000"> 
                <td align="left" valign="top" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="top" colspan="2"> 
                  <table width="100%" border="0" cellspacing="5" cellpadding="5">
                    <tr> 
                      <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#FF0000">Edit 
                        a link:<br>
                        </font></b></font> 
                        <form name="form3" method="get" action="links.asp">
                          <b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">ENTER 
                          LINK ID:</font></b> 
                          <input type="text" name="link" size="5">
                          <input type="submit" name="Submit3" value="Submit">
                        </form>
                        <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#FF0000"> 
                        </font></b></font></td>
                    </tr>
                    <% If Not rsAdminLinkEdit.EOF Or Not rsAdminLinkEdit.BOF Then %>
                    <tr> 
                      <td>&nbsp; 
                        <form method="POST" action="<%=MM_editAction%>" name="form4">
                          <table align="center" width="100%" cellpadding="5" cellspacing="5">
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">CAT_ID:</font></td>
                              <td> 
                                <input type="text" name="CAT_ID" value="<%=(rsAdminLinkEdit.Fields.Item("CAT_ID").Value)%>" size="32">
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">SUB_ID:</font></td>
                              <td> 
                                <input type="text" name="SUB_ID" value="<%=(rsAdminLinkEdit.Fields.Item("SUB_ID").Value)%>" size="32">
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">LINK_NAME:</font></td>
                              <td> 
                                <input type="text" name="LINK_NAME" value="<%=(rsAdminLinkEdit.Fields.Item("LINK_NAME").Value)%>" size="32">
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">LINK_URL:</font></td>
                              <td> 
                                <input type="text" name="LINK_URL" value="<%=(rsAdminLinkEdit.Fields.Item("LINK_URL").Value)%>" size="32">
                              </td>
                            </tr>
                            <tr> 
                              <td nowrap align="right" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">LINK_DESC:</font></td>
                              <td valign="baseline"> 
                                <textarea name="LINK_DESC" cols="50" rows="5"><%=(rsAdminLinkEdit.Fields.Item("LINK_DESC").Value)%></textarea>
                              </td>
                            </tr>
                            <tr valign="baseline"> 
                              <td nowrap align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font></td>
                              <td> 
                                <input type="submit" value="Update Link">
                              </td>
                            </tr>
                          </table>
                          <input type="hidden" name="MM_update" value="true">
                          <input type="hidden" name="MM_recordId" value="<%= rsAdminLinkEdit.Fields.Item("LINK_ID").Value %>">
                        </form>
                        <p>&nbsp;</p>
                      </td>
                    </tr>
                    <% End If ' end Not rsAdminLinkEdit.EOF Or NOT rsAdminLinkEdit.BOF %>
                  </table>
                </td>
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
rsAdminLink.Close()
%>
<%
rsCat.Close()
%>
<%
rsAdminLinkEdit.Close()
%>

