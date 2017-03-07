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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then

  MM_editConnection = MM_connDUportal_STRING
  MM_editTable = "NEWS"
  MM_editRedirectUrl = "newsAdded.asp"
  MM_fieldsStr  = "NEWS_TYPE|value|NEWS_TITLE|value|NEWS_DESC|value|NEWS_URL|value|NEWS_SOURCE|value|NEWS_ADDER|value"
  MM_columnsStr = "NEWS_TYPE|none,none,NULL|NEWS_TITLE|',none,''|NEWS_DESC|',none,''|NEWS_URL|',none,''|NEWS_SOURCE|',none,''|NEWS_ADDER|',none,''"

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
set rsNewsType = Server.CreateObject("ADODB.Recordset")
rsNewsType.ActiveConnection = MM_connDUportal_STRING
rsNewsType.Source = "SELECT * FROM NEWS_TYPES ORDER BY TYPE_NAME ASC"
rsNewsType.CursorType = 0
rsNewsType.CursorLocation = 2
rsNewsType.LockType = 3
rsNewsType.Open()
rsNewsType_numRows = 0
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
                  <div class = "links"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td align="left" valign="middle" height="20" class = "bg_navigator"><font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp;<b>ADD 
                          NEWS</b></font></td>
                      </tr>
                      <tr> 
                        <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                      </tr>
                      <tr> 
                        <td align="left" valign="top"> 
                          <form method="POST" action="<%=MM_editAction%>" name="form1">
                            <table align="center" cellpadding="5" cellspacing="5">
                              <tr valign="baseline"> 
                                <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">NEWS 
                                  TYPE:</font></b></td>
                                <td> 
                                  <select name="NEWS_TYPE">
                                    <%
While (NOT rsNewsType.EOF)
%>
                                    <option value="<%=(rsNewsType.Fields.Item("TYPE_ID").Value)%>" <%if (CStr(rsNewsType.Fields.Item("TYPE_ID").Value) = CStr(rsNewsType.Fields.Item("TYPE_NAME").Value)) then Response.Write("SELECTED") : Response.Write("")%> ><%=(rsNewsType.Fields.Item("TYPE_NAME").Value)%></option>
                                    <%
  rsNewsType.MoveNext()
Wend
If (rsNewsType.CursorType > 0) Then
  rsNewsType.MoveFirst
Else
  rsNewsType.Requery
End If
%>
                                  </select>
                                </td>
                              </tr>
                              <tr valign="baseline"> 
                                <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">NEWS 
                                  TITLE:</font></b></td>
                                <td> 
                                  <input type="text" name="NEWS_TITLE" value="" size="45" maxlength="45">
                                </td>
                              </tr>
                              <tr valign="baseline"> 
                                <td nowrap align="right" valign="top"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">NEWS 
                                  DETAIL:</font></b></td>
                                <td> 
                                  <textarea name="NEWS_DESC" cols="60" rows="5"></textarea>
                                </td>
                              </tr>
                              <tr valign="baseline"> 
                                <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">NEWS 
                                  URL:</font></b></td>
                                <td> 
                                  <input type="text" name="NEWS_URL" value="" size="60">
                                </td>
                              </tr>
                              <tr valign="baseline"> 
                                <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">NEWS 
                                  SOURCE:</font></b></td>
                                <td> 
                                  <input type="text" name="NEWS_SOURCE" value="" size="40" maxlength="40">
                                </td>
                              </tr>
                              <tr valign="baseline"> 
                                <td nowrap align="right"> 
                                  <input type="hidden" name="NEWS_ADDER" value="<%= Session("MM_Username") %>">
                                </td>
                                <td> 
                                  <input type="submit" value="Add News" name="submit">
                                </td>
                              </tr>
                            </table>
                            <input type="hidden" name="MM_insert" value="true">
                          </form>
                        </td>
                      </tr>
                    </table>
                    <p>&nbsp;</p>
                  </div>
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
rsNewsType.Close()
%>

