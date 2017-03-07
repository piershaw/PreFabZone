<%@LANGUAGE="VBSCRIPT"%>
<% Response.Buffer = "true" %>

<!--#include file="../Connections/connDUportal.asp" -->

<%
if(Request.QueryString("id") <> "") then cmdNewsRead__varID = Request.QueryString("id")
%>
<%
set cmdNewsRead = Server.CreateObject("ADODB.Command")
cmdNewsRead.ActiveConnection = MM_connDUportal_STRING
cmdNewsRead.CommandText = "UPDATE NEWS  SET NEWS_HITS = NEWS_HITS +1 WHERE NEWS_ID = " + Replace(cmdNewsRead__varID, "'", "''") + " "
cmdNewsRead.CommandType = 1
cmdNewsRead.CommandTimeout = 0
cmdNewsRead.Prepared = true
cmdNewsRead.Execute()
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
  MM_editTable = "COMMENTS"
  MM_editRedirectUrl = "newsDetail.asp"
  MM_fieldsStr  = "COM_HEADER|value|COM_COMMENT|value|COM_AUTHOR|value|RESOURCE_TYPE|value|RESOURCE_ID|value"
  MM_columnsStr = "COM_HEADER|',none,''|COM_COMMENT|',none,''|COM_AUTHOR|',none,''|RESOURCE_TYPE|',none,''|RESOURCE_ID|none,none,NULL"

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
Dim rsNewsDetail__varID
rsNewsDetail__varID = "999"
if (Request.QueryString("id")  <> "") then rsNewsDetail__varID = Request.QueryString("id") 
%>
<%
set rsNewsDetail = Server.CreateObject("ADODB.Recordset")
rsNewsDetail.ActiveConnection = MM_connDUportal_STRING
rsNewsDetail.Source = "SELECT *  FROM NEWS, NEWS_TYPES  WHERE NEWS_APPROVED = Yes AND NEWS_TYPE = TYPE_ID AND NEWS_ID = " + Replace(rsNewsDetail__varID, "'", "''") + ""
rsNewsDetail.CursorType = 0
rsNewsDetail.CursorLocation = 2
rsNewsDetail.LockType = 3
rsNewsDetail.Open()
rsNewsDetail_numRows = 0
%>
<%
Dim rsComment__MMColParam
rsComment__MMColParam = "1"
if (Request.QueryString("id") <> "") then rsComment__MMColParam = Request.QueryString("id")
%>
<%
set rsComment = Server.CreateObject("ADODB.Recordset")
rsComment.ActiveConnection = MM_connDUportal_STRING
rsComment.Source = "SELECT *  FROM COMMENTS  WHERE RESOURCE_ID = " + Replace(rsComment__MMColParam, "'", "''") + " AND RESOURCE_TYPE = 'NEWS'  ORDER BY COM_DATE DESC"
rsComment.CursorType = 0
rsComment.CursorLocation = 2
rsComment.LockType = 3
rsComment.Open()
rsComment_numRows = 0
%>
<%
Dim rsCOM_AUTHOR__MMColParam
rsCOM_AUTHOR__MMColParam = "1"
if (Session("MM_Username") <> "") then rsCOM_AUTHOR__MMColParam = Session("MM_Username")
%>
<%
set rsCOM_AUTHOR = Server.CreateObject("ADODB.Recordset")
rsCOM_AUTHOR.ActiveConnection = MM_connDUportal_STRING
rsCOM_AUTHOR.Source = "SELECT U_FIRST + ' ' + U_LAST AS AUTHOR  FROM USERS  WHERE U_ID = '" + Replace(rsCOM_AUTHOR__MMColParam, "'", "''") + "'"
rsCOM_AUTHOR.CursorType = 0
rsCOM_AUTHOR.CursorLocation = 2
rsCOM_AUTHOR.LockType = 3
rsCOM_AUTHOR.Open()
rsCOM_AUTHOR_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsComment_numRows = rsComment_numRows + Repeat1__numRows
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
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
                        <td align="left" valign="middle" height="20"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>&nbsp;<a href="../default.asp">HOME</a> 
                          &gt; <a href="default.asp">NEWS</a> &gt; <A HREF="news.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "type_id=" & rsNewsDetail.Fields.Item("TYPE_ID").Value %>"><%= UCase((rsNewsDetail.Fields.Item("TYPE_NAME").Value)) %></A> &gt; DETAIL VIEW:</b></font></td>
                      </tr>
                      <tr> 
                        <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                      </tr>
                      <tr> 
                        <td align="left" valign="top"> 
                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td align="left" valign="middle" class = "bg_navigator" height="20"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><font size="1"> 
                                &nbsp;<%= UCase((rsNewsDetail.Fields.Item("NEWS_TITLE").Value)) %></font></font></b> <i><font size="1" face="Verdana, Arial, Helvetica, sans-serif">(from 
                                <%=(rsNewsDetail.Fields.Item("NEWS_SOURCE").Value)%> on <%=(rsNewsDetail.Fields.Item("NEWS_DATE").Value)%> ) </font></i> </td>
                            </tr>
                            <tr> 
                              <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                            </tr>
                            <tr> 
                              <td align="left" valign="top"> 
                                <table width="100%" border="0" cellspacing="0" cellpadding="3">
                                  <tr> 
                                    <td align="left" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsNewsDetail.Fields.Item("NEWS_DESC").Value)%></font></td>
                                  </tr>
                                  <tr> 
                                    <td align="right" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><a href="<%=(rsNewsDetail.Fields.Item("NEWS_URL").Value)%>" target="_blank">&gt;&gt;&gt; 
                                      READ MORE</a></b></font></td>
                                  </tr>
                                </table>
                              </td>
                            </tr>
                            <tr> 
                              <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                            </tr>
                            <tr> 
                              <td align="left" valign="middle" height="20" class = "bg_navigator"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>&nbsp;NEWS 
                                REVIEWING</b></font></td>
                            </tr>
                            <tr> 
                              <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                            <tr> 
                              <td align="left" valign="top"> 
                                <% If Not rsComment.EOF Or Not rsComment.BOF Then %>
                                <% 
While ((Repeat1__numRows <> 0) AND (NOT rsComment.EOF)) 
%>
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td align="left" valign="middle" bgcolor="#CCCCCC" height="20"><font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp;<b><%=(rsComment.Fields.Item("COM_HEADER").Value)%></b> <i>(<%=(rsComment.Fields.Item("COM_AUTHOR").Value)%> - <%=(rsComment.Fields.Item("COM_DATE").Value)%>)</i></font></td>
                                  </tr>
                                  <tr> 
                                    <td align="left" valign="top"> 
                                      <table width="100%" border="0" cellspacing="2" cellpadding="3">
                                        <tr> 
                                          <td align="left" valign="top">&nbsp;<font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsComment.Fields.Item("COM_COMMENT").Value)%></font></td>
                                        </tr>
                                      </table>
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                                  </tr>
                                </table>
                                <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsComment.MoveNext()
Wend
%>
                                <% End If ' end Not rsComment.EOF Or NOT rsComment.BOF %>
                              </td>
                            </tr>
                            <tr> 
                              <td align="right" valign="top"> 
                                <table border="0" cellspacing="2" cellpadding="5">
                                  <form name="COMMENTS" method="POST" action="<%=MM_editAction%>">
                                    <tr align="left" valign="middle"> 
                                      <td align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">SUBJECT</font></b></td>
                                      <td> 
                                        <input type="text" name="COM_HEADER" size="51" class = "fields">
                                      </td>
                                    </tr>
                                    <tr align="left" valign="middle"> 
                                      <td align="right" valign="top"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">MESSAGES</font></b></td>
                                      <td> 
                                        <textarea name="COM_COMMENT" cols="50" rows="4" class = "fields"></textarea>
                                      </td>
                                    </tr>
                                    <% If Not rsCOM_AUTHOR.EOF Or Not rsCOM_AUTHOR.BOF Then %>
                                    <tr align="left" valign="middle"> 
                                      <td align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"></font></b></td>
                                      <td> 
                                        <input type="submit" name="Submit3" value="Submit" class = "buttons">
                                        <input type="hidden" name="COM_AUTHOR" value="<%=(rsCOM_AUTHOR.Fields.Item("AUTHOR").Value)%>">
                                        <input type="hidden" name="RESOURCE_TYPE" value="NEWS">
                                        <input type="hidden" name="RESOURCE_ID" value="<%=Request.QueryString("id")%>">
                                      </td>
                                    </tr>
                                    <% Else %>
                                    <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color = "ff0000">To 
                                    comment this link, please <a href="../DUhome/login.asp">login</a> 
                                    or <a href="../DUhome/register.asp">register</a> 
                                    first</font> 
                                    <% End If ' end Not rsCOM_AUTHOR.EOF Or NOT rsCOM_AUTHOR.BOF %>
                                    <input type="hidden" name="MM_insert" value="true">
                                  </form>
                                </table>
                              </td>
                            </tr>
							<tr> 
                                    <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                                  </tr>
                          </table>
                        </td>
                      </tr>
                      <tr> 
                        <td align="left" valign="top"> 
                          <!--#include file="../DUnews/inc_news_hot.asp" -->
                          <!--#include file="../DUnews/inc_news_new.asp" -->
                        </td>
                      </tr>
                    </table>
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
rsNewsDetail.Close()
%>
<%
rsComment.Close()
%>
<%
rsCOM_AUTHOR.Close()
%>
