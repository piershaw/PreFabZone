<%@LANGUAGE="VBSCRIPT"%>
<% Response.Buffer = "True" %>
<!--#include file="../Connections/connDUportal.asp" -->
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
  MM_editRedirectUrl = "dirRate.asp"
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
<%
Dim rsLink__ID
rsLink__ID = "1"
if (Request.QueryString("id") <> "") then rsLink__ID = Request.QueryString("id")
%>
<%
set rsLink = Server.CreateObject("ADODB.Recordset")
rsLink.ActiveConnection = MM_connDUportal_STRING
rsLink.Source = "SELECT *, (LINK_RATE/NO_RATES) AS RATING  FROM (LINKS INNER JOIN LINK_CATS ON LINKS.CAT_ID = LINK_CATS.CAT_ID) INNER JOIN LINK_SUBS ON LINKS.SUB_ID = LINK_SUBS.SUB_ID  WHERE LINK_ID = " + Replace(rsLink__ID, "'", "''") + ""
rsLink.CursorType = 0
rsLink.CursorLocation = 2
rsLink.LockType = 3
rsLink.Open()
rsLink_numRows = 0
%>
<%
Dim rsComment__MMColParam
rsComment__MMColParam = "1"
if (Request.QueryString("id") <> "") then rsComment__MMColParam = Request.QueryString("id")
%>
<%
set rsComment = Server.CreateObject("ADODB.Recordset")
rsComment.ActiveConnection = MM_connDUportal_STRING
rsComment.Source = "SELECT *  FROM COMMENTS  WHERE RESOURCE_ID = " + Replace(rsComment__MMColParam, "'", "''") + " AND RESOURCE_TYPE = 'LINKS'  ORDER BY COM_DATE DESC"
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
Repeat1__numRows = 15
Dim Repeat1__index
Repeat1__index = 0
rsComment_numRows = rsComment_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
rsComment_total = rsComment.RecordCount

' set the number of rows displayed on this page
If (rsComment_numRows < 0) Then
  rsComment_numRows = rsComment_total
Elseif (rsComment_numRows = 0) Then
  rsComment_numRows = 1
End If

' set the first and last displayed record
rsComment_first = 1
rsComment_last  = rsComment_first + rsComment_numRows - 1

' if we have the correct record count, check the other stats
If (rsComment_total <> -1) Then
  If (rsComment_first > rsComment_total) Then rsComment_first = rsComment_total
  If (rsComment_last > rsComment_total) Then rsComment_last = rsComment_total
  If (rsComment_numRows > rsComment_total) Then rsComment_numRows = rsComment_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsComment_total = -1) Then

  ' count the total records by iterating through the recordset
  rsComment_total=0
  While (Not rsComment.EOF)
    rsComment_total = rsComment_total + 1
    rsComment.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsComment.CursorType > 0) Then
    rsComment.MoveFirst
  Else
    rsComment.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsComment_numRows < 0 Or rsComment_numRows > rsComment_total) Then
    rsComment_numRows = rsComment_total
  End If

  ' set the first and last displayed record
  rsComment_first = 1
  rsComment_last = rsComment_first + rsComment_numRows - 1
  If (rsComment_first > rsComment_total) Then rsComment_first = rsComment_total
  If (rsComment_last > rsComment_total) Then rsComment_last = rsComment_total

End If
%>
<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = rsComment
MM_rsCount   = rsComment_total
MM_size      = rsComment_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  r = Request.QueryString("index")
  If r = "" Then r = Request.QueryString("offset")
  If r <> "" Then MM_offset = Int(r)

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  i = 0
  While ((Not MM_rs.EOF) And (i < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    i = i + 1
  Wend
  If (MM_rs.EOF) Then MM_offset = i  ' set MM_offset to the last possible record

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  i = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or i < MM_offset + MM_size))
    MM_rs.MoveNext
    i = i + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = i
    If (MM_size < 0 Or MM_size > MM_rsCount) Then MM_size = MM_rsCount
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  i = 0
  While (Not MM_rs.EOF And i < MM_offset)
    MM_rs.MoveNext
    i = i + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsComment_first = MM_offset + 1
rsComment_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (rsComment_first > MM_rsCount) Then rsComment_first = MM_rsCount
  If (rsComment_last > MM_rsCount) Then rsComment_last = MM_rsCount
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 0) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    params = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For i = 0 To UBound(params)
      nextItem = Left(params(i), InStr(params(i),"=") - 1)
      If (StrComp(nextItem,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & params(i)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then MM_keepMove = MM_keepMove & "&"
urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="
MM_moveFirst = urlStr & "0"
MM_moveLast  = urlStr & "-1"
MM_moveNext  = urlStr & Cstr(MM_offset + MM_size)
prev = MM_offset - MM_size
If (prev < 0) Then prev = 0
MM_movePrev  = urlStr & Cstr(prev)
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
                              <td align="left" valign="middle" height="20" class = "bg_navigator" colspan="2"><div class = "login">&nbsp;<b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="../default.asp"><font size="1">HOME</font></a> 
                                <font size="1"> &gt; <a href="default.asp">DIRECTORY</a> 
                                &gt; <a href="dirCat.asp?id=<%=(rsLink.Fields.Item("CAT_ID").Value)%>"><%= UCase((rsLink.Fields.Item("CAT_NAME").Value)) %></a> &gt; <a href="dirSub.asp?catid=<%=(rsLink.Fields.Item("CAT_ID").Value)%>&subid=<%=(rsLink.Fields.Item("SUB_ID").Value)%>"><%= UCase((rsLink.Fields.Item("SUB_NAME").Value)) %></a> &gt; LINK RATING:</font></font></b></div></td>
                          </tr>
                          <tr> 
                            <td align="left" valign="middle" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                          </tr>
                          
                      <tr> 
                        <td align="left" valign="top" colspan="2"> 
                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td align="left" valign="top" bgcolor="#FFFFFF"> 
                                <table width="100%" border="0" cellspacing="0" cellpadding="3">
                                  <tr valign="middle" bgcolor="#CCCCCC"> 
                                    <td align="left" height="20"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="../DUdirectory/dirHitting.asp?id=<%=(rsLink.Fields.Item("LINK_ID").Value)%>&url=<%=(rsLink.Fields.Item("LINK_URL").Value)%>" target="_blank" onClick="window.location.reload(true);"><b><%= UCase((rsLink.Fields.Item("LINK_NAME").Value)) %></b></a> <i>(<%=(rsLink.Fields.Item("LINK_URL").Value)%>)&nbsp;<font color="#FF0000"> 
                                      <% If rsLink.Fields.Item("RATING").Value > 4.0 Then %>
                                      <b>HOT!</b> 
                                      <% End If %>
                                      </font>&nbsp;<font color="#0000FF"> 
                                      <% If rsLink.Fields.Item("NO_HITS").Value > 50 Then %>
                                      <b><font color="#FF6633">POPULAR!</font></b> 
                                      <% End If %>
                                      </font>&nbsp;<font color="#0000FF"> 
                                      <% If rsLink.Fields.Item("LINK_DATE").Value > date() - 7 Then %>
                                      <b><font color="#FF00FF">NEW!</font></b> 
                                      <% End If %>
                                      </font></i></font></td>
                                    <td align="right" height="20"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsLink.Fields.Item("LINK_DATE").Value)%></font></td>
                                  </tr>
                                  <tr valign="middle"> 
                                    <td align="left" height="25"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Hits: 
                                      <%=(rsLink.Fields.Item("NO_HITS").Value)%> </b> | <b>Rating: </b> <img src="../assets/<%= FormatNumber((rsLink.Fields.Item("RATING").Value), 1, -2, -2, -2) %>.gif" align="absmiddle"> 
                                      <b>(<%=(rsLink.Fields.Item("NO_RATES").Value)%>)</b> </font></td>
                                    <td align="right" height="25"><b></b></td>
                                  </tr>
                                  <tr> 
                                    <td align="left" valign="top" colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Description:</b> 
                                      <%=(rsLink.Fields.Item("LINK_DESC").Value)%> </font></td>
                                  </tr>
                                  <tr align="right"> 
                                    <td valign="top" colspan="2"> 
                                      <% If Not rsCOM_AUTHOR.EOF Or Not rsCOM_AUTHOR.BOF Then %>
                                      <table border="0" cellspacing="1" cellpadding="0" bgcolor="#FFFFFF">
                                        <tr> 
                                          <form name="RATE" method="get" action="../DUdirectory/dirRating.asp">
                                            <td align="right" valign="top" bgcolor="#FFFFFF"> 
                                              <table border="0" cellspacing="2" cellpadding="2">
                                                <tr> 
                                                  <td align="center" valign="middle"><img src="../assets/star.gif" width="13" height="12" align="absmiddle"></td>
                                                  <td align="center" valign="middle"><img src="../assets/star.gif" width="13" height="12" align="absmiddle"></td>
                                                  <td align="center" valign="middle"><img src="../assets/star.gif" width="13" height="12" align="absmiddle"></td>
                                                  <td align="center" valign="middle"><img src="../assets/star.gif" width="13" height="12" align="absmiddle"></td>
                                                  <td align="center" valign="middle"><img src="../assets/star.gif" width="13" height="12" align="absmiddle"></td>
                                                  <td rowspan="2" align="center" valign="middle"> 
                                                    <input type="hidden" name="id" value="<%=(rsLink.Fields.Item("LINK_ID").Value)%>">
                                                    <input type="hidden" name="catid" value="<%=(rsLink.Fields.Item("CAT_ID").Value)%>">
                                                    <input type="submit" name="Submit2" value="Rate" class = "buttons">
                                                  </td>
                                                </tr>
                                                <tr> 
                                                  <td align="center" valign="middle"> 
                                                    <input type="radio" name="rate_value" value="1">
                                                  </td>
                                                  <td align="center" valign="middle"> 
                                                    <input type="radio" name="rate_value" value="2">
                                                  </td>
                                                  <td align="center" valign="middle"> 
                                                    <input type="radio" name="rate_value" value="3">
                                                  </td>
                                                  <td align="center" valign="middle"> 
                                                    <input type="radio" name="rate_value" value="4">
                                                  </td>
                                                  <td align="center" valign="middle"> 
                                                    <input type="radio" name="rate_value" value="5" checked>
                                                  </td>
                                                </tr>
                                              </table>
                                            </td>
                                          </form>
                                        </tr>
                                      </table><% Else %>
                                      <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color = "ff0000">To 
                                      rate this link, please <a href="../DUhome/login.asp">login</a> 
                                      or <a href="../DUhome/register.asp">register</a> 
                                      first</font> 
                                      <% End If ' end Not rsCOM_AUTHOR.EOF Or NOT rsCOM_AUTHOR.BOF %>
                                    </td>
                                  </tr>
                                </table>
                              </td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                      <tr> 
                        <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                      </tr>
                      <tr> 
                        <td align="left" valign="middle" class = "bg_navigator" height="20">&nbsp;<b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">LINK 
                          REVIEWING</font></b></td>
                        <td align="right" valign="middle" class = "bg_navigator" height="20"> 
                          <font face="Verdana, Arial, Helvetica, sans-serif"> 
                          <font size="1"> <b> 
                          <%
For i = 1 to rsComment_total Step MM_size
TM_endCount = i + MM_size - 1
if TM_endCount > rsComment_total Then TM_endCount = rsComment_total
if i <> MM_offset + 1 Then
Response.Write("<a href=""" & Request.ServerVariables("URL") & "?" & MM_keepMove & "offset=" & i-1 & """>")
Response.Write(i & "-" & TM_endCount & "</a>")
else
Response.Write("<b>" & i & "-" & TM_endCount & "</b>")
End if
if(TM_endCount <> rsComment_total) then Response.Write(" | ")
next
 %>
                          </b></font></font>&nbsp;</td>
                      </tr>
                      <tr> 
                        <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                      </tr>
                      <tr align="right"> 
                        <td valign="top" colspan="2"> 
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
                                  <input type="hidden" name="RESOURCE_TYPE" value="LINKS">
                                  <input type="hidden" name="RESOURCE_ID" value="<%=(rsLink.Fields.Item("LINK_ID").Value)%>">
                                </td>
                              </tr><% Else %>
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
                          <tr bgcolor="#000000"> 
                            <td align="left" valign="top" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                          </tr>
                        </table>
                      </div></td>
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
rsLink.Close()
%>
<%
rsComment.Close()
%>
<%
rsCOM_AUTHOR.Close()
%>
