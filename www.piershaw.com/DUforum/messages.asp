<%@LANGUAGE="VBSCRIPT"%>
<% Response.Buffer = "true" %>
<!--#include file="../Connections/connDUportal.asp" -->
<%
Dim rsMessages__varID
rsMessages__varID = "9999"
if (Request.QueryString("for_id") <> "") then rsMessages__varID = Request.QueryString("for_id")
%>
<%
set rsMessages = Server.CreateObject("ADODB.Recordset")
rsMessages.ActiveConnection = MM_connDUportal_STRING
rsMessages.Source = "SELECT *, U_EMAIL  FROM MESSAGES INNER JOIN FORUMS ON FORUMS.FOR_ID = MESSAGES.FOR_ID, USERS  WHERE U_ID = MSG_AUTHOR AND MESSAGES.FOR_ID = " + Replace(rsMessages__varID, "'", "''") + "  ORDER BY MSG_LAST_POST DESC"
rsMessages.CursorType = 0
rsMessages.CursorLocation = 2
rsMessages.LockType = 3
rsMessages.Open()
rsMessages_numRows = 0
%>
<%
Dim rsPoster__MMColParam
rsPoster__MMColParam = "1"
if (Session("MM_Username") <> "") then rsPoster__MMColParam = Session("MM_Username")
%>
<%
set rsPoster = Server.CreateObject("ADODB.Recordset")
rsPoster.ActiveConnection = MM_connDUportal_STRING
rsPoster.Source = "SELECT * FROM USERS WHERE U_ID = '" + Replace(rsPoster__MMColParam, "'", "''") + "'"
rsPoster.CursorType = 0
rsPoster.CursorLocation = 2
rsPoster.LockType = 3
rsPoster.Open()
rsPoster_numRows = 0
%>
<%
Dim rsForum__MMColParam
rsForum__MMColParam = "1"
if (Request.QueryString("for_id") <> "") then rsForum__MMColParam = Request.QueryString("for_id")
%>
<%
set rsForum = Server.CreateObject("ADODB.Recordset")
rsForum.ActiveConnection = MM_connDUportal_STRING
rsForum.Source = "SELECT * FROM FORUMS WHERE FOR_ID = " + Replace(rsForum__MMColParam, "'", "''") + ""
rsForum.CursorType = 0
rsForum.CursorLocation = 2
rsForum.LockType = 3
rsForum.Open()
rsForum_numRows = 0
%>
<%
Dim RepeatLongTopicNew__numRows
RepeatLongTopicNew__numRows = 30
Dim RepeatLongTopicNew__index
RepeatLongTopicNew__index = 0
rsMessages_numRows = rsMessages_numRows + RepeatLongTopicNew__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
rsMessages_total = rsMessages.RecordCount

' set the number of rows displayed on this page
If (rsMessages_numRows < 0) Then
  rsMessages_numRows = rsMessages_total
Elseif (rsMessages_numRows = 0) Then
  rsMessages_numRows = 1
End If

' set the first and last displayed record
rsMessages_first = 1
rsMessages_last  = rsMessages_first + rsMessages_numRows - 1

' if we have the correct record count, check the other stats
If (rsMessages_total <> -1) Then
  If (rsMessages_first > rsMessages_total) Then rsMessages_first = rsMessages_total
  If (rsMessages_last > rsMessages_total) Then rsMessages_last = rsMessages_total
  If (rsMessages_numRows > rsMessages_total) Then rsMessages_numRows = rsMessages_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsMessages_total = -1) Then

  ' count the total records by iterating through the recordset
  rsMessages_total=0
  While (Not rsMessages.EOF)
    rsMessages_total = rsMessages_total + 1
    rsMessages.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsMessages.CursorType > 0) Then
    rsMessages.MoveFirst
  Else
    rsMessages.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsMessages_numRows < 0 Or rsMessages_numRows > rsMessages_total) Then
    rsMessages_numRows = rsMessages_total
  End If

  ' set the first and last displayed record
  rsMessages_first = 1
  rsMessages_last = rsMessages_first + rsMessages_numRows - 1
  If (rsMessages_first > rsMessages_total) Then rsMessages_first = rsMessages_total
  If (rsMessages_last > rsMessages_total) Then rsMessages_last = rsMessages_total

End If
%>
<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = rsMessages
MM_rsCount   = rsMessages_total
MM_size      = rsMessages_numRows
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
rsMessages_first = MM_offset + 1
rsMessages_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (rsMessages_first > MM_rsCount) Then rsMessages_first = MM_rsCount
  If (rsMessages_last > MM_rsCount) Then rsMessages_last = MM_rsCount
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
                  <% If Not rsMessages.EOF Or Not rsMessages.BOF Then %>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td align="left" valign="middle"  height="20" colspan="2"> 
                        <div class = "links">&nbsp;<b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="default.asp">MESSAGES 
                          BOARDS</a> &gt; <a href="messages.asp?for_id=<%=(rsMessages.Fields.Item("FOR_ID").Value)%>"><%= UCase((rsMessages.Fields.Item("FOR_NAME").Value)) %></a> :</font></b></div>
                      </td>
                    </tr>
                    <tr> 
                      <td align="left" valign="middle" colspan="2" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="middle" class = "bg_navigator" height="20"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp;MESSAGE 
                        LISTING</font></b></td>
                      <td align="right" valign="middle" class = "bg_navigator" height="20"> 
                        <font size="1"> <font face="Verdana, Arial, Helvetica, sans-serif"> 
                        MESSAGES 
                        <%
For i = 1 to rsMessages_total Step MM_size
TM_endCount = i + MM_size - 1
if TM_endCount > rsMessages_total Then TM_endCount = rsMessages_total
if i <> MM_offset + 1 Then
Response.Write("<a href=""" & Request.ServerVariables("URL") & "?" & MM_keepMove & "offset=" & i-1 & """>")
Response.Write(i & "-" & TM_endCount & "</a>")
else
Response.Write("<b>" & i & "-" & TM_endCount & "</b>")
End if
if(TM_endCount <> rsMessages_total) then Response.Write(" | ")
next
 %>
                        &nbsp; </font></font> </td>
                    </tr>
                    <tr> 
                      <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                    </tr>
                    <tr valign="top" align="center"> 
                      <td colspan="2" align="left"> 
                        <div class = "links"> 
                          <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#999999">
                            <tr align="center" valign="middle" class = "bg_login"> 
                              <td align="left" height="20"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Topic</font></b></td>
                              <td width="60"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Dated</font></b></td>
                              <td width="40"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Replies</font></b></td>
                              <td width="40"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Reads</font></b></td>
                              <td width="130"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Last 
                                Post</font></b></td>
                            </tr>
                            <% 
While ((RepeatLongTopicNew__numRows <> 0) AND (NOT rsMessages.EOF)) 
%>
                            <tr align="center" valign="middle" bgcolor="#FFFFFF"> 
                              <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><a href="msgDetail.asp?msg_id=<%=(rsMessages.Fields.Item("MSG_ID").Value)%>&for_id=<%=(rsMessages.Fields.Item("FOR_ID").Value)%>"><%=(rsMessages.Fields.Item("MSG_SUBJECT").Value)%></a></b> by <a href="mailto:<%=(rsMessages.Fields.Item("U_EMAIL").Value)%>"><%=(rsMessages.Fields.Item("MSG_AUTHOR").Value)%></a></font></td>
                              <td><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsMessages.Fields.Item("MSG_DATE").Value)%></font></td>
                              <td><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsMessages.Fields.Item("MSG_REPLY_COUNT").Value)%></font></td>
                              <td><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsMessages.Fields.Item("MSG_READ_COUNT").Value)%></font></td>
                              <td><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsMessages.Fields.Item("MSG_LAST_POST").Value)%></font></td>
                            </tr>
                            <% 
  RepeatLongTopicNew__index=RepeatLongTopicNew__index+1
  RepeatLongTopicNew__numRows=RepeatLongTopicNew__numRows-1
  rsMessages.MoveNext()
Wend
%>
                          </table>
                        </div>
                      </td>
                    </tr>
                  </table>
                  <% End If ' end Not rsMessages.EOF Or NOT rsMessages.BOF %>
                </td>
              </tr>
              <tr> 
                <td align="left" valign="top"> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="middle" class = "bg_navigator" height="20"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp;POST 
                        NEW TOPIC</font></b></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                    </tr>
                    <tr> 
                      <form name="POST" method="get" action="msgAdding.asp">
                        <td align="left" valign="top"> 
                          <table width="100%" border="0" cellspacing="5" cellpadding="5">
                            <tr align="left" valign="middle"> 
                              <td align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Poster:</font></b></td>
                              <td> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                                <% If Not rsPoster.EOF Or Not rsPoster.BOF Then %>
                                <input type="hidden" name="MSG_AUTHOR" value="<%=(rsPoster.Fields.Item("U_ID").Value)%>">
                                <%=(rsPoster.Fields.Item("U_ID").Value)%> 
                                <% Else %>
                                <font color = "ff0000"> To post a message, please 
                                login or register first.</font> 
                                <% End If ' end Not rsPoster.EOF Or NOT rsPoster.BOF %>
                                </font></b> </td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Forum:</font></b></td>
                              <td> 
                                <div class = "links"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                                  <input type="hidden" name="FOR_ID" value="<%=(rsForum.Fields.Item("FOR_ID").Value)%>">
                                  <a href="messages.asp?for_id=<%=(rsForum.Fields.Item("FOR_ID").Value)%>"><%=(rsForum.Fields.Item("FOR_NAME").Value)%></a></font></b></div>
                              </td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Subject:</font></b></td>
                              <td> 
                                <input type="text" name="MSG_SUBJECT" size="60" maxlength="60" class = "fields">
                              </td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td align="right" valign="top" rowspan="2"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Message:</font></b></td>
                              <td valign="top"><font size="1"><i><font face="Verdana, Arial, Helvetica, sans-serif">If 
                                your message contains HTML or ASP codes, please 
                                replace &lt; with [ and &gt; with ]. If not, your 
                                codes won't display correctly. </font> </i> </font></td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td> 
                                <textarea name="MSG_BODY" cols="60" rows="10" class = "fields"></textarea>
                              </td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font></b></td>
                              <td> 
                                <% If Not rsPoster.EOF Or Not rsPoster.BOF Then %>
                                <input type="submit" name="SUBMIT" value="POST" class = "buttons">
                                <% End If %>
                              </td>
                            </tr>
                          </table>
                        </td>
                      </form>
                    </tr>
                    <tr> 
                      <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
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
rsMessages.Close()
%>
<%
rsPoster.Close()
%>
<%
rsForum.Close()
%>

