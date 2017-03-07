<!--#include file="../Connections/connDUportal.asp" -->
<%
set rsHotPoll = Server.CreateObject("ADODB.Recordset")
rsHotPoll.ActiveConnection = MM_connDUportal_STRING
rsHotPoll.Source = "SELECT * FROM QUESTIONS ORDER BY TOTAL_VOTES DESC"
rsHotPoll.CursorType = 0
rsHotPoll.CursorLocation = 2
rsHotPoll.LockType = 3
rsHotPoll.Open()
rsHotPoll_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = 25
Dim Repeat1__index
Repeat1__index = 0
rsHotPoll_numRows = rsHotPoll_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
rsHotPoll_total = rsHotPoll.RecordCount

' set the number of rows displayed on this page
If (rsHotPoll_numRows < 0) Then
  rsHotPoll_numRows = rsHotPoll_total
Elseif (rsHotPoll_numRows = 0) Then
  rsHotPoll_numRows = 1
End If

' set the first and last displayed record
rsHotPoll_first = 1
rsHotPoll_last  = rsHotPoll_first + rsHotPoll_numRows - 1

' if we have the correct record count, check the other stats
If (rsHotPoll_total <> -1) Then
  If (rsHotPoll_first > rsHotPoll_total) Then rsHotPoll_first = rsHotPoll_total
  If (rsHotPoll_last > rsHotPoll_total) Then rsHotPoll_last = rsHotPoll_total
  If (rsHotPoll_numRows > rsHotPoll_total) Then rsHotPoll_numRows = rsHotPoll_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsHotPoll_total = -1) Then

  ' count the total records by iterating through the recordset
  rsHotPoll_total=0
  While (Not rsHotPoll.EOF)
    rsHotPoll_total = rsHotPoll_total + 1
    rsHotPoll.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsHotPoll.CursorType > 0) Then
    rsHotPoll.MoveFirst
  Else
    rsHotPoll.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsHotPoll_numRows < 0 Or rsHotPoll_numRows > rsHotPoll_total) Then
    rsHotPoll_numRows = rsHotPoll_total
  End If

  ' set the first and last displayed record
  rsHotPoll_first = 1
  rsHotPoll_last = rsHotPoll_first + rsHotPoll_numRows - 1
  If (rsHotPoll_first > rsHotPoll_total) Then rsHotPoll_first = rsHotPoll_total
  If (rsHotPoll_last > rsHotPoll_total) Then rsHotPoll_last = rsHotPoll_total

End If
%>
<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = rsHotPoll
MM_rsCount   = rsHotPoll_total
MM_size      = rsHotPoll_numRows
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
rsHotPoll_first = MM_offset + 1
rsHotPoll_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (rsHotPoll_first > MM_rsCount) Then rsHotPoll_first = MM_rsCount
  If (rsHotPoll_last > MM_rsCount) Then rsHotPoll_last = MM_rsCount
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
<table width="100%" border="0" cellspacing="0" cellpadding="0"><div class = "links">
  <tr valign="middle" class = "bg_navigator"> 
    <td align="left" height="20">&nbsp;<font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font size="1">HOT 
      POLLS</font></b></font></td>
    <td align="right" valign="middle"> <font face="Verdana, Arial, Helvetica, sans-serif"> 
      <font size="1"><b>SHOWING POLLS 
      <%
For i = 1 to rsHotPoll_total Step MM_size
TM_endCount = i + MM_size - 1
if TM_endCount > rsHotPoll_total Then TM_endCount = rsHotPoll_total
if i <> MM_offset + 1 Then
Response.Write("<a href=""" & Request.ServerVariables("URL") & "?" & MM_keepMove & "offset=" & i-1 & """>")
Response.Write(i & "-" & TM_endCount & "</a>")
else
Response.Write("<b>" & i & "-" & TM_endCount & "</b>")
End if
if(TM_endCount <> rsHotPoll_total) then Response.Write(" | ")
next
 %>
      </b></font></font>&nbsp;</td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr> 
    <td align="left" valign="top" colspan="2"> 
      <% 
While ((Repeat1__numRows <> 0) AND (NOT rsHotPoll.EOF)) 
%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="left" valign="top"> 
            <table width="100%" border="0" cellspacing="3" cellpadding="3">
              <tr align="left" valign="middle"> 
                <td> <font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp;<img src="../assets/bullet.gif" width="11" height="11" align="absmiddle"> 
                  <b><A HREF="pollDetail.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "id=" & rsHotPoll.Fields.Item("QUEST_ID").Value %>"><%=(rsHotPoll.Fields.Item("QUESTION").Value)%></A></b> <i>(<%=(rsHotPoll.Fields.Item("QUEST_DATED").Value)%>)</i><br>
                  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=(rsHotPoll.Fields.Item("QUEST_DESCRIPTION").Value)%></font></td>
                <td width="120"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Total 
                  Votes:</b> <%=(rsHotPoll.Fields.Item("TOTAL_VOTES").Value)%></font></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
        </tr>
      </table>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsHotPoll.MoveNext()
Wend
%>
    </td>
  </tr></div>
</table>
<%
rsHotPoll.Close()
%>
