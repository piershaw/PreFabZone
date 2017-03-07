<!--#include file="../Connections/connDUportal.asp" -->
<%
set rsUsers = Server.CreateObject("ADODB.Recordset")
rsUsers.ActiveConnection = MM_connDUportal_STRING
rsUsers.Source = "SELECT * FROM USERS ORDER BY U_ID ASC"
rsUsers.CursorType = 0
rsUsers.CursorLocation = 2
rsUsers.LockType = 3
rsUsers.Open()
rsUsers_numRows = 0
%>
<%
Dim HLooper1__numRows
HLooper1__numRows = 50
Dim HLooper1__index
HLooper1__index = 0
rsUsers_numRows = rsUsers_numRows + HLooper1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
rsUsers_total = rsUsers.RecordCount

' set the number of rows displayed on this page
If (rsUsers_numRows < 0) Then
  rsUsers_numRows = rsUsers_total
Elseif (rsUsers_numRows = 0) Then
  rsUsers_numRows = 1
End If

' set the first and last displayed record
rsUsers_first = 1
rsUsers_last  = rsUsers_first + rsUsers_numRows - 1

' if we have the correct record count, check the other stats
If (rsUsers_total <> -1) Then
  If (rsUsers_first > rsUsers_total) Then rsUsers_first = rsUsers_total
  If (rsUsers_last > rsUsers_total) Then rsUsers_last = rsUsers_total
  If (rsUsers_numRows > rsUsers_total) Then rsUsers_numRows = rsUsers_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsUsers_total = -1) Then

  ' count the total records by iterating through the recordset
  rsUsers_total=0
  While (Not rsUsers.EOF)
    rsUsers_total = rsUsers_total + 1
    rsUsers.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsUsers.CursorType > 0) Then
    rsUsers.MoveFirst
  Else
    rsUsers.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsUsers_numRows < 0 Or rsUsers_numRows > rsUsers_total) Then
    rsUsers_numRows = rsUsers_total
  End If

  ' set the first and last displayed record
  rsUsers_first = 1
  rsUsers_last = rsUsers_first + rsUsers_numRows - 1
  If (rsUsers_first > rsUsers_total) Then rsUsers_first = rsUsers_total
  If (rsUsers_last > rsUsers_total) Then rsUsers_last = rsUsers_total

End If
%>
<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = rsUsers
MM_rsCount   = rsUsers_total
MM_size      = rsUsers_numRows
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
rsUsers_first = MM_offset + 1
rsUsers_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (rsUsers_first > MM_rsCount) Then rsUsers_first = MM_rsCount
  If (rsUsers_last > MM_rsCount) Then rsUsers_last = MM_rsCount
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
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
<link rel="stylesheet" href="../css/default.css" type="text/css">
<div class = "links">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
      <td align="left" valign="middle" class = "bg_navigator" height="20"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>&nbsp;USERS 
        DIRECTORY </b> (Page 
        <%
TM_counter = 0
For i = 1 to rsUsers_total Step MM_size
TM_counter = TM_counter + 1
TM_PageEndCount = i + MM_size - 1
if TM_PageEndCount > rsUsers_total Then TM_PageEndCount = rsUsers_total
if i <> MM_offset + 1 then
Response.Write("<a href=""" & Request.ServerVariables("URL") & "?" & MM_keepMove & "offset=" & i-1 & """>")
Response.Write(TM_counter & "</a>")
else
Response.Write("<b>" & TM_counter & "</b>")
End if
if(TM_PageEndCount <> rsUsers_total) then Response.Write(" | ")
next
 %>
      ) </font></td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr> 
    <td align="center" valign="top"> 
      <table>
        <%
startrw = 0
endrw = HLooper1__index
numberColumns = 2
numrows = 25
while((numrows <> 0) AND (Not rsUsers.EOF))
	startrw = endrw + 1
	endrw = endrw + numberColumns
 %>
        <tr align="center" valign="top"> 
          <%
While ((startrw <= endrw) AND (Not rsUsers.EOF))
%>
          <td align="left"> 
            <table border="0" cellspacing="2" cellpadding="3">
              <tr> 
                <td width="11" align="center" valign="middle"><font color="#666666"><i><font size="2"><img src="../assets/bullet.gif" width="11" height="11" align="absmiddle"></font></i></font></td>
                <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color = "ff0000"><b><%=(rsUsers.Fields.Item("U_ID").Value)%></b></font> <font face="Verdana, Arial, Helvetica, sans-serif" size="2">(<%=(rsUsers.Fields.Item("U_FIRST").Value)%>&nbsp;<%=(rsUsers.Fields.Item("U_LAST").Value)%> - <a href="mailto:<%=(rsUsers.Fields.Item("U_EMAIL").Value)%>"><%=(rsUsers.Fields.Item("U_EMAIL").Value)%></a>)</font></td>
              </tr>
            </table>
          </td>
          <%
	startrw = startrw + 1
	rsUsers.MoveNext()
	Wend
	%>
        </tr>
        <%
 numrows=numrows-1
 Wend
 %>
      </table>
    </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
</table>
</div>
<%
rsUsers.Close()
%>
