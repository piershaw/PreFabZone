<%@LANGUAGE="VBSCRIPT"%>
<% Response.Buffer = "true" %>
<!--#include file="../Connections/connDUportal.asp" -->
<%
Dim rsNewsListing__varTYPE
rsNewsListing__varTYPE = "999"
if (Request.QueryString("type_id") <> "") then rsNewsListing__varTYPE = Request.QueryString("type_id")
%>
<%
set rsNewsListing = Server.CreateObject("ADODB.Recordset")
rsNewsListing.ActiveConnection = MM_connDUportal_STRING
rsNewsListing.Source = "SELECT *  FROM NEWS, NEWS_TYPES  WHERE NEWS_APPROVED = Yes AND NEWS_TYPE = TYPE_ID AND TYPE_ID = " + Replace(rsNewsListing__varTYPE, "'", "''") + "  ORDER BY NEWS_DATE DESC"
rsNewsListing.CursorType = 0
rsNewsListing.CursorLocation = 2
rsNewsListing.LockType = 3
rsNewsListing.Open()
rsNewsListing_numRows = 0
%>
<%
Dim RepeatNewsListing__numRows
RepeatNewsListing__numRows = 30
Dim RepeatNewsListing__index
RepeatNewsListing__index = 0
rsNewsListing_numRows = rsNewsListing_numRows + RepeatNewsListing__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
rsNewsListing_total = rsNewsListing.RecordCount

' set the number of rows displayed on this page
If (rsNewsListing_numRows < 0) Then
  rsNewsListing_numRows = rsNewsListing_total
Elseif (rsNewsListing_numRows = 0) Then
  rsNewsListing_numRows = 1
End If

' set the first and last displayed record
rsNewsListing_first = 1
rsNewsListing_last  = rsNewsListing_first + rsNewsListing_numRows - 1

' if we have the correct record count, check the other stats
If (rsNewsListing_total <> -1) Then
  If (rsNewsListing_first > rsNewsListing_total) Then rsNewsListing_first = rsNewsListing_total
  If (rsNewsListing_last > rsNewsListing_total) Then rsNewsListing_last = rsNewsListing_total
  If (rsNewsListing_numRows > rsNewsListing_total) Then rsNewsListing_numRows = rsNewsListing_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsNewsListing_total = -1) Then

  ' count the total records by iterating through the recordset
  rsNewsListing_total=0
  While (Not rsNewsListing.EOF)
    rsNewsListing_total = rsNewsListing_total + 1
    rsNewsListing.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsNewsListing.CursorType > 0) Then
    rsNewsListing.MoveFirst
  Else
    rsNewsListing.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsNewsListing_numRows < 0 Or rsNewsListing_numRows > rsNewsListing_total) Then
    rsNewsListing_numRows = rsNewsListing_total
  End If

  ' set the first and last displayed record
  rsNewsListing_first = 1
  rsNewsListing_last = rsNewsListing_first + rsNewsListing_numRows - 1
  If (rsNewsListing_first > rsNewsListing_total) Then rsNewsListing_first = rsNewsListing_total
  If (rsNewsListing_last > rsNewsListing_total) Then rsNewsListing_last = rsNewsListing_total

End If
%>
<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = rsNewsListing
MM_rsCount   = rsNewsListing_total
MM_size      = rsNewsListing_numRows
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
rsNewsListing_first = MM_offset + 1
rsNewsListing_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (rsNewsListing_first > MM_rsCount) Then rsNewsListing_first = MM_rsCount
  If (rsNewsListing_last > MM_rsCount) Then rsNewsListing_last = MM_rsCount
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
                          &gt; <a href="default.asp">NEWS</a> &gt; <A HREF="news.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "type_id=" & rsNewsListing.Fields.Item("TYPE_ID").Value %>"><%= UCase((rsNewsListing.Fields.Item("TYPE_NAME").Value)) %></A> :</b></font></td>
                      </tr>
                      <tr> 
                        <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                      </tr>
                      <tr> 
                        <td align="left" valign="top"> 
                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td align="left" valign="middle" class = "bg_navigator" height="20"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><font size="1"> 
                                &nbsp;NEWS LISTING FOR <%= UCase((rsNewsListing.Fields.Item("TYPE_NAME").Value)) %></font></font></b></td>
                              <td align="right" valign="middle" class = "bg_navigator" height="20"> 
                                <font size="1"> <b> <font face="Verdana, Arial, Helvetica, sans-serif">
                                SHOWING <%
For i = 1 to rsNewsListing_total Step MM_size
TM_endCount = i + MM_size - 1
if TM_endCount > rsNewsListing_total Then TM_endCount = rsNewsListing_total
if i <> MM_offset + 1 Then
Response.Write("<a href=""" & Request.ServerVariables("URL") & "?" & MM_keepMove & "offset=" & i-1 & """>")
Response.Write(i & "-" & TM_endCount & "</a>")
else
Response.Write("<b>" & i & "-" & TM_endCount & "</b>")
End if
if(TM_endCount <> rsNewsListing_total) then Response.Write(" | ")
next
 %>&nbsp;
                                </font></b></font> </td>
                            </tr>
                            <tr> 
                              <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                            </tr>
                            <tr> 
                              <td align="left" valign="top" colspan="2"> 
                                <% 
While ((RepeatNewsListing__numRows <> 0) AND (NOT rsNewsListing.EOF)) 
%>
                                <table width="100%" border="0" cellspacing="0" cellpadding="3">
                                  <tr> 
                                    <td align="left" valign="middle"> 
                                      <div class = "links"><img src="../assets/bullet.gif" width="11" height="11" align="absmiddle"> 
                                        <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font size="1"><a href="newsDetail.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "id=" & rsNewsListing.Fields.Item("NEWS_ID").Value %>"><%=(rsNewsListing.Fields.Item("NEWS_TITLE").Value)%></a></font></b></font> <font size="2"><i>(<%=(rsNewsListing.Fields.Item("NEWS_SOURCE").Value)%>)</i></font></div>
                                    </td>
                                    <td align="right" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                                      <b>Dated:</b> <%=(rsNewsListing.Fields.Item("NEWS_DATE").Value)%></font></td>
                                  </tr>
                                  <tr> 
                                    <td align="left" valign="top" colspan="2"> 
                                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                        <tr> 
                                          <td width="14">&nbsp;</td>
                                          <td align="left" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                                            <% =(DoTrimProperly((rsNewsListing.Fields.Item("NEWS_DESC").Value), 250, 1, 1, " ...")) %>
                                            </font></td>
                                        </tr>
                                      </table>
                                    </td>
                                  </tr>
                                </table>
                                <% 
  RepeatNewsListing__index=RepeatNewsListing__index+1
  RepeatNewsListing__numRows=RepeatNewsListing__numRows-1
  rsNewsListing.MoveNext()
Wend
%>
                              </td>
                            </tr>
                            <tr> 
                              <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
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
rsNewsListing.Close()
%>

