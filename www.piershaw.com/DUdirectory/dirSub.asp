<%@LANGUAGE="VBSCRIPT"%>
<% Response.Buffer = "True" %>
<!--#include file="../Connections/connDUportal.asp" -->
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
Dim rsSub__varCat
rsSub__varCat = "4"
if (Request.QueryString("catid")  <> "") then rsSub__varCat = Request.QueryString("catid") 
%>
<%
Dim rsSub__varSub
rsSub__varSub = "1"
if (Request.QueryString("subid")  <> "") then rsSub__varSub = Request.QueryString("subid") 
%>
<%
set rsSub = Server.CreateObject("ADODB.Recordset")
rsSub.ActiveConnection = MM_connDUportal_STRING
rsSub.Source = "SELECT *, (SELECT COUNT (*) FROM LINKS WHERE LINKS.SUB_ID = LINK_SUBS.SUB_ID) AS LINK_COUNT  FROM LINK_SUBS  WHERE CAT_ID = " + Replace(rsSub__varCat, "'", "''") + " AND SUB_ID = " + Replace(rsSub__varSub, "'", "''") + "  ORDER BY SUB_NAME ASC"
rsSub.CursorType = 0
rsSub.CursorLocation = 2
rsSub.LockType = 3
rsSub.Open()
rsSub_numRows = 0
%>
<%
Dim rsCat__MMColParam
rsCat__MMColParam = "1"
if (Request.QueryString("catid")  <> "") then rsCat__MMColParam = Request.QueryString("catid") 
%>
<%
set rsCat = Server.CreateObject("ADODB.Recordset")
rsCat.ActiveConnection = MM_connDUportal_STRING
rsCat.Source = "SELECT *, (SELECT COUNT (*) FROM LINKS WHERE LINKS.CAT_ID =  LINK_CATS.CAT_ID) AS LINK_COUNT  FROM LINK_CATS  WHERE CAT_ID = " + Replace(rsCat__MMColParam, "'", "''") + ""
rsCat.CursorType = 0
rsCat.CursorLocation = 2
rsCat.LockType = 3
rsCat.Open()
rsCat_numRows = 0
%>
<%
Dim rsFeat__varCat
rsFeat__varCat = "1"
if (Request.QueryString("catid")  <> "") then rsFeat__varCat = Request.QueryString("catid") 
%>
<%
Dim rsFeat__varSub
rsFeat__varSub = "1"
if (Request.QueryString("subid")  <> "") then rsFeat__varSub = Request.QueryString("subid") 
%>
<%
set rsFeat = Server.CreateObject("ADODB.Recordset")
rsFeat.ActiveConnection = MM_connDUportal_STRING
rsFeat.Source = "SELECT *, (LINK_RATE/NO_RATES) AS RATING, (SELECT COUNT(*) FROM COMMENTS WHERE RESOURCE_ID = LINK_ID AND RESOURCE_TYPE = 'LINKS') AS COMMENTS  FROM LINKS  WHERE SUB_ID = " + Replace(rsFeat__varSub, "'", "''") + " AND CAT_ID = " + Replace(rsFeat__varCat, "'", "''") + "  AND LINK_APPROVED = Yes  ORDER BY LINK_RATE DESC"
rsFeat.CursorType = 0
rsFeat.CursorLocation = 2
rsFeat.LockType = 3
rsFeat.Open()
rsFeat_numRows = 0
%>
<%
Dim rsLink__varCat
rsLink__varCat = "1"
if (Request.QueryString("catid")  <> "") then rsLink__varCat = Request.QueryString("catid") 
%>
<%
Dim rsLink__varSub
rsLink__varSub = "1"
if (Request.QueryString("subid")  <> "") then rsLink__varSub = Request.QueryString("subid") 
%>
<%
set rsLink = Server.CreateObject("ADODB.Recordset")
rsLink.ActiveConnection = MM_connDUportal_STRING
rsLink.Source = "SELECT *, (LINK_RATE/NO_RATES) AS RATING, (SELECT COUNT(*) FROM COMMENTS WHERE RESOURCE_ID = LINK_ID AND RESOURCE_TYPE = 'LINKS') AS COMMENTS  FROM LINKS  WHERE SUB_ID =" + Replace(rsLink__varSub, "'", "''") + " AND CAT_ID = " + Replace(rsLink__varCat, "'", "''") + "  AND LINK_APPROVED = Yes  ORDER BY LINK_NAME ASC"
rsLink.CursorType = 0
rsLink.CursorLocation = 2
rsLink.LockType = 3
rsLink.Open()
rsLink_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = 30
Dim Repeat1__index
Repeat1__index = 0
rsLink_numRows = rsLink_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Repeat2__numRows = 3
Dim Repeat2__index
Repeat2__index = 0
rsFeat_numRows = rsFeat_numRows + Repeat2__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
rsLink_total = rsLink.RecordCount

' set the number of rows displayed on this page
If (rsLink_numRows < 0) Then
  rsLink_numRows = rsLink_total
Elseif (rsLink_numRows = 0) Then
  rsLink_numRows = 1
End If

' set the first and last displayed record
rsLink_first = 1
rsLink_last  = rsLink_first + rsLink_numRows - 1

' if we have the correct record count, check the other stats
If (rsLink_total <> -1) Then
  If (rsLink_first > rsLink_total) Then rsLink_first = rsLink_total
  If (rsLink_last > rsLink_total) Then rsLink_last = rsLink_total
  If (rsLink_numRows > rsLink_total) Then rsLink_numRows = rsLink_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsLink_total = -1) Then

  ' count the total records by iterating through the recordset
  rsLink_total=0
  While (Not rsLink.EOF)
    rsLink_total = rsLink_total + 1
    rsLink.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsLink.CursorType > 0) Then
    rsLink.MoveFirst
  Else
    rsLink.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsLink_numRows < 0 Or rsLink_numRows > rsLink_total) Then
    rsLink_numRows = rsLink_total
  End If

  ' set the first and last displayed record
  rsLink_first = 1
  rsLink_last = rsLink_first + rsLink_numRows - 1
  If (rsLink_first > rsLink_total) Then rsLink_first = rsLink_total
  If (rsLink_last > rsLink_total) Then rsLink_last = rsLink_total

End If
%>
<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = rsLink
MM_rsCount   = rsLink_total
MM_size      = rsLink_numRows
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
rsLink_first = MM_offset + 1
rsLink_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (rsLink_first > MM_rsCount) Then rsLink_first = MM_rsCount
  If (rsLink_last > MM_rsCount) Then rsLink_last = MM_rsCount
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
                        <td align="left" valign="middle" height="20" colspan="2"> 
                          <div class = "links"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>&nbsp;<a href="../default.asp"><font size="1">HOME</font></a><font size="1"> 
                            &gt; <a href="default.asp">DIRECTORY</a> &gt; <a href="dirCat.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "id=" & rsCat.Fields.Item("CAT_ID").Value %>"><%= UCase((rsCat.Fields.Item("CAT_NAME").Value)) %></a> (<%=(rsCat.Fields.Item("LINK_COUNT").Value)%>) &gt; <a href="dirSub.asp?catid=<%=(rsSub.Fields.Item("CAT_ID").Value)%>&subid=<%=(rsSub.Fields.Item("SUB_ID").Value)%>"><%= UCase((rsSub.Fields.Item("SUB_NAME").Value))%></a> (<%=(rsSub.Fields.Item("LINK_COUNT").Value)%>) :</font></b></font></div>
                        </td>
                      </tr>
                      <tr> 
                        <td align="left" valign="middle" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                      </tr>
                      <tr> 
                        <td align="left" valign="middle" bgcolor="#00CC99" class = "bg_navigator" colspan="2" height="20">&nbsp;<b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">FEATURED 
                          LINKS</font></b></td>
                      </tr>
                      <tr> 
                        <td align="left" valign="top" colspan="2"> 
                          <% 
While ((Repeat2__numRows <> 0) AND (NOT rsFeat.EOF)) 
%>
                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td align="left" valign="middle" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                            </tr>
                          </table>
                          <table width="100%" border="0" cellspacing="0" cellpadding="5">
                            <tr valign="middle"> 
                              <td align="left" bgcolor="#CCCCCC" height="20"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="../DUdirectory/dirHitting.asp?id=<%=(rsFeat.Fields.Item("LINK_ID").Value)%>&url=<%=(rsFeat.Fields.Item("LINK_URL").Value)%>" target="_blank" onClick="window.location.reload(true);"><b><%= UCase((rsFeat.Fields.Item("LINK_NAME").Value)) %></b></a> <i>(<%=(rsFeat.Fields.Item("LINK_URL").Value)%>)&nbsp;<font color="#FF0000"> 
                                <% If rsFeat.Fields.Item("RATING").Value > 4.0 Then %>
                                <b>HOT!</b> 
                                <% End If %>
                                </font>&nbsp;<font color="#0000FF"> 
                                <% If rsFeat.Fields.Item("NO_HITS").Value > 50 Then %>
                                <b><font color="#FF6633">POPULAR!</font></b> 
                                <% End If %>
                                </font>&nbsp;<font color="#0000FF"> 
                                <% If rsFeat.Fields.Item("LINK_DATE").Value > date() - 7 Then %>
                                <b><font color="#FF00FF">NEW!</font></b> 
                                <% End If %>
                                </font></i></font></td>
                              <td align="right" bgcolor="#CCCCCC"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsFeat.Fields.Item("LINK_DATE").Value)%></font></td>
                            </tr>
                            <tr valign="middle"> 
                              <td align="left" height="25"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Hits:</b> 
                                <%=(rsFeat.Fields.Item("NO_HITS").Value)%> | <b>Rating: </b> <img src="../assets/<%= FormatNumber((rsFeat.Fields.Item("RATING").Value), 1, -2, -2, -2) %>.gif" align="absmiddle"> 
                                </font></td>
                              <td align="right"><b><font size="1"><font face="Verdana, Arial, Helvetica, sans-serif"> 
                                <a href="dirRate.asp?id=<%=(rsFeat.Fields.Item("LINK_ID").Value)%>">Review 
                                (<%=(rsFeat.Fields.Item("COMMENTS").Value)%>) </a> | <a href="dirRate.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "id=" & rsFeat.Fields.Item("LINK_ID").Value %>"> 
                                Rate (<%=(rsFeat.Fields.Item("NO_RATES").Value)%>) </a></font></font></b></td>
                            </tr>
                            <tr> 
                              <td align="left" valign="top" colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Description:</b> 
                                <%=(rsFeat.Fields.Item("LINK_DESC").Value)%> </font></td>
                            </tr>
                          </table>
                          <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  rsFeat.MoveNext()
Wend
%>
                        </td>
                      </tr>
                      <tr> 
                        <td align="left" valign="middle" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                      </tr>
                      <tr> 
                        <td align="left" valign="middle" bgcolor="#00CC99" height="20" class = "bg_navigator">&nbsp;<b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%= UCase((rsSub.Fields.Item("SUB_NAME").Value)) %> LINKS</font></b></td>
                        <td align="right" valign="middle" height="20" class = "bg_navigator"> 
                          <font size="1"> <font face="Verdana, Arial, Helvetica, sans-serif"> 
                          <b> SHOWING LINKS 
                          <%
For i = 1 to rsLink_total Step MM_size
TM_endCount = i + MM_size - 1
if TM_endCount > rsLink_total Then TM_endCount = rsLink_total
if i <> MM_offset + 1 Then
Response.Write("<a href=""" & Request.ServerVariables("URL") & "?" & MM_keepMove & "offset=" & i-1 & """>")
Response.Write(i & "-" & TM_endCount & "</a>")
else
Response.Write("<b>" & i & "-" & TM_endCount & "</b>")
End if
if(TM_endCount <> rsLink_total) then Response.Write(" | ")
next
 %>
                          </b></font></font>&nbsp;</td>
                      </tr>
                      <tr> 
                        <td align="left" valign="middle" colspan="2"> 
                          <% 
While ((Repeat1__numRows <> 0) AND (NOT rsLink.EOF)) 
%>
                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td align="left" valign="middle" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                            </tr>
                          </table>
                          <table width="100%" border="0" cellspacing="0" cellpadding="5">
                            <tr valign="middle"> 
                              <td align="left" bgcolor="#CCCCCC" height="20"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="../DUdirectory/dirHitting.asp?id=<%=(rsLink.Fields.Item("LINK_ID").Value)%>&url=<%=(rsLink.Fields.Item("LINK_URL").Value)%>" target="_blank" onClick="window.location.reload(true);"><b><%= UCase((rsLink.Fields.Item("LINK_NAME").Value)) %></b></a> <i>(<%=(rsLink.Fields.Item("LINK_URL").Value)%>)&nbsp;<font color="#FF0000"> 
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
                              <td align="right" bgcolor="#CCCCCC"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsLink.Fields.Item("LINK_DATE").Value)%></font></td>
                            </tr>
                            <tr valign="middle"> 
                              <td align="left" height="25"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Hits: 
                                <%=(rsLink.Fields.Item("NO_HITS").Value)%> </b> | <b>Rating: </b> <img src="../assets/<%= FormatNumber((rsLink.Fields.Item("RATING").Value), 1, -2, -2, -2) %>.gif" align="absmiddle"> 
                                </font></td>
                              <td align="right"><b><font size="1"><font face="Verdana, Arial, Helvetica, sans-serif"> 
                                <a href="dirRate.asp?id=<%=(rsLink.Fields.Item("LINK_ID").Value)%>">Review 
                                (<%=(rsLink.Fields.Item("COMMENTS").Value)%>)</a> |<a href="dirRate.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "id=" & rsLink.Fields.Item("LINK_ID").Value %>"> 
                                Rate (<%=(rsLink.Fields.Item("NO_RATES").Value)%>) </a></font></font></b></td>
                            </tr>
                            <tr> 
                              <td align="left" valign="top" colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Description:</b> 
                                <%=(rsLink.Fields.Item("LINK_DESC").Value)%> </font></td>
                            </tr>
                          </table>
                          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsLink.MoveNext()
Wend
%>
                        </td>
                      </tr>
                      <tr> 
                        <td align="left" valign="middle" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
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
rsSub.Close()
%>
<%
rsCat.Close()
%>
<%
rsFeat.Close()
%>
<%
rsLink.Close()
%>

