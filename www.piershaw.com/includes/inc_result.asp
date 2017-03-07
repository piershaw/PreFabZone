<!--#include file="../Connections/connDUportal.asp" -->
<%
'****************************************************************************************
'**  Copyright Notice                                                               
'**  Copyright 2003 DUware All Rights Reserved.                                
'**  This program is free software; you can modify (at your own risk) any part of it 
'**  under the terms of the License that accompanies this software and use it both 
'**  privately and commercially.
'**  All copyright notices must remain in tacked in the scripts and the 
'**  outputted HTML.
'**  You may use parts of this program in your own private work, but you may NOT
'**  redistribute, repackage, or sell the whole or any part of this program even 
'**  if it is modified or reverse engineered in whole or in part without express 
'**  permission from the author.
'**  You may not pass the whole or any part of this application off as your own work.
'**  All links to DUware and powered by logo's must remain unchanged and in place
'**  and must remain visible when the pages are viewed unless permission is first granted
'**  by the copyright holder.
'**  This program is distributed in the hope that it will be useful,
'**  but WITHOUT ANY WARRANTY; without even the implied warranty of
'**  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE OR ANY OTHER 
'**  WARRANTIES WHETHER EXPRESSED OR IMPLIED.
'**  No official support is available for this program but you may post support questions at: -
'**  http://www.duware.com/support
'****************************************************************************************
%>
<%
Dim strChannel
strChannel = Request.QueryString("iChannel")
If strChannel = "All" Then
sqlChannel = ""
Else
sqlChannel = " AND CHA_ID = " & strChannel
End If

Dim strKeyword
strKeyword = Replace(Request.QueryString("keyword"),"'", "''")
If strKeyword = "" Then
sqlKeyword = ""
Else
sqlKeyword = " AND (DAT_NAME LIKE '%" & strKeyword & "%' OR DAT_DESCRIPTION LIKE '%" & strKeyword & "%') "
End If
%>


<%
set rsSearchResult = Server.CreateObject("ADODB.Recordset")
rsSearchResult.ActiveConnection = MM_connDUportal_STRING
rsSearchResult.Source = "SELECT *  FROM DATAS, CATEGORIES, CHANNELS  WHERE DAT_CATEGORY = CAT_ID AND CHA_ID = CAT_CHANNEL " & sqlChannel & sqlKeyword & " AND DAT_APPROVED = 1 AND CHA_ACTIVE=1 AND DAT_EXPIRED > DATE() AND DAT_PARENT=0 ORDER BY CHA_MENU, CAT_NAME, DAT_NAME"
rsSearchResult.CursorType = 0
rsSearchResult.CursorLocation = 2
rsSearchResult.LockType = 3
rsSearchResult.Open()
rsSearchResult_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 50
Repeat1__index = 0
rsSearchResult_numRows = rsSearchResult_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsSearchResult_total
Dim rsSearchResult_first
Dim rsSearchResult_last

' set the record count
rsSearchResult_total = rsSearchResult.RecordCount

' set the number of rows displayed on this page
If (rsSearchResult_numRows < 0) Then
  rsSearchResult_numRows = rsSearchResult_total
Elseif (rsSearchResult_numRows = 0) Then
  rsSearchResult_numRows = 1
End If

' set the first and last displayed record
rsSearchResult_first = 1
rsSearchResult_last  = rsSearchResult_first + rsSearchResult_numRows - 1

' if we have the correct record count, check the other stats
If (rsSearchResult_total <> -1) Then
  If (rsSearchResult_first > rsSearchResult_total) Then
    rsSearchResult_first = rsSearchResult_total
  End If
  If (rsSearchResult_last > rsSearchResult_total) Then
    rsSearchResult_last = rsSearchResult_total
  End If
  If (rsSearchResult_numRows > rsSearchResult_total) Then
    rsSearchResult_numRows = rsSearchResult_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsSearchResult_total = -1) Then

  ' count the total records by iterating through the recordset
  rsSearchResult_total=0
  While (Not rsSearchResult.EOF)
    rsSearchResult_total = rsSearchResult_total + 1
    rsSearchResult.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsSearchResult.CursorType > 0) Then
    rsSearchResult.MoveFirst
  Else
    rsSearchResult.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsSearchResult_numRows < 0 Or rsSearchResult_numRows > rsSearchResult_total) Then
    rsSearchResult_numRows = rsSearchResult_total
  End If

  ' set the first and last displayed record
  rsSearchResult_first = 1
  rsSearchResult_last = rsSearchResult_first + rsSearchResult_numRows - 1
  
  If (rsSearchResult_first > rsSearchResult_total) Then
    rsSearchResult_first = rsSearchResult_total
  End If
  If (rsSearchResult_last > rsSearchResult_total) Then
    rsSearchResult_last = rsSearchResult_total
  End If

End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = rsSearchResult
MM_rsCount   = rsSearchResult_total
MM_size      = rsSearchResult_numRows
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
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

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
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
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
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsSearchResult_first = MM_offset + 1
rsSearchResult_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsSearchResult_first > MM_rsCount) Then
    rsSearchResult_first = MM_rsCount
  End If
  If (rsSearchResult_last > MM_rsCount) Then
    rsSearchResult_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

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

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = MM_keepMove & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>

 <link href="../assets/DUportal.css" rel="stylesheet" type="text/css"> 
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  
  <tr>
    <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#003399">
              <tr> 
                <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../assets/bg_header.gif">
                    <tr> 
                      <td width="10"><img src="../assets/header_end_left.gif"></td>
                      <td align="left" valign="middle" class="textBoldColor">SEARCH 
                        RESULT </td>
                      <td width="28" align="right" valign="middle"><img src="../assets/header_end_right.gif"></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif"></td>
                <td align="left" valign="top" class="bgTable"><table width="100%" border="0" cellspacing="2" cellpadding="2">
                    <tr>
                      <td align="left" valign="middle" class="textBold">
                        <%
TFM_MiddlePages = 10
TFM_delimiter = " | "
TFM_startLink = MM_offset + 1 - MM_size * (int(TFM_middlePages/2))
If MM_offset > 0 Then TFM_LimitPageEndCount = int(TFM_startLink/MM_size)
If TFM_startLink < 1 Then 
	TFM_startLink = 1
	TFM_LimitPageEndCount = 0
End If
TFM_endLink = MM_size * TFM_MiddlePages + TFM_startLink - 1
If TFM_endLink > rsSearchResult_total Then TFM_endLink = rsSearchResult_total 
For i = TFM_startLink to TFM_endLink Step MM_size
  TFM_LimitPageEndCount = TFM_LimitPageEndCount + 1
  if i <> MM_offset + 1 Then
    Response.Write("<a href=""" & Request.ServerVariables("URL") & "?" & MM_keepMove & "offset=" & i-1 & """>")
    Response.Write(TFM_LimitPageEndCount & "</a>")
  else
    Response.Write("<b>Page ")
    Response.Write(TFM_LimitPageEndCount & "</b>")
  End if
  if(i <= TFM_endLink - MM_size) then Response.Write(TFM_delimiter)
Next
%>
                      </td>
                    </tr>
					
                   
                    <tr> 
                      <td align="left" valign="top"> 
                        <table width="100%" border="0" cellpadding="4" cellspacing="1" class="bgMouseOver">
                          <tr align="center" valign="middle" bgcolor="#CCCCCC" class="textBold">
                            <td class="text">&nbsp;</td>
                            <td class="text">CHANNEL</td>
                            <td class="text">CATEGORY</td>
                            <td class="text">NAME</td>
                            <td class="text">DATE</td>
                          </tr>
                          <% 
While ((Repeat1__numRows <> 0) AND (NOT rsSearchResult.EOF)) 
%>
                          <tr valign="middle" class="bgTable">
                            <td width="5" align="center"><img src="../assets/icon_bullet_square.gif" width="5" height="5" align="absmiddle"></td>
                            <td align="left" class="text"><a href="../home/channel.asp?iChannel=<%=(rsSearchResult.Fields.Item("CAT_CHANNEL").Value)%>&nChannel=<%=(rsSearchResult.Fields.Item("CHA_NAME").Value)%>"> 
                              <%=(rsSearchResult.Fields.Item("CHA_MENU").Value)%></a></td>
                            <td align="left" class="text"><a href="../home/type.asp?iCat=<%=(rsSearchResult.Fields.Item("DAT_CATEGORY").Value)%>&iChannel=<%=(rsSearchResult.Fields.Item("CAT_CHANNEL").Value)%>&nChannel=<%=(rsSearchResult.Fields.Item("CHA_NAME").Value)%>"> 
                              <%=(rsSearchResult.Fields.Item("CAT_NAME").Value)%></a></td>
                            <td align="left" class="text"> 
                              <a href="../home/detail.asp?iData=<%=(rsSearchResult.Fields.Item("DAT_ID").Value)%>&iCat=<%=(rsSearchResult.Fields.Item("DAT_CATEGORY").Value)%>&iChannel=<%=(rsSearchResult.Fields.Item("CAT_CHANNEL").Value)%>&nChannel=<%=(rsSearchResult.Fields.Item("CHA_NAME").Value)%>"> 
                              <%= (rsSearchResult.Fields.Item("DAT_NAME").Value) %> </a></td>
                            <td align="center" width="80" class="text"><%=(rsSearchResult.Fields.Item("DAT_DATED").Value)%></td>
                          </tr>
                          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsSearchResult.MoveNext()
Wend
%>
                        </table>

                        </td>
                    </tr>
					



                    <tr> 
                      <td align="right" valign="middle" class="textBold"> 
                        <%
TFM_MiddlePages = 10
TFM_delimiter = " | "
TFM_startLink = MM_offset + 1 - MM_size * (int(TFM_middlePages/2))
If MM_offset > 0 Then TFM_LimitPageEndCount = int(TFM_startLink/MM_size)
If TFM_startLink < 1 Then 
	TFM_startLink = 1
	TFM_LimitPageEndCount = 0
End If
TFM_endLink = MM_size * TFM_MiddlePages + TFM_startLink - 1
If TFM_endLink > rsSearchResult_total Then TFM_endLink = rsSearchResult_total 
For i = TFM_startLink to TFM_endLink Step MM_size
  TFM_LimitPageEndCount = TFM_LimitPageEndCount + 1
  if i <> MM_offset + 1 Then
    Response.Write("<a href=""" & Request.ServerVariables("URL") & "?" & MM_keepMove & "offset=" & i-1 & """>")
    Response.Write(TFM_LimitPageEndCount & "</a>")
  else
    Response.Write("<b>Page ")
    Response.Write(TFM_LimitPageEndCount & "</b>")
  End if
  if(i <= TFM_endLink - MM_size) then Response.Write(TFM_delimiter)
Next
%>
                      </td>
                    </tr>
                  </table></td>
                <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif"></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td align="center" valign="top" background="../assets/bg_header_bottom.gif"></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="7" align="left" valign="top"><img src="../assets/_spacer.gif" width="1" height="1"></td>
  </tr>
</table>

<%
rsSearchResult.Close()
%>
