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
Dim rsTypeForums__MMColParam
rsTypeForums__MMColParam = "0"
if (Request.QueryString("iCat") <> "") then rsTypeForums__MMColParam = Request.QueryString("iCat")
%>
<%
set rsTypeForums = Server.CreateObject("ADODB.Recordset")
rsTypeForums.ActiveConnection = MM_connDUportal_STRING
rsTypeForums.Source = "SELECT * FROM DATAS, CATEGORIES, CHANNELS  WHERE DAT_CATEGORY = CAT_ID AND CAT_CHANNEL = CHA_ID AND DAT_PARENT=0 AND DAT_CATEGORY = " + Replace(rsTypeForums__MMColParam, "'", "''") + "  ORDER BY DAT_LAST DESC"
rsTypeForums.CursorType = 0
rsTypeForums.CursorLocation = 2
rsTypeForums.LockType = 3
rsTypeForums.Open()
rsTypeForums_numRows = 0
%>
<%
Dim rsTypeForums__numRows
Dim rsTypeForums__index

rsTypeForums__numRows = 50
rsTypeForums__index = 0
rsTypeForums_numRows = rsTypeForums_numRows + rsTypeForums__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsTypeForums_total
Dim rsTypeForums_first
Dim rsTypeForums_last

' set the record count
rsTypeForums_total = rsTypeForums.RecordCount

' set the number of rows displayed on this page
If (rsTypeForums_numRows < 0) Then
  rsTypeForums_numRows = rsTypeForums_total
Elseif (rsTypeForums_numRows = 0) Then
  rsTypeForums_numRows = 1
End If

' set the first and last displayed record
rsTypeForums_first = 1
rsTypeForums_last  = rsTypeForums_first + rsTypeForums_numRows - 1

' if we have the correct record count, check the other stats
If (rsTypeForums_total <> -1) Then
  If (rsTypeForums_first > rsTypeForums_total) Then
    rsTypeForums_first = rsTypeForums_total
  End If
  If (rsTypeForums_last > rsTypeForums_total) Then
    rsTypeForums_last = rsTypeForums_total
  End If
  If (rsTypeForums_numRows > rsTypeForums_total) Then
    rsTypeForums_numRows = rsTypeForums_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsTypeForums_total = -1) Then

  ' count the total records by iterating through the recordset
  rsTypeForums_total=0
  While (Not rsTypeForums.EOF)
    rsTypeForums_total = rsTypeForums_total + 1
    rsTypeForums.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsTypeForums.CursorType > 0) Then
    rsTypeForums.MoveFirst
  Else
    rsTypeForums.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsTypeForums_numRows < 0 Or rsTypeForums_numRows > rsTypeForums_total) Then
    rsTypeForums_numRows = rsTypeForums_total
  End If

  ' set the first and last displayed record
  rsTypeForums_first = 1
  rsTypeForums_last = rsTypeForums_first + rsTypeForums_numRows - 1
  
  If (rsTypeForums_first > rsTypeForums_total) Then
    rsTypeForums_first = rsTypeForums_total
  End If
  If (rsTypeForums_last > rsTypeForums_total) Then
    rsTypeForums_last = rsTypeForums_total
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

Set MM_rs    = rsTypeForums
MM_rsCount   = rsTypeForums_total
MM_size      = rsTypeForums_numRows
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
rsTypeForums_first = MM_offset + 1
rsTypeForums_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsTypeForums_first > MM_rsCount) Then
    rsTypeForums_first = MM_rsCount
  End If
  If (rsTypeForums_last > MM_rsCount) Then
    rsTypeForums_last = MM_rsCount
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
                      <td align="left" valign="middle" class="textBoldColor"><a href="default.asp">HOME</a> 
                        &raquo; <% If Not rsTypeForums.EOF Or Not rsTypeForums.BOF Then %>
                        <a href="channel.asp?iChannel=<%=(rsTypeForums.Fields.Item("CHA_ID").Value)%>&nChannel=<%=(rsTypeForums.Fields.Item("CHA_NAME").Value)%>"><%=UCASE(rsTypeForums.Fields.Item("CHA_MENU").Value)%></a> &raquo; <a href="type.asp?iCat=<%=(rsTypeForums.Fields.Item("CAT_ID").Value)%>&iChannel=<%=(rsTypeForums.Fields.Item("CHA_ID").Value)%>&nChannel=<%=(rsTypeForums.Fields.Item("CHA_NAME").Value)%>"><%=UCASE(rsTypeForums.Fields.Item("CAT_NAME").Value)%></a> 
                        <% End If ' end Not rsTypeForums.EOF Or NOT rsTypeForums.BOF %> </td>
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
                <td align="left" valign="top" class="bgTable">
				
				
				<% If Not rsTypeForums.EOF Or Not rsTypeForums.BOF Then %>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="100%"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr align="center" valign="middle" class="textBoldColor"> 
                            <td>TOPIC</td>
                            <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                            <td width="60" >AUTHOR</td>
                            <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                            <td width="60" >REPLIES</td>
                            <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                            <td width="60" >READS</td>
                            <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                            <td width="80" >LAST POST</td>
                          </tr>
                          <% 
While ((rsTypeForums__numRows <> 0) AND (NOT rsTypeForums.EOF)) 
%>
                          <tr align="center" valign="middle" class="textBoldColor"> 
                            <td height="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                            <td height="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                            <td height="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                            <td height="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                            <td height="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                            <td height="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                            <td height="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                            <td height="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                            <td height="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                          </tr>
                          <tr align="center" valign="middle" class="text"> 
                            <td align="left"><table width="100%" border="0" cellspacing="2" cellpadding="2">
                                <tr> 
                                  <td width="5"><img src="../assets/icon_folder.gif" hspace="0" vspace="0" align="absmiddle"></td>
                                  <td align="left" valign="middle" class="textBoldColor"><a href="detail.asp?iData=<%=(rsTypeForums.Fields.Item("DAT_ID").Value)%>&iCat=<%=(rsTypeForums.Fields.Item("CAT_ID").Value)%>&iChannel=<%=(rsTypeForums.Fields.Item("CHA_ID").Value)%>&nChannel=<%=(rsTypeForums.Fields.Item("CHA_NAME").Value)%>"><%=(rsTypeForums.Fields.Item("DAT_NAME").Value)%></a></td>
                                </tr>
                              </table></td>
                            <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                            <td ><%=(rsTypeForums.Fields.Item("DAT_USER").Value)%></td>
                            <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                            <td ><%=(rsTypeForums.Fields.Item("DAT_COUNT").Value)%></td>
                            <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                            <td ><%=(rsTypeForums.Fields.Item("DAT_HITS").Value)%></td>
                            <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                            <td><%=(rsTypeForums.Fields.Item("DAT_LAST").Value)%></td>
                          </tr>
                          <% 
  rsTypeForums__index=rsTypeForums__index+1
  rsTypeForums__numRows=rsTypeForums__numRows-1
  rsTypeForums.MoveNext()
Wend
%>
                        </table></td>
                    </tr>
                  </table>
<% End If ' end Not rsTypeForums.EOF Or NOT rsTypeForums.BOF %>
				
				
				
				
				</td>
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

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr align="left" valign="middle" class="textBold">
    <td>
<%
TFM_MiddlePages = 5
TFM_delimiter = " | "
TFM_startLink = MM_offset + 1 - MM_size * (int(TFM_middlePages/2))
If MM_offset > 0 Then TFM_LimitPageEndCount = int(TFM_startLink/MM_size)
If TFM_startLink < 1 Then 
	TFM_startLink = 1
	TFM_LimitPageEndCount = 0
End If
TFM_endLink = MM_size * TFM_MiddlePages + TFM_startLink - 1
If TFM_endLink > rsTypeForums_total Then TFM_endLink = rsTypeForums_total 
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
%></td>
    <td align="right"><img src="../assets/icon_topic_new.gif" hspace="4" vspace="1" align="absmiddle"><a href="../home/post.asp?iData=0&iCat=<%=Request.QueryString("iCat") %>">POST 
      NEW TOPIC</a></td>
  </tr>
</table>
<%
rsTypeForums.Close()
%>