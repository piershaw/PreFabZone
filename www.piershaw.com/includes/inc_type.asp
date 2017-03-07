
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
Dim rsType__MMColParam
rsType__MMColParam = "0"
if (Request.QueryString("iCat") <> "") then rsType__MMColParam = Request.QueryString("iCat")
%>
<%
set rsType = Server.CreateObject("ADODB.Recordset")
rsType.ActiveConnection = MM_connDUportal_STRING
rsType.Source = "SELECT *  FROM DATAS,  CATEGORIES, CHANNELS  WHERE DAT_CATEGORY = CAT_ID AND CAT_CHANNEL = CHA_ID AND DAT_APPROVED=1 AND CHA_ACTIVE=1 AND DAT_EXPIRED > DATE() AND DAT_CATEGORY = " + Replace(rsType__MMColParam, "'", "''") + "  ORDER BY DAT_DATED DESC"
rsType.CursorType = 0
rsType.CursorLocation = 2
rsType.LockType = 3
rsType.Open()
rsType_numRows = 0
%>
<%
Dim rsType__numRows
rsType__numRows = 20
Dim rsType__index
rsType__index = 0
rsType_numRows = rsType_numRows + rsType__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsType_total
Dim rsType_first
Dim rsType_last

' set the record count
rsType_total = rsType.RecordCount

' set the number of rows displayed on this page
If (rsType_numRows < 0) Then
  rsType_numRows = rsType_total
Elseif (rsType_numRows = 0) Then
  rsType_numRows = 1
End If

' set the first and last displayed record
rsType_first = 1
rsType_last  = rsType_first + rsType_numRows - 1

' if we have the correct record count, check the other stats
If (rsType_total <> -1) Then
  If (rsType_first > rsType_total) Then
    rsType_first = rsType_total
  End If
  If (rsType_last > rsType_total) Then
    rsType_last = rsType_total
  End If
  If (rsType_numRows > rsType_total) Then
    rsType_numRows = rsType_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsType_total = -1) Then

  ' count the total records by iterating through the recordset
  rsType_total=0
  While (Not rsType.EOF)
    rsType_total = rsType_total + 1
    rsType.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsType.CursorType > 0) Then
    rsType.MoveFirst
  Else
    rsType.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsType_numRows < 0 Or rsType_numRows > rsType_total) Then
    rsType_numRows = rsType_total
  End If

  ' set the first and last displayed record
  rsType_first = 1
  rsType_last = rsType_first + rsType_numRows - 1
  
  If (rsType_first > rsType_total) Then
    rsType_first = rsType_total
  End If
  If (rsType_last > rsType_total) Then
    rsType_last = rsType_total
  End If

End If
%>
<%
' *** Move To Record and Go To Record: declare variables
Set MM_rs    = rsType
MM_rsCount   = rsType_total
MM_size      = rsType_numRows
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
rsType_first = MM_offset + 1
rsType_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsType_first > MM_rsCount) Then
    rsType_first = MM_rsCount
  End If
  If (rsType_last > MM_rsCount) Then
    rsType_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters
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
                        &raquo; <% If Not rsType.EOF Or Not rsType.BOF Then %>
                        <a href="channel.asp?iChannel=<%=(rsType.Fields.Item("CHA_ID").Value)%>&nChannel=<%=(rsType.Fields.Item("CHA_NAME").Value)%>"><%=UCASE(rsType.Fields.Item("CHA_MENU").Value)%></a> &raquo; <a href="type.asp?iCat=<%=(rsType.Fields.Item("CAT_ID").Value)%>&iChannel=<%=(rsType.Fields.Item("CHA_ID").Value)%>&nChannel=<%=(rsType.Fields.Item("CHA_NAME").Value)%>"><%=UCASE(rsType.Fields.Item("CAT_NAME").Value)%></a> 
                        <% End If ' end Not rsType.EOF Or NOT rsType.BOF %> </td>
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
				
				
				<% If Not rsType.EOF Or Not rsType.BOF Then %>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td> <% 
While ((rsType__numRows <> 0) AND (NOT rsType.EOF)) 
%>
                        <%
Dim dat_rated
Dim dat_rate_count 
Dim dat_rate_value
dat_rate_count = rsType.Fields.Item("DAT_RATES").Value
dat_rate_value = rsType.Fields.Item("DAT_RATED").Value
If dat_rate_count > 0 Then 
dat_rated = (dat_rate_value/dat_rate_count)
else
dat_rated = 0
end if
%>
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td align="left" valign="top"> <table width="100%" border="0" cellspacing="2" cellpadding="2">
                                <tr valign="top"> 
                                  <td colspan="2" align="left" class="text"><b>Name:</b> 
                                    <a href="detail.asp?iData=<%=(rsType.Fields.Item("DAT_ID").Value)%>&iCat=<%=(rsType.Fields.Item("DAT_CATEGORY").Value)%>&iChannel=<%=(rsType.Fields.Item("CAT_CHANNEL").Value)%>&nChannel=<%=(rsType.Fields.Item("CHA_NAME").Value)%>"><%=(rsType.Fields.Item("DAT_NAME").Value)%></a></td>
                                </tr>
                                <tr valign="top"> 
                                  <td width="50%" align="left" class="text"><strong>Category:</strong> 
                                    <a href="type.asp?iCat=<%=(rsType.Fields.Item("CAT_ID").Value)%>"><%=(rsType.Fields.Item("CAT_NAME").Value)%></a></td>
                                  <td width="50%" align="left" class="text"><strong>Views:</strong> 
                                    <%=(rsType.Fields.Item("DAT_HITS").Value)%></td>
                                </tr>
                                <tr valign="top"> 
                                  <td width="50%" align="left" class="text"><strong>Rating:</strong> 
                                    <img src="../assets/<%= FormatNumber(dat_rated, 1, -2, -2, -2) %>.gif" align="absmiddle"> 
                                    (<%= FormatNumber(dat_rated, 1, -2, -2, -2) %>) </td>
                                  <td width="50%" align="left" class="text"><strong>By:</strong> 
                                    <%=(rsType.Fields.Item("DAT_RATES").Value)%> users</td>
                                </tr>
                                <tr> 
                                  <td colspan="2" align="left" valign="middle" class="text"><b>Description:</b> 
                                    <% =TrimBody(DoTrimProperly((rsType.Fields.Item("DAT_DESCRIPTION").Value), 100, 1, 1, " ...")) %> </td>
                                </tr>
                              </table></td>
                          </tr>
                          <tr> 
                            <td align="left" valign="top" class="bgTableBorder" ><img src="../assets/_spacer.gif" width="1" height="1"></td>
                          </tr>
                        </table>
                        <% 
  rsType__index=rsType__index+1
  rsType__numRows=rsType__numRows-1
  rsType.MoveNext()
Wend
%> </td>
                    </tr>
                    <tr> 
                      <td height="20" align="center" valign="middle" class="textBold">
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
If TFM_endLink > rsType_total Then TFM_endLink = rsType_total 
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
%> </td>
                    </tr>
                  </table>
<% End If ' end Not rsType.EOF Or NOT rsType.BOF %>
				
				
				
				
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
<%
rsType.Close()
%>
