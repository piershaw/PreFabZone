<!--#include file="../Connections/connDUportal.asp" -->

<!--#include file="inc_restriction.asp" -->

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
Dim tfm_orderby, tfm_order
tfm_orderby = "DAT_DATED"
tfm_order = "DESC"
If(CStr(Request.QueryString("tfm_orderby")) <> "") Then
	tfm_orderby = Cstr(Request.QueryString("tfm_orderby"))
End If
If(Cstr(Request.QueryString("tfm_order")) <> "") Then
	tfm_order = Cstr(Request.QueryString("tfm_order"))
End If

Dim sql_orderby
sql_orderby = " " & tfm_orderby & " " & tfm_order
%>
<%
Dim rsDatas__sql_orderby
rsDatas__sql_orderby = "DAT_DATED"
if (sql_orderby <> "") then rsDatas__sql_orderby = sql_orderby
%>
<%
set rsDatas = Server.CreateObject("ADODB.Recordset")
rsDatas.ActiveConnection = MM_connDUportal_STRING
rsDatas.Source = "SELECT *  FROM DATAS, CATEGORIES, CHANNELS WHERE DAT_CATEGORY = CAT_ID  AND CHA_ID = CAT_CHANNEL AND DAT_EXPIRED > DATE() AND DAT_PARENT=0 ORDER BY " + Replace(rsDatas__sql_orderby, "'", "''") + ""
rsDatas.CursorType = 0
rsDatas.CursorLocation = 2
rsDatas.LockType = 3
rsDatas.Open()
rsDatas_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = 30
Dim Repeat1__index
Repeat1__index = 0
rsDatas_numRows = rsDatas_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
rsDatas_total = rsDatas.RecordCount

' set the number of rows displayed on this page
If (rsDatas_numRows < 0) Then
  rsDatas_numRows = rsDatas_total
Elseif (rsDatas_numRows = 0) Then
  rsDatas_numRows = 1
End If

' set the first and last displayed record
rsDatas_first = 1
rsDatas_last  = rsDatas_first + rsDatas_numRows - 1

' if we have the correct record count, check the other stats
If (rsDatas_total <> -1) Then
  If (rsDatas_first > rsDatas_total) Then rsDatas_first = rsDatas_total
  If (rsDatas_last > rsDatas_total) Then rsDatas_last = rsDatas_total
  If (rsDatas_numRows > rsDatas_total) Then rsDatas_numRows = rsDatas_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsDatas_total = -1) Then

  ' count the total records by iterating through the recordset
  rsDatas_total=0
  While (Not rsDatas.EOF)
    rsDatas_total = rsDatas_total + 1
    rsDatas.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsDatas.CursorType > 0) Then
    rsDatas.MoveFirst
  Else
    rsDatas.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsDatas_numRows < 0 Or rsDatas_numRows > rsDatas_total) Then
    rsDatas_numRows = rsDatas_total
  End If

  ' set the first and last displayed record
  rsDatas_first = 1
  rsDatas_last = rsDatas_first + rsDatas_numRows - 1
  If (rsDatas_first > rsDatas_total) Then rsDatas_first = rsDatas_total
  If (rsDatas_last > rsDatas_total) Then rsDatas_last = rsDatas_total

End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = rsDatas
MM_rsCount   = rsDatas_total
MM_size      = rsDatas_numRows
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
rsDatas_first = MM_offset + 1
rsDatas_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (rsDatas_first > MM_rsCount) Then rsDatas_first = MM_rsCount
  If (rsDatas_last > MM_rsCount) Then rsDatas_last = MM_rsCount
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
<%
'sort column headers for rsDatas
Dim tfm_saveParams, tfm_keepParams, tfm_orderbyURL
tfm_saveParams = ""
tfm_keepParams = ""
If tfm_order = "ASC" Then
	tfm_order = "DESC"
Else
	tfm_order = "ASC"
End If
		
If tfm_saveParams <> "" Then
	tfm_params = Split(tfm_saveParams,",")
	For i = 0 to UBound(tfm_params)
		If Cstr(Request(tfm_params(i))) <> "" Then
			tfm_keepParams = tfm_keepParams & LCase(tfm_params(i)) & "=" & Server.URLEncode(Request(tfm_params(i))) & "&"
		End If
	Next
End If
tfm_orderbyURL = Request.ServerVariables("URL") & "?" & tfm_keepParams & "tfm_order=" & tfm_order & "&tfm_orderby="
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
                      <td align="left" valign="middle" class="textBoldColor">DATA</td>
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
                <td align="left" valign="top" class="bgTable"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="bgTable">
                          <tr> 
                            <td align="left" valign="top"> <table width="100%" border="0" cellspacing="2" cellpadding="2">
                                <tr> 
                                  <td align="right" valign="middle" class="textBold"> 
                                    <%
TM_counter = 0
For i = 1 to rsDatas_total Step MM_size
TM_counter = TM_counter + 1
TM_PageEndCount = i + MM_size - 1
if TM_PageEndCount > rsDatas_total Then TM_PageEndCount = rsDatas_total
if i <> MM_offset + 1 then
Response.Write("<a href=""" & Request.ServerVariables("URL") & "?" & MM_keepMove & "offset=" & i-1 & """>")
Response.Write(TM_counter & "</a>")
else
Response.Write("<b>Page " & TM_counter & "</b>")
End if
if(TM_PageEndCount <> rsDatas_total) then Response.Write(" : ")
next
 %>
                                  </td>
                                </tr>
                                <tr> 
                                  <td align="left" valign="top"> <table width="100%" border="0" cellpadding="3" cellspacing="1" class="bgTableBorder">
                                      <tr align="center" valign="middle" bgcolor="#CCCCCC" class="textBoldColor"> 
                                        <td height="18"><a href="<%=tfm_orderbyURL%>DAT_NAME">NAME</a></td>
                                        <td height="18"><a href="<%=tfm_orderbyURL%>CHA_MENU">CHANNEL</a></td>
                                        <td height="18"><a href="<%=tfm_orderbyURL%>CAT_NAME">CATEGORY</a></td>
                                        <td height="18"><a href="<%=tfm_orderbyURL%>DAT_DATED">DATED</a></td>
                                        <td height="18"><a href="<%=tfm_orderbyURL%>DAT_USER">AUTHOR</a></td>
                                        <td height="18">APPROVE</td>
                                        <td>EDIT</td>
                                        <td>DELETE</td>
                                      </tr>
                                      <% 
While ((Repeat1__numRows <> 0) AND (NOT rsDatas.EOF)) 
%>
                                      <tr align="center" valign="middle" class="text"> 
                                        <td align="left" bgcolor="#FFFFFF" class="textBold"><a href="../home/detail.asp?iData=<%=(rsDatas.Fields.Item("DAT_ID").Value)%>&iCat=<%=(rsDatas.Fields.Item("DAT_CATEGORY").Value)%>&iChannel=<%=(rsDatas.Fields.Item("CHA_ID").Value)%>&nChannel=<%=(rsDatas.Fields.Item("CHA_NAME").Value)%>" target="_blank"><%=(rsDatas.Fields.Item("DAT_NAME").Value)%></a></td>
                                        <td align="left" bgcolor="#FFFFFF"><%=(rsDatas.Fields.Item("CHA_MENU").Value)%></td>
                                        <td align="left" bgcolor="#FFFFFF"><%=(rsDatas.Fields.Item("CAT_NAME").Value)%></td>
                                        <td align="center" bgcolor="#FFFFFF"><%=(rsDatas.Fields.Item("DAT_DATED").Value)%></td>
                                        <td align="left" bgcolor="#FFFFFF"><%=(rsDatas.Fields.Item("DAT_USER").Value)%></td>
                                        <td bgcolor="#FFFFFF"><img src="../assets/icon_<%=(rsDatas.Fields.Item("DAT_APPROVED").Value)%>.gif" align="absmiddle" border="0"></a></td>
                                        <td bgcolor="#FFFFFF"><a href="datasEdit.asp?iData=<%=(rsDatas.Fields.Item("DAT_ID").Value)%>&iCat=<%=(rsDatas.Fields.Item("DAT_CATEGORY").Value)%>&iChannel=<%=(rsDatas.Fields.Item("CHA_ID").Value)%>&nChannel=<%=(rsDatas.Fields.Item("CHA_NAME").Value)%>"><img src="../assets/icon_edit_data.gif" alt="EDIT"  border="0" align="absmiddle"></a></td>
                                        <td bgcolor="#FFFFFF"><a href="datasDelete.asp?iData=<%=(rsDatas.Fields.Item("DAT_ID").Value)%>&iCat=<%=(rsDatas.Fields.Item("DAT_CATEGORY").Value)%>&iChannel=<%=(rsDatas.Fields.Item("CHA_ID").Value)%>&nChannel=<%=(rsDatas.Fields.Item("CHA_NAME").Value)%>"><img src="../assets/icon_delete_data.gif" alt="DELETE" border="0" align="absmiddle"></a></td>
                                      </tr>
                                      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsDatas.MoveNext()
Wend
%>
                                    </table></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
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
rsDatas.Close()
%>