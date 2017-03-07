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
Dim rsHotTopics__MMColParam
rsHotTopics__MMColParam = "0"
if (Request.QueryString("iChannel") <> "") then rsHotTopics__MMColParam = Request.QueryString("iChannel")
%>
<%
set rsHotTopics = Server.CreateObject("ADODB.Recordset")
rsHotTopics.ActiveConnection = MM_connDUportal_STRING
rsHotTopics.Source = "SELECT *  FROM DATAS,  CATEGORIES, CHANNELS  WHERE DAT_CATEGORY = CAT_ID AND CAT_CHANNEL = CHA_ID AND DAT_APPROVED=1  AND DAT_PARENT=0 AND CHA_ACTIVE=1 AND CAT_CHANNEL = " + Replace(rsHotTopics__MMColParam, "'", "''") + "  AND DAT_EXPIRED > DATE() ORDER BY DAT_HITS DESC"
rsHotTopics.CursorType = 0
rsHotTopics.CursorLocation = 2
rsHotTopics.LockType = 3
rsHotTopics.Open()
rsHotTopics_numRows = 0
%>
<%
Dim rsHotTopics__numRows
rsHotTopics__numRows = 5
Dim rsHotTopics__index
rsHotTopics__index = 0
rsHotTopics_numRows = rsHotTopics_numRows + rsHotTopics__numRows
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
                      <td align="left" valign="middle" class="textBoldColor">HOME 
                        &raquo; <% If Not rsHotTopics.EOF Or Not rsHotTopics.BOF Then %>
                        <%=UCASE(rsHotTopics.Fields.Item("CHA_MENU").Value)%> &raquo; HOT <%=UCASE(rsHotTopics.Fields.Item("CHA_NAME").Value)%> 
                        <% End If ' end Not rsHotTopics.EOF Or NOT rsHotTopics.BOF %> </td>
                      <td width="28" align="right" valign="middle"><img src="../assets/header_end_right.gif"></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td align="left" valign="top">
		  
		  <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif"></td>
                <td align="left" valign="top" class="bgTable">
				
				
				
				
				
				
<% If Not rsHotTopics.EOF Or Not rsHotTopics.BOF Then %>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
 
  <tr> 
    <td colspan="2"> <% 
While ((rsHotTopics__numRows <> 0) AND (NOT rsHotTopics.EOF)) 
%>

      <table width="100%" border="0" cellspacing="0" cellpadding="0">
       
        <tr> 
          <td align="left" valign="top"> <table width="100%" border="0" cellspacing="2" cellpadding="2">
                                <tr valign="top"> 
                                  <td colspan="2" align="left" class="text"><b>Topic:</b> 
                                    <a href="detail.asp?iData=<%=(rsHotTopics.Fields.Item("DAT_ID").Value)%>&iCat=<%=(rsHotTopics.Fields.Item("DAT_CATEGORY").Value)%>&iChannel=<%=(rsHotTopics.Fields.Item("CAT_CHANNEL").Value)%>&nChannel=<%=(rsHotTopics.Fields.Item("CHA_NAME").Value)%>"><%=(rsHotTopics.Fields.Item("DAT_NAME").Value)%></a></td>
                                </tr>
                                <tr valign="top"> 
                                  <td width="50%" align="left" class="text"><strong>Forum:</strong> 
                                    <a href="type.asp?iCat=<%=(rsHotTopics.Fields.Item("CAT_ID").Value)%>&iChannel=<%=(rsHotTopics.Fields.Item("CHA_ID").Value)%>&nChannel=<%=(rsHotTopics.Fields.Item("CHA_NAME").Value)%>"><%=(rsHotTopics.Fields.Item("CAT_NAME").Value)%></a></td>
                                  <td width="50%" align="left" class="text"><strong>Reads:</strong> 
                                    <%=(rsHotTopics.Fields.Item("DAT_HITS").Value)%></td>
                                </tr>
                                <tr> 
                                  <td colspan="2" align="left" valign="middle" class="text"><b>Message:</b> 
                                    <% =TrimBody(DoTrimProperly((rsHotTopics.Fields.Item("DAT_DESCRIPTION").Value), 200, 1, 1, " ...")) %> </td>
                                </tr>
                              </table></td>
        </tr>
		 <tr> 
          <td align="left" valign="top" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
        </tr>
      </table>
      <% 
  rsHotTopics__index=rsHotTopics__index+1
  rsHotTopics__numRows=rsHotTopics__numRows-1
  rsHotTopics.MoveNext()
Wend
%> </td>
  </tr>
</table>
<% End If ' end Not rsHotTopics.EOF Or NOT rsHotTopics.BOF %>
				
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
rsHotTopics.Close()
%>