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
Dim rsHot__MMColParam
rsHot__MMColParam = "0"
if (Request.QueryString("iChannel") <> "") then rsHot__MMColParam = Request.QueryString("iChannel")
%>
<%
set rsHot = Server.CreateObject("ADODB.Recordset")
rsHot.ActiveConnection = MM_connDUportal_STRING
rsHot.Source = "SELECT *  FROM DATAS,  CATEGORIES, CHANNELS  WHERE DAT_CATEGORY = CAT_ID AND CAT_CHANNEL = CHA_ID AND DAT_APPROVED=1  AND DAT_PARENT=0 AND CHA_ACTIVE=1 AND CAT_CHANNEL = " + Replace(rsHot__MMColParam, "'", "''") + "  AND DAT_EXPIRED > DATE() ORDER BY DAT_HITS DESC"
rsHot.CursorType = 0
rsHot.CursorLocation = 2
rsHot.LockType = 3
rsHot.Open()
rsHot_numRows = 0
%>
<%
Dim rsHot__numRows
rsHot__numRows = 5
Dim rsHot__index
rsHot__index = 0
rsHot_numRows = rsHot_numRows + rsHot__numRows
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
                        &raquo; <% If Not rsHot.EOF Or Not rsHot.BOF Then %>
                        <%=UCASE(rsHot.Fields.Item("CHA_MENU").Value)%> &raquo; HOT <%=UCASE(rsHot.Fields.Item("CHA_NAME").Value)%> 
                        <% End If ' end Not rsHot.EOF Or NOT rsHot.BOF %> </td>
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
				
				
				
				
				
				
<% If Not rsHot.EOF Or Not rsHot.BOF Then %>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
 
  <tr> 
    <td colspan="2"> <% 
While ((rsHot__numRows <> 0) AND (NOT rsHot.EOF)) 
%>
      <%
Dim dat_rated
Dim dat_rate_count 
Dim dat_rate_value
dat_rate_count = rsHot.Fields.Item("DAT_RATES").Value
dat_rate_value = rsHot.Fields.Item("DAT_RATED").Value
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
                                    <a href="detail.asp?iData=<%=(rsHot.Fields.Item("DAT_ID").Value)%>&iCat=<%=(rsHot.Fields.Item("DAT_CATEGORY").Value)%>&iChannel=<%=(rsHot.Fields.Item("CAT_CHANNEL").Value)%>&nChannel=<%=(rsHot.Fields.Item("CHA_NAME").Value)%>"><%=(rsHot.Fields.Item("DAT_NAME").Value)%></a></td>
                                </tr>
                                <tr valign="top"> 
                                  <td width="50%" align="left" class="text"><strong>Category:</strong> 
                                    <a href="type.asp?iCat=<%=(rsHot.Fields.Item("CAT_ID").Value)%>&iChannel=<%=(rsHot.Fields.Item("CHA_ID").Value)%>&nChannel=<%=(rsHot.Fields.Item("CHA_NAME").Value)%>"><%=(rsHot.Fields.Item("CAT_NAME").Value)%></a></td>
                                  <td width="50%" align="left" class="text"><strong>Views:</strong> 
                                    <%=(rsHot.Fields.Item("DAT_HITS").Value)%></td>
                                </tr>
                                <tr valign="top"> 
                                  <td width="50%" align="left" class="text"><strong>Rating:</strong> 
                                    <img src="../assets/<%= FormatNumber(dat_rated, 1, -2, -2, -2) %>.gif" align="absmiddle"> 
                                    (<%= FormatNumber(dat_rated, 1, -2, -2, -2) %>) </td>
                                  <td width="50%" align="left" class="text"><strong>By:</strong> 
                                    <%=(rsHot.Fields.Item("DAT_RATES").Value)%> users</td>
                                </tr>
                                <tr> 
                                  <td colspan="2" align="left" valign="middle" class="text"><b>Description:</b> 
                                    <% =TrimBody(DoTrimProperly((rsHot.Fields.Item("DAT_DESCRIPTION").Value), 100, 1, 1, " ...")) %> </td>
                                </tr>
                              </table></td>
        </tr>
		 <tr> 
          <td align="left" valign="top" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
        </tr>
      </table>
      <% 
  rsHot__index=rsHot__index+1
  rsHot__numRows=rsHot__numRows-1
  rsHot.MoveNext()
Wend
%> </td>
  </tr>
</table>
<% End If ' end Not rsHot.EOF Or NOT rsHot.BOF %>
				
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
rsHot.Close()
%>