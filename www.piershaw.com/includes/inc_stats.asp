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
Dim rsStats
Dim rsStats_numRows

Set rsStats = Server.CreateObject("ADODB.Recordset")
rsStats.ActiveConnection = MM_connDUportal_STRING
rsStats.Source = "SELECT (SELECT COUNT(*) FROM USERS) AS U_COUNT, CHA_NAME, COUNT(DAT_ID) AS D_COUNT FROM CHANNELS, DATAS, CATEGORIES  WHERE CHA_ID = CAT_CHANNEL AND CAT_ID = DAT_CATEGORY AND CHA_STATS=1 AND DAT_APPROVED=1 AND DAT_EXPIRED > DATE()  AND DAT_PARENT=0 GROUP BY CHA_NAME ORDER BY CHA_NAME ASC"
rsStats.CursorType = 0
rsStats.CursorLocation = 2
rsStats.LockType = 1
rsStats.Open()

rsStats_numRows = 0
%>
<%
Dim rsStats__numRows
Dim rsStats__index

rsStats__numRows = -1
rsStats__index = 0
rsStats_numRows = rsStats_numRows + rsStats__numRows
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
                      <td align="left" valign="middle" class="textBoldColor">STATISTICS</td>
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
                <td align="left" valign="top" class="bgTable"> <% If Not rsStats.EOF Or Not rsStats.BOF Then %>
                  <table width="100%" border="0" cellpadding="2" cellspacing="2">
                    <tr align="left" valign="middle"> 
                      <td width="5" class="textBold"><img src="../assets/icon_bullet_square.gif" align="absmiddle"></td>
                      <td height="16" class="textBold">Users:</td>
                      <td align="right" class="text"><%= (rsStats.Fields.Item("U_COUNT").Value) %></td>
                    </tr>
                    <% 
While ((rsStats__numRows <> 0) AND (NOT rsStats.EOF)) 
%>
                    <tr align="left" valign="middle"> 
                      <td class="textBold"><img src="../assets/icon_bullet_square.gif" align="absmiddle"></td>
                      <td height="16" class="textBold"><%= ((rsStats.Fields.Item("CHA_NAME").Value)) %>:</td>
                      <td align="right" class="text"><%= (rsStats.Fields.Item("D_COUNT").Value) %></td>
                    </tr>
                    <% 
  rsStats__index=rsStats__index+1
  rsStats__numRows=rsStats__numRows-1
  rsStats.MoveNext()
Wend
%>
                  </table>
                  <% End If ' end Not rsStats.EOF Or NOT rsStats.BOF %> </td>
                <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif"></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td align="center" valign="top" background="../assets/bg_header_bottom.gif"><table border="0" cellpadding="0" cellspacing="0" class="bgTable" >
              <tr> 
                <td><img src="../assets/header_bottom.gif"></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="5" align="left" valign="top"><img src="../assets/_spacer.gif" width="1" height="1"></td>
  </tr>
</table>
<%
rsStats.Close()
Set rsStats = Nothing
%>