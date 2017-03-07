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
Dim rsMenu
Dim rsMenu_numRows

Set rsMenu = Server.CreateObject("ADODB.Recordset")
rsMenu.ActiveConnection = MM_connDUportal_STRING
rsMenu.Source = "SELECT DISTINCT *  FROM CHANNELS  WHERE CHA_ACTIVE=1  ORDER BY CHA_MENU ASC"
rsMenu.CursorType = 0
rsMenu.CursorLocation = 2
rsMenu.LockType = 1
rsMenu.Open()

rsMenu_numRows = 0
%>
<%
Dim rsMenu__numRows
Dim rsMenu__index

rsMenu__numRows = -1
rsMenu__index = 0
rsMenu_numRows = rsMenu_numRows + rsMenu__numRows
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
                      <td align="left" valign="middle" class="textBoldColor">CHANNELS</td>
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
                <td align="left" valign="top" class="bgTable"> <table width="100%" border="0" cellpadding="3" cellspacing="0">
                    <tr align="left" valign="middle" onmouseover="this.className='bgMouseOver';" onmouseout="this.className='bgMouseOff';"> 
                      <td width="5" ><img src="../assets/icon_cross.gif" hspace="0" vspace="2"></td>
                      <td class="textBoldColor"><a href="../home/">HOME</a></td>
                    </tr>
                    <% 
While ((rsMenu__numRows <> 0) AND (NOT rsMenu.EOF)) 
%>
                    <tr align="left" valign="middle"  onmouseover="this.className='bgMouseOver';" onmouseout="this.className='bgMouseOff';"> 
                      <td><img src="../assets/icon_cross.gif" hspace="0" vspace="2"></td>
                      <td class="textBoldColor"><a href="../home/channel.asp?iChannel=<%=(rsMenu.Fields.Item("CHA_ID").Value)%>&nChannel=<%=(rsMenu.Fields.Item("CHA_NAME").Value)%>"><%= UCase((rsMenu.Fields.Item("CHA_MENU").Value)) %></a></td>
                      <% 
  rsMenu__index=rsMenu__index+1
  rsMenu__numRows=rsMenu__numRows-1
  rsMenu.MoveNext()
Wend
%>
                  </table></td>
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
rsMenu.Close()
Set rsMenu = Nothing
%>