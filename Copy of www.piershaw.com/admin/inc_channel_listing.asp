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
Dim rsChannelsListing
Dim rsChannelsListing_numRows

Set rsChannelsListing = Server.CreateObject("ADODB.Recordset")
rsChannelsListing.ActiveConnection = MM_connDUportal_STRING
rsChannelsListing.Source = "SELECT *, (SELECT COUNT(*) FROM DATAS, CATEGORIES WHERE CAT_ID = DAT_CATEGORY AND CAT_CHANNEL=CHA_ID AND DAT_PARENT=0) AS CHA_COUNT FROM CHANNELS ORDER BY CHA_MENU ASC"
rsChannelsListing.CursorType = 0
rsChannelsListing.CursorLocation = 2
rsChannelsListing.LockType = 1
rsChannelsListing.Open()

rsChannelsListing_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsChannelsListing_numRows = rsChannelsListing_numRows + Repeat1__numRows
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
                      <td align="left" valign="middle" class="textBoldColor">CHANNELS 
                        LISTING</td>
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
                            <td align="left" valign="top"><table width="100%" border="0" cellspacing="2" cellpadding="2">
                                <tr>
                                  <td align="left" valign="top"><table width="100%" border="0" cellpadding="4" cellspacing="1" class="bgTableBorder">
                                      <tr align="left" valign="middle" bgcolor="#CCCCCC" class="textBoldColor"> 
                                        <td width="20" class="textBoldColor">&nbsp;</td>
                                        <td align="center">MENU NAME </td>
                                        <td align="center">DATA NAME</td>
                                        <td width="80" align="center">DATA COUNT</td>
                                        <td width="80" align="center">USER SUBMIT</td>
                                        <td width="80" align="center">MENU ACTIVE</td>
                                        <td width="90" align="center">STATS ACTIVE</td>
                                        <td width="40" align="center">EDIT</td>
                                      </tr>
                                      <% 
While ((Repeat1__numRows <> 0) AND (NOT rsChannelsListing.EOF)) 
%>
                                      <tr align="left" valign="middle" class="bgTable"> 
                                        <td align="center" class="textBoldColor"><img src="../assets/icon_folder.gif" align="absmiddle"></td>
                                        <td class="textBold"> 
                                          <a href="../home/channel.asp?iChannel=<%=(rsChannelsListing.Fields.Item("CHA_ID").Value)%>&nChannel=<%=(rsChannelsListing.Fields.Item("CHA_NAME").Value)%>" target="_blank"><%=(rsChannelsListing.Fields.Item("CHA_MENU").Value)%></a></td>
                                        <td align="left" class="text"><%=(rsChannelsListing.Fields.Item("CHA_NAME").Value)%></td>
                                        <td align="center" class="text"><%=(rsChannelsListing.Fields.Item("CHA_COUNT").Value)%></td>
                                        <td align="center" class="text"><img src="../assets/icon_<%=(rsChannelsListing.Fields.Item("CHA_SUBMIT").Value)%>.gif" width="15" height="15"></td>
                                        <td align="center" class="text"><img src="../assets/icon_<%=(rsChannelsListing.Fields.Item("CHA_ACTIVE").Value)%>.gif" width="15" height="15"></td>
                                        <td align="center" class="text"><img src="../assets/icon_<%=(rsChannelsListing.Fields.Item("CHA_STATS").Value)%>.gif" width="15" height="15"></td>
                                        <td align="center"><a href="channelsEdit.asp?iChannel=<%=(rsChannelsListing.Fields.Item("CHA_ID").Value)%>"><img src="../assets/icon_edit_data.gif" alt="EDIT" border="0"></a></td>
                                      </tr>
                                      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsChannelsListing.MoveNext()
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
rsChannelsListing.Close()
Set rsChannelsListing = Nothing
%>
