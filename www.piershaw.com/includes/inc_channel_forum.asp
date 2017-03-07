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
Dim rsChannelsForums__varChannel
rsChannelsForums__varChannel = "0"
If (Request.QueryString("iChannel") <> "") Then 
  rsChannelsForums__varChannel = Request.QueryString("iChannel")
End If
%>


<%
set rsChannelsForums = Server.CreateObject("ADODB.Recordset")
rsChannelsForums.ActiveConnection = MM_connDUportal_STRING
rsChannelsForums.Source = "SELECT * FROM (SELECT *, (SELECT COUNT(*)  FROM DATAS  WHERE DAT_CATEGORY = CAT_ID) AS DAT_COUNT, (SELECT COUNT(*)  FROM DATAS  WHERE DAT_CATEGORY = CAT_ID AND DAT_PARENT=0) AS DAT_TOPIC_COUNT, (SELECT SUM(DAT_HITS)  FROM DATAS  WHERE DAT_CATEGORY = CAT_ID) AS DAT_READ_COUNT, (SELECT MAX(DAT_LAST)  FROM DATAS  WHERE DAT_CATEGORY = CAT_ID) AS DAT_LAST  FROM CATEGORIES, CHANNELS WHERE CAT_CHANNEL = CHA_ID AND CAT_CHANNEL = " + Replace(rsChannelsForums__varChannel, "'", "''") + ")  ORDER BY DAT_LAST DESC"
rsChannelsForums.CursorType = 0
rsChannelsForums.CursorLocation = 2
rsChannelsForums.LockType = 3
rsChannelsForums.Open()
rsChannelsForums_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsChannelsForums_numRows = rsChannelsForums_numRows + Repeat1__numRows
%>
<%
Dim rsChannelsForums__numRows
rsChannelsForums__numRows = -2
Dim rsChannelsForums__index
rsChannelsForums__index = 0
rsChannelsForums_numRows = rsChannelsForums_numRows + rsChannelsForums__numRows
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
                        &raquo; <% If Not rsChannelsForums.EOF Or Not rsChannelsForums.BOF Then %>
                        <a href="../home/channel.asp?iChannel=<%=(rsChannelsForums.Fields.Item("CHA_ID").Value)%>&nChannel=<%=(rsChannelsForums.Fields.Item("CHA_NAME").Value)%>"><%=UCASE(rsChannelsForums.Fields.Item("CHA_MENU").Value)%></a> 
                        <% End If ' end Not rsChannelsForums.EOF Or NOT rsChannelsForums.BOF %> </td>
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
                    <tr align="center" valign="middle" class="textBoldColor"> 
                      <td>FORUMS</td>
                      <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                      <td width="60" >READS</td>
                      <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                      <td width="60" >TOPICS</td>
                      <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                      <td width="60" >POSTS</td>
                      <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                      <td width="80" >LAST POST</td>
                    </tr>
					  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsChannelsForums.EOF)) 
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
                            <td align="left" valign="middle" class="textBoldColor"><a href="type.asp?iCat=<%=(rsChannelsForums.Fields.Item("CAT_ID").Value)%>&iChannel=<%=(rsChannelsForums.Fields.Item("CHA_ID").Value)%>&nChannel=<%=(rsChannelsForums.Fields.Item("CHA_NAME").Value)%>"><%=(rsChannelsForums.Fields.Item("CAT_NAME").Value)%></a></td>
                          </tr>
                          <tr>
                           <td width="5">
                            <td align="left" valign="middle" class="text"><%=(rsChannelsForums.Fields.Item("CAT_DESCRIPTION").Value)%></td>
                          </tr>
                        </table></td>
                      <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                      <td ><%=(rsChannelsForums.Fields.Item("DAT_READ_COUNT").Value)%></td>
                      <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                      <td ><%=(rsChannelsForums.Fields.Item("DAT_TOPIC_COUNT").Value)%></td>
                      <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                      <td ><%=(rsChannelsForums.Fields.Item("DAT_COUNT").Value)%></td>
                      <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                      <td><%=(rsChannelsForums.Fields.Item("DAT_LAST").Value)%></td>
                    </tr>
                    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsChannelsForums.MoveNext()
Wend
%>
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
rsChannelsForums.Close()
%>