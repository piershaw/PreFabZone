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
Dim rsChannels__varChannel
rsChannels__varChannel = "0"
If (Request.QueryString("iChannel") <> "") Then 
  rsChannels__varChannel = Request.QueryString("iChannel")
End If
%>


<%
set rsChannels = Server.CreateObject("ADODB.Recordset")
rsChannels.ActiveConnection = MM_connDUportal_STRING
rsChannels.Source = "SELECT *, (SELECT COUNT(*)  FROM DATAS  WHERE DAT_CATEGORY = CAT_ID AND DAT_APPROVED=1 AND DAT_EXPIRED > DATE()) AS DAT_COUNT  FROM CATEGORIES, CHANNELS WHERE CAT_CHANNEL = CHA_ID AND CAT_CHANNEL = " + Replace(rsChannels__varChannel, "'", "''") + "  ORDER BY CAT_NAME ASC"
rsChannels.CursorType = 0
rsChannels.CursorLocation = 2
rsChannels.LockType = 3
rsChannels.Open()
rsChannels_numRows = 0
%>

<%
Dim rsChannels__numRows
rsChannels__numRows = -2
Dim rsChannels__index
rsChannels__index = 0
rsChannels_numRows = rsChannels_numRows + rsChannels__numRows
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
                        &raquo; <% If Not rsChannels.EOF Or Not rsChannels.BOF Then %>
                        <a href="../home/channel.asp?iChannel=<%=(rsChannels.Fields.Item("CHA_ID").Value)%>&nChannel=<%=(rsChannels.Fields.Item("CHA_NAME").Value)%>"><%=UCASE(rsChannels.Fields.Item("CHA_MENU").Value)%></a> 
                        <% End If ' end Not rsChannels.EOF Or NOT rsChannels.BOF %> </td>
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
                <td align="left" valign="top" class="bgTable"> <% If Not rsChannels.EOF Or Not rsChannels.BOF Then %>
                  <table width="100%" cellpadding="2" cellspacing="2">
                    <%
startrw = 0
endrw = rsChannels__index
numberColumns = 2
numrows = -1
while((numrows <> 0) AND (Not rsChannels.EOF))
	startrw = endrw + 1
	endrw = endrw + numberColumns
 %>
                    <tr align="center" valign="top"> 
                      <%
While ((startrw <= endrw) AND (Not rsChannels.EOF))
%>
                      <td> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr align="left" valign="middle"> 
                            <td width="5" align="center"><img src="../assets/icon_folder.gif" hspace="3" vspace="0" align="absmiddle"></td>
                            <td class="textBoldColor"><a href="type.asp?iCat=<%=(rsChannels.Fields.Item("CAT_ID").Value)%>&iChannel=<%=(rsChannels.Fields.Item("CHA_ID").Value)%>&nChannel=<%=(rsChannels.Fields.Item("CHA_NAME").Value)%>"><%=(rsChannels.Fields.Item("CAT_NAME").Value)%></a> (<%=(rsChannels.Fields.Item("DAT_COUNT").Value)%>)</td>
                          </tr>
                          <tr align="left" valign="middle"> 
                            <td>&nbsp;</td>
                            <td valign="top" class="text"><%=(rsChannels.Fields.Item("CAT_DESCRIPTION").Value)%></td>
                          </tr>
                        </table></td>
                      <%
	startrw = startrw + 1
	rsChannels.MoveNext()
	Wend
	%>
                    </tr>
                    <%
 numrows=numrows-1
 Wend
 %>
                  </table>
                  <% End If ' end Not rsChannels.EOF Or NOT rsChannels.BOF %> </td>
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
rsChannels.Close()
%>