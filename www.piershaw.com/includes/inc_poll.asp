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
Dim rsActivePoll
Dim rsActivePoll_numRows

Set rsActivePoll = Server.CreateObject("ADODB.Recordset")
rsActivePoll.ActiveConnection = MM_connDUportal_STRING
rsActivePoll.Source = "SELECT * FROM DATAS, CATEGORIES, CHANNELS WHERE DAT_CATEGORY = CAT_ID AND CAT_CHANNEL = CHA_ID AND DAT_ACTIVE = 1 AND DAT_PARENT=0 AND CHA_NAME = 'POLLS'"
rsActivePoll.CursorType = 0
rsActivePoll.CursorLocation = 2
rsActivePoll.LockType = 1
rsActivePoll.Open()

rsActivePoll_numRows = 0
%>
<% If Not rsActivePoll.EOF Or Not rsActivePoll.BOF Then %>
<%
Dim rsChoices
Dim rsChoices_numRows

Set rsChoices = Server.CreateObject("ADODB.Recordset")
rsChoices.ActiveConnection = MM_connDUportal_STRING
rsChoices.Source = "SELECT *  FROM DATAS WHERE DAT_PARENT = " & rsActivePoll.Fields.Item("DAT_ID").Value & " ORDER BY DAT_ID ASC"
rsChoices.CursorType = 0
rsChoices.CursorLocation = 2
rsChoices.LockType = 1
rsChoices.Open()

rsChoices_numRows = 0
%>
<%
Dim rsChoices__numRows
Dim rsChoices__index

rsChoices__numRows = -1
rsChoices__index = 0
rsChoices_numRows = rsChoices_numRows + rsChoices__numRows
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
                      <td align="left" valign="middle" class="textBoldColor"> 
                        ACTIVE POLL</td>
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
              
                  <td align="left" valign="top" class="bgTable"><table width="100%" border="0" cellspacing="2" cellpadding="2">
                    <tr> 
                     <form name="poll" method="get" action="../includes/inc_poll_voting.asp">
    <td align="left" valign="top"> 
                          <table width="100%" border="0" cellspacing="2" cellpadding="2">
                            <tr> 
                              <td align="left" valign="top" class="textBold"><%=(rsActivePoll.Fields.Item("DAT_NAME").Value)%> <input name="DAT_PARENT" type="hidden" id="DAT_PARENT" value="<%=(rsActivePoll.Fields.Item("DAT_ID").Value)%>"> 
                                <input name="DAT_CATEGORY" type="hidden" id="DAT_CATEGORY" value="<%=(rsActivePoll.Fields.Item("DAT_CATEGORY").Value)%>"> 
                                <input name="CHA_ID" type="hidden" id="CHA_ID" value="<%=(rsActivePoll.Fields.Item("CHA_ID").Value)%>"> 
                                <input name="CHA_NAME" type="hidden" id="CHA_NAME" value="<%=(rsActivePoll.Fields.Item("CHA_NAME").Value)%>"> 
                              </td>
                            </tr>
                            <% 
While ((rsChoices__numRows <> 0) AND (NOT rsChoices.EOF)) 
%>
                            <tr> 
                              <td align="left" valign="top"><table border="0" cellspacing="0" cellpadding="0">
                                  <tr align="left" valign="middle"> 
                                    <td><input name="DAT_ID" type="radio" value="<%=(rsChoices.Fields.Item("DAT_ID").Value)%>" checked></td>
                                    <td width="5"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                                    <td class="text"><%=(rsChoices.Fields.Item("DAT_NAME").Value)%></td>
                                  </tr>
                                </table></td>
                            </tr>
                            <% 
  rsChoices__index=rsChoices__index+1
  rsChoices__numRows=rsChoices__numRows-1
  rsChoices.MoveNext()
Wend
%>
                            <tr> 
                              <td align="left" valign="top"><input name="Submit" type="submit" class="button" value="Vote"></td>
                            </tr>
                          </table>
                         </td>
                      </form>
                    </tr>
                  </table> </td>
               
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
rsChoices.Close()
Set rsChoices = Nothing
%>
<% End If ' end Not rsActivePoll.EOF Or NOT rsActivePoll.BOF %>
<%
rsActivePoll.Close()
Set rsActivePoll = Nothing
%>
