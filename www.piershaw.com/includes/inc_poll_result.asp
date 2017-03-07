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
Dim rsActivePollResult
Dim rsActivePollResult_numRows

Set rsActivePollResult = Server.CreateObject("ADODB.Recordset")
rsActivePollResult.ActiveConnection = MM_connDUportal_STRING
rsActivePollResult.Source = "SELECT *  FROM DATAS, CATEGORIES, CHANNELS  WHERE DAT_CATEGORY = CAT_ID AND CAT_CHANNEL = CHA_ID AND DAT_ACTIVE = 1 AND DAT_PARENT=0 AND CHA_NAME = 'POLLS'"
rsActivePollResult.CursorType = 0
rsActivePollResult.CursorLocation = 2
rsActivePollResult.LockType = 1
rsActivePollResult.Open()

rsActivePollResult_numRows = 0
%>
<%
Dim rsChoicesResult
Dim rsChoicesResult_numRows

Set rsChoicesResult = Server.CreateObject("ADODB.Recordset")
rsChoicesResult.ActiveConnection = MM_connDUportal_STRING
rsChoicesResult.Source = "SELECT *  FROM DATAS WHERE DAT_PARENT = " & rsActivePollResult.Fields.Item("DAT_ID").Value & " ORDER BY DAT_ID ASC"
rsChoicesResult.CursorType = 0
rsChoicesResult.CursorLocation = 2
rsChoicesResult.LockType = 1
rsChoicesResult.Open()

rsChoicesResult_numRows = 0
%>

<%
Dim rsChoicesResult__numRows
Dim rsChoicesResult__index

rsChoicesResult__numRows = -1
rsChoicesResult__index = 0
rsChoicesResult_numRows = rsChoicesResult_numRows + rsChoicesResult__numRows
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
                      <td align="left" valign="middle" class="textBoldColor">POLL 
                        RESULT</td>
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
                <td align="left" valign="top" class="bgTable"> <table width="100%" border="0" cellspacing="2" cellpadding="2">
                    <tr> 
                      <td align="left" valign="top" class="text"><TABLE width="100%" border="0" cellpadding="0" cellspacing="0">
                          <TR> 
                            <TD  align="left" valign="top"> 
                              <TABLE width="100%" border="0" cellpadding="2" cellspacing="2">
                                <TR> 
                                  <TD colspan="2" align="left" valign="middle" class="textBold"><img src="../assets/icon_folder.gif" width="15" height="13" align="absmiddle"> 
                                    <%=(rsActivePollResult.Fields.Item("DAT_NAME").Value)%> </TD>
                                </TR>
                                <TR> 
                                  <TD width="50%" align="left" valign="middle" class="text">From 
                                    <%=(rsActivePollResult.Fields.Item("DAT_DATED").Value)%> to <%=(rsActivePollResult.Fields.Item("DAT_LAST").Value)%></TD>
                                  <TD width="50%" align="right" valign="middle" class="text">Total 
                                    Votes: <%=(rsActivePollResult.Fields.Item("DAT_COUNT").Value)%></TD>
                                </TR>
                                <TR align="left"> 
                                  <TD colspan="2" valign="middle"> 
                                    <% 
While ((rsChoicesResult__numRows <> 0) AND (NOT rsChoicesResult.EOF)) 
Dim percent, total_parent, total_child
total_parent = rsActivePollResult.Fields.Item("DAT_COUNT").Value
total_child = rsChoicesResult.Fields.Item("DAT_COUNT").Value
If total_parent = 0 or total_child = 0 then
percent = 0
else
percent = (total_child/total_parent)*100
end if
%>
                                    <table width="100%" border="0" cellpadding="0" cellspacing="2">
                                      <tr align="left" valign="middle"> 
                                        <td width="5" rowspan="2">&nbsp;</td>
                                        <td class="text"><strong> <%=(rsChoicesResult.Fields.Item("DAT_NAME").Value)%></strong> (<%=(rsChoicesResult.Fields.Item("DAT_COUNT").Value)%> 
                                          votes)</td>
                                      </tr>
                                      <tr align="left" valign="middle"> 
                                        <td class="text"><img src="../assets/bg_header.gif" width="<%= FormatNumber((percent), 0, -2, -2, -2) %>%" height="16" align="absmiddle"><%= FormatNumber((percent), 0, -2, -2, -2) %>%</td>
                                      </tr>
                                      <tr align="left" valign="middle"> 
                                        <td height="5" colspan="2"><img src="assets/_spacer.gif" width="1" height="1"></td>
                                      </tr>
                                    </table>
                                    <% 
  rsChoicesResult__index=rsChoicesResult__index+1
  rsChoicesResult__numRows=rsChoicesResult__numRows-1
  rsChoicesResult.MoveNext()
Wend
%>
                                  </TD>
                                </TR>
                                <% If Request.QueryString("action") = "voted"  then %>
                                <TR align="left"> 
                                  <TD colspan="2" valign="middle" class="textRed">Your 
                                    vote was not accepted because you have voted 
                                    on this poll before.</TD>
                                </TR>
                                <% End If %>
                              </TABLE></TD>
                          </TR>
                        </TABLE> </td>
                    </tr>
                  </table>
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
rsChoicesResult.Close()
Set rsChoicesResult = Nothing
%>
<%
rsActivePollResult.Close()
Set rsActivePollResult = Nothing
%>
