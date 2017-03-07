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
Dim rsRelated__MMColParam
rsRelated__MMColParam = "0"
if (Request.QueryString("iCat") <> "") then rsRelated__MMColParam = Request.QueryString("iCat")
%>
<%
Dim rsRelated__var_id
rsRelated__var_id = "0"
if (Request.QueryString("iData") <> "") then rsRelated__var_id = Request.QueryString("iData")
%>
<%
set rsRelated = Server.CreateObject("ADODB.Recordset")
rsRelated.ActiveConnection = MM_connDUportal_STRING
rsRelated.Source = "SELECT *  FROM DATAS, CATEGORIES, CHANNELS  WHERE CAT_CHANNEL = CHA_ID AND DAT_CATEGORY = CAT_ID AND CHA_ACTIVE = 1 AND DAT_CATEGORY = " + Replace(rsRelated__MMColParam, "'", "''") + " AND DAT_ID <> " + Replace(rsRelated__var_id, "'", "''") + "  AND DAT_APPROVED=1 AND DAT_EXPIRED > DATE() ORDER BY DAT_DATED DESC"
rsRelated.CursorType = 0
rsRelated.CursorLocation = 2
rsRelated.LockType = 3
rsRelated.Open()
rsRelated_numRows = 0
%>

<%
Dim rsRelated__numRows
Dim rsRelated__index

rsRelated__numRows = 15
rsRelated__index = 0
rsRelated_numRows = rsRelated_numRows + rsRelated__numRows
%>
<link href="../assets/DUportal.css" rel="stylesheet" type="text/css">
<% If Not rsRelated.EOF Or Not rsRelated.BOF Then %>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif"></td>
                <td align="left" valign="top" class="bgTable"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td colspan="2" align="left" valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td align="left" valign="middle"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr> 
                                  <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" id="related">
                                      <tr> 
                                        <td height="18" align="left" valign="middle" class="bgHeader"><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#003399">
                                            <tr> 
                                              <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../assets/bg_header.gif">
                                                  <tr> 
                                                    <td width="10"><img src="../assets/header_end_left.gif"></td>
                                                    <td align="left" valign="middle" class="textBoldColor"> 
                                                      <%=UCASE(rsRelated.Fields.Item("CHA_NAME").Value)%> &raquo; <%=UCASE(rsRelated.Fields.Item("CAT_NAME").Value)%> &raquo; RELATED</td>
                                                    <td width="28" align="right" valign="middle"><img src="../assets/header_end_right.gif"></td>
                                                  </tr>
                                                </table></td>
                                            </tr>
                                          </table></td>
                                      </tr>
                                      <tr> 
                                        <td align="left" valign="top"> <% 
While ((rsRelated__numRows <> 0) AND (NOT rsRelated.EOF)) 
%>
                                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                            <tr> 
                                              <td align="left" valign="top"> <table width="100%" border="0" cellspacing="2" cellpadding="2">
                                                  <tr valign="top"> 
                                                    <td colspan="2" align="left" class="text"><b>Name:</b> 
                                                      <a href="detail.asp?iData=<%=(rsRelated.Fields.Item("DAT_ID").Value)%>&iCat=<%=(rsRelated.Fields.Item("DAT_CATEGORY").Value)%>&iChannel=<%=(rsRelated.Fields.Item("CAT_CHANNEL").Value)%>&nChannel=<%=(rsRelated.Fields.Item("CHA_NAME").Value)%>"><%=(rsRelated.Fields.Item("DAT_NAME").Value)%></a>&nbsp;</td>
                                                  </tr>
                                                  <tr valign="top"> 
                                                    <td width="50%" align="left" class="text"><strong>Category:</strong> 
                                                      <%=(rsRelated.Fields.Item("CAT_NAME").Value)%></td>
                                                    <td width="50%" align="left" class="text"><strong>Views:</strong> 
                                                      <%=(rsRelated.Fields.Item("DAT_HITS").Value)%></td>
                                                  </tr>
                                                </table></td>
                                            </tr>
                                            <tr> 
                                              <td align="left" valign="top" bgcolor="#000000"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                                            </tr>
                                          </table>
                                          <% 
  rsRelated__index=rsRelated__index+1
  rsRelated__numRows=rsRelated__numRows-1
  rsRelated.MoveNext()
Wend
%> </td>
                                      </tr>
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
<% End If ' end Not rsRelated.EOF Or NOT rsRelated.BOF %>
<%
rsRelated.Close()
%>
