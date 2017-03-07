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
Dim rsAdvertising
Dim rsAdvertising_numRows

Set rsAdvertising = Server.CreateObject("ADODB.Recordset")
rsAdvertising.ActiveConnection = MM_connDUportal_STRING
rsAdvertising.Source = "SELECT * FROM DATAS, CATEGORIES, CHANNELS WHERE DAT_PICTURE <> '' AND CHA_ID = CAT_CHANNEL AND DAT_CATEGORY = CAT_ID AND DAT_EXPIRED > DATE() AND CHA_NAME = 'BANNERS' AND CAT_NAME = 'LEFT BANNER'"
rsAdvertising.CursorType = 3
rsAdvertising.CursorLocation = 2
rsAdvertising.LockType = 1
rsAdvertising.Open()

rsAdvertising_numRows = 0
If Not rsAdvertising.EOF Or Not rsAdvertising.BOF Then
Dim rndMaxAd
rndMaxAd = CInt(rsAdvertising.RecordCount)
rsAdvertising.MoveFirst

Dim rndNumberAd
Randomize Timer
rndNumberAd = Int(RND * rndMaxAd)
rsAdvertising.Move rndNumberAd


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
                        ADVERTISING </td>
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
                <td align="left" valign="top" class="bgTable"><table width="100%" border="0" cellspacing="2" cellpadding="2">
                    <tr> 
                      <td align="center" valign="middle"><table border="1" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td align="left" valign="top"><a href="<%=(rsAdvertising.Fields.Item("DAT_URL").Value)%>" target="_blank"><img src="../pictures/<%= (rsAdvertising.Fields.Item("DAT_PICTURE").Value) %>" alt="<%= (rsAdvertising.Fields.Item("DAT_NAME").Value) %>" border="0"></a></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
                <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif"></td>
              </tr>
            </table>
           </td>
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
 <% End If ' end Not rsAdvertising.EOF Or NOT rsAdvertising.BOF %> 
<%
rsAdvertising.Close()
Set rsAdvertising = Nothing
%>
