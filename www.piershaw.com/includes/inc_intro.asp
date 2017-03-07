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
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>
function Thumbnail(tmb_suff,tmb_filename)
  Dim tmb_NewFilename, tmb_Path, tmb_PosPath, tmb_PosExt
  if not isnull(tmb_filename) then
    tmb_PosPath = InStrRev(tmb_filename,"/")
    tmb_Path = ""
    if tmb_PosPath > 0 then
      tmb_Path = mid(tmb_filename,1,tmb_PosPath)
    end if
    tmb_PosExt = InStrRev(tmb_filename,".")
    if tmb_PosExt > 0 then
      tmb_NewFilename = tmb_Path & mid(tmb_filename,tmb_PosPath+1,tmb_PosExt-(tmb_PosPath+1)) & tmb_suff & ".jpg"
    else
      tmb_NewFilename = tmb_Path & mid(tmb_filename,tmb_PosPath+1,len(tmb_filename)-tmb_PosPath) & tmb_suff & ".jpg"
    end if
  end if
  Thumbnail = tmb_NewFilename
end function
</SCRIPT>
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
<%
Dim rsLastestNews
Dim rsLastestNews_numRows

Set rsLastestNews = Server.CreateObject("ADODB.Recordset")
rsLastestNews.ActiveConnection = MM_connDUportal_STRING
rsLastestNews.Source = "SELECT * FROM DATAS, CHANNELS, CATEGORIES WHERE CAT_CHANNEL = CHA_ID AND DAT_CATEGORY = CAT_ID AND DAT_APPROVED=1 AND CHA_ACTIVE = 1 AND CHA_NAME = 'NEWS' ORDER BY DAT_DATED DESC"
rsLastestNews.CursorType = 0
rsLastestNews.CursorLocation = 2
rsLastestNews.LockType = 1
rsLastestNews.Open()

rsLastestNews_numRows = 0
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
                      <td align="left" valign="middle" class="textBoldColor">LATEST 
                        NEWS </td>
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
                <td align="left" valign="top" class="bgTable"> <% If Not rsLastestNews.EOF Or Not rsLastestNews.BOF Then %>
                  <table width="100%" border="0" cellspacing="2" cellpadding="2">
                    <tr> 
                      <td align="left" valign="top" class="text"><a href="detail.asp?iData=<%=(rsLastestNews.Fields.Item("DAT_ID").Value)%>&iCat=<%=(rsLastestNews.Fields.Item("DAT_CATEGORY").Value)%>&iChannel=<%=(rsLastestNews.Fields.Item("CAT_CHANNEL").Value)%>&nChannel=<%=(rsLastestNews.Fields.Item("CHA_NAME").Value)%>"><img src="../pictures/<%= Thumbnail("_small",(rsLastestNews.Fields.Item("DAT_PICTURE").Value)) %>" border="1" align="right" alt="<%=(rsLastestNews.Fields.Item("DAT_NAME").Value)%>"></a><strong class="textBoldColor"><a href="detail.asp?iData=<%=(rsLastestNews.Fields.Item("DAT_ID").Value)%>&iCat=<%=(rsLastestNews.Fields.Item("DAT_CATEGORY").Value)%>&iChannel=<%=(rsLastestNews.Fields.Item("CAT_CHANNEL").Value)%>&nChannel=<%=(rsLastestNews.Fields.Item("CHA_NAME").Value)%>"><%=(rsLastestNews.Fields.Item("DAT_NAME").Value)%></a></strong><br> <% =TrimBody(DoTrimProperly((rsLastestNews.Fields.Item("DAT_DESCRIPTION").Value), 600, 1, 1, " ... ")) %> </td>
                    </tr>
                  </table>
                  <% End If ' end Not rsLastestNews.EOF Or NOT rsLastestNews.BOF %> </td>
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
rsLastestNews.Close()
Set rsLastestNews = Nothing
%>