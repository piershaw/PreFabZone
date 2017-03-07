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
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="../home/"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
Response.Cookies("DUportalUser").Expires = Date - 300
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
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
set rsNewChannels = Server.CreateObject("ADODB.Recordset")
rsNewChannels.ActiveConnection = MM_connDUportal_STRING
rsNewChannels.Source = "SELECT * FROM CHANNELS WHERE CHA_ACTIVE=1 AND CHA_SUBMIT=1 ORDER BY CHA_NAME ASC"
rsNewChannels.CursorType = 0
rsNewChannels.CursorLocation = 2
rsNewChannels.LockType = 3
rsNewChannels.Open()
rsNewChannels_numRows = 0
%>
<%
set rsNewDatas = Server.CreateObject("ADODB.Recordset")
rsNewDatas.ActiveConnection = MM_connDUportal_STRING
rsNewDatas.Source = "SELECT * FROM DATAS, CATEGORIES WHERE DAT_CATEGORY = CAT_ID AND DAT_USER = '" & Session("MM_Username") & "' ORDER BY DAT_DATED DESC"
rsNewDatas.CursorType = 0
rsNewDatas.CursorLocation = 2
rsNewDatas.LockType = 3
rsNewDatas.Open()
rsNewDatas_numRows = 0
%>
<%
Dim HLooper1__numRows
HLooper1__numRows = -1
Dim HLooper1__index
HLooper1__index = 0
rsNewChannels_numRows = rsNewChannels_numRows + HLooper1__numRows
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
                      <td align="left" valign="middle" class="textBoldColor">MY 
                        PORTAL </td>
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
                <td align="left" valign="top" class="bgTable"><table cellpadding="0" cellspacing="0" width="100%">
                    <%
startrw = 0
endrw = HLooper1__index
numberColumns = 1
numrows = -1
while((numrows <> 0) AND (Not rsNewChannels.EOF))
	startrw = endrw + 1
	endrw = endrw + numberColumns
 %>
                    <tr align="center" valign="top"> 
                      <%
While ((startrw <= endrw) AND (Not rsNewChannels.EOF))
rsNewDatas.Filter = "CAT_CHANNEL = "  & rsNewChannels.Fields.Item("CHA_ID").Value
Dim rsNewDatas__numRows
rsNewDatas__numRows = -1
Dim rsNewDatas__index
rsNewDatas__index = 0
rsNewDatas_numRows = rsNewDatas_numRows + rsNewDatas__numRows
%>
                      
                      <td width="50%" > <table width="100%" border="0" cellspacing="2" cellpadding="2">
                          <tr> 
                            <td align="left" valign="top" class="text" colspan="2"> 
                              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr > 
                                  <td align="left" valign="middle" class="textBold"><img src="../assets/icon_round.gif" hspace="0" vspace="2" align="absmiddle">&nbsp;<%=UCASE(rsNewChannels.Fields.Item("CHA_NAME").Value)%></td>
                                  <td align="right" valign="middle" class="textBold"><a href="../home/submit.asp?iChannel=<%=(rsNewChannels.Fields.Item("CHA_ID").Value)%>&nChannel=<%=(rsNewChannels.Fields.Item("CHA_NAME").Value)%>">ADD 
                                    NEW</a> </td>
                                </tr>
                                <tr> 
                                  <td  colspan="2" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                                </tr>
                                <% 
While ((rsNewDatas__numRows <> 0) AND (NOT rsNewDatas.EOF)) 
%>
                                <tr> 
                                  <td colspan="2"> <table width="100%" border="0" cellspacing="2" cellpadding="2">
                                      <tr valign="middle"> 
									  <td width="1" align="center" class="text"><img src="../assets/icon_<%=(rsNewDatas.Fields.Item("DAT_APPROVED").Value)%>.gif" alt="APPROVED"  border="0" align="absmiddle"></td>
									  
                                        <td align="left" class="text"> 
                                          <a href="../home/detail.asp?iData=<%=(rsNewDatas.Fields.Item("DAT_ID").Value)%>&iCat=<%=(rsNewDatas.Fields.Item("DAT_CATEGORY").Value)%>&iChannel=<%=(rsNewDatas.Fields.Item("CAT_CHANNEL").Value)%>&nChannel=<%=(rsNewChannels.Fields.Item("CHA_NAME").Value)%>"> 
                                          <% =(DoTrimProperly((rsNewDatas.Fields.Item("DAT_NAME").Value), 100, 1, 1, " ...")) %>
                                          </a></td>
										  
                                        <td align="right" width="80" class="text"><%=(rsNewDatas.Fields.Item("DAT_DATED").Value)%></td>
										
                                        <td width="1" align="center" class="text"><a href="../home/edit.asp?iData=<%=(rsNewDatas.Fields.Item("DAT_ID").Value)%>&iCat=<%=(rsNewDatas.Fields.Item("DAT_CATEGORY").Value)%>&iChannel=<%=(rsNewDatas.Fields.Item("CAT_CHANNEL").Value)%>&nChannel=<%=(rsNewChannels.Fields.Item("CHA_NAME").Value)%>"><img src="../assets/icon_edit_data.gif" alt="EDIT" border="0" align="absmiddle"></a></td>
										
                                        <td width="1" align="center" class="text"><a href="../home/delete.asp?iData=<%=(rsNewDatas.Fields.Item("DAT_ID").Value)%>&iCat=<%=(rsNewDatas.Fields.Item("DAT_CATEGORY").Value)%>&iChannel=<%=(rsNewDatas.Fields.Item("CAT_CHANNEL").Value)%>&nChannel=<%=(rsNewChannels.Fields.Item("CHA_NAME").Value)%>"><img src="../assets/icon_delete_data.gif" alt="DELETE" border="0" align="absmiddle"></a></td>
										
										
										
										
                                      </tr>
                                    </table></td>
                                </tr>
                                <% 
  rsNewDatas__index=rsNewDatas__index+1
  rsNewDatas__numRows=rsNewDatas__numRows-1
  rsNewDatas.MoveNext()
Wend
%>
                              </table></td>
                          </tr>
                        </table></td>
                      <%
	startrw = startrw + 1
	rsNewChannels.MoveNext()
	Wend
	%>
                    </tr>
                    <%
 numrows=numrows-1
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
rsNewChannels.Close()
%>
<%
rsNewDatas.Close()
%>