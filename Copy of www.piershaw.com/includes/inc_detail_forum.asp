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
if(Request.QueryString("iData") <> "") then cmdHits__varData = Request.QueryString("iData")
%>
<%

set cmdHits = Server.CreateObject("ADODB.Command")
cmdHits.ActiveConnection = MM_connDUportal_STRING
cmdHits.CommandText = "UPDATE DATAS  SET DAT_HITS = DAT_HITS + 1  WHERE DAT_ID = " + Replace(cmdHits__varData, "'", "''") + ""
cmdHits.CommandType = 1
cmdHits.CommandTimeout = 0
cmdHits.Prepared = true
cmdHits.Execute()

%>

<%
Dim rsMsgDetail__MMColParam
rsMsgDetail__MMColParam = "0"
if (Request.QueryString("iData") <> "") then rsMsgDetail__MMColParam = Request.QueryString("iData")
%>
<%
set rsMsgDetail = Server.CreateObject("ADODB.Recordset")
rsMsgDetail.ActiveConnection = MM_connDUportal_STRING
rsMsgDetail.Source = "SELECT *, (SELECT COUNT(*) FROM DATAS, CATEGORIES, CHANNELS WHERE DAT_CATEGORY = CAT_ID AND CAT_CHANNEL = CHA_ID AND  DAT_USER = U_ID AND CHA_NAME ='TOPICS') AS POST_COUNT  FROM DATAS,  CATEGORIES, CHANNELS, USERS  WHERE DAT_ID = " + Replace(rsMsgDetail__MMColParam, "'", "''") + " AND DAT_CATEGORY = CAT_ID AND CAT_CHANNEL = CHA_ID AND U_ID = DAT_USER"
rsMsgDetail.CursorType = 0
rsMsgDetail.CursorLocation = 2
rsMsgDetail.LockType = 3
rsMsgDetail.Open()
rsMsgDetail_numRows = 0
%>


<%
Dim rsRepDetail__MMColParam
rsRepDetail__MMColParam = "0"
if (Request.QueryString("iData") <> "") then rsRepDetail__MMColParam = Request.QueryString("iData")
%>
<%
set rsRepDetail = Server.CreateObject("ADODB.Recordset")
rsRepDetail.ActiveConnection = MM_connDUportal_STRING
rsRepDetail.Source = "SELECT *, (SELECT COUNT(*) FROM DATAS, CATEGORIES, CHANNELS WHERE DAT_CATEGORY = CAT_ID AND CAT_CHANNEL = CHA_ID AND  DAT_USER = U_ID AND CHA_NAME ='TOPICS') AS POST_COUNT  FROM DATAS,  CATEGORIES, CHANNELS, USERS  WHERE DAT_PARENT = " + Replace(rsMsgDetail__MMColParam, "'", "''") + " AND DAT_CATEGORY = CAT_ID AND CAT_CHANNEL = CHA_ID AND U_ID = DAT_USER"
rsRepDetail.CursorType = 0
rsRepDetail.CursorLocation = 2
rsRepDetail.LockType = 3
rsRepDetail.Open()
rsRepDetail_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsRepDetail_numRows = rsRepDetail_numRows + Repeat1__numRows
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
                      <td align="left" valign="middle" class="textBoldColor"><a href="default.asp">HOME 
                        </a> &raquo; <a href="channel.asp?iChannel=<%=(rsMsgDetail.Fields.Item("CAT_CHANNEL").Value)%>&nChannel=<%=(rsMsgDetail.Fields.Item("CHA_NAME").Value)%>"><%=UCASE(rsMsgDetail.Fields.Item("CHA_NAME").Value)%></a> 
                        &raquo; <a href="type.asp?iCat=<%=(rsMsgDetail.Fields.Item("DAT_CATEGORY").Value)%>&iChannel=<%=(rsMsgDetail.Fields.Item("CAT_CHANNEL").Value)%>&nChannel=<%=(rsMsgDetail.Fields.Item("CHA_NAME").Value)%>"><%=UCASE(rsMsgDetail.Fields.Item("CAT_NAME").Value)%></a> 
                        &raquo; DETAIL</td>
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
                      <td colspan="2" align="left" valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                         
                          <tr> 
                            <td align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr class="textBold"> 
                                  <td width="110" height="20" align="center" valign="middle">Author</td>
                                  <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                                  <td height="20">&nbsp;<%=(rsMsgDetail.Fields.Item("DAT_NAME").Value)%></td>
                                </tr>
                                <tr> 
                                  <td height="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                                  <td height="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                                  <td height="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                                </tr>
                                <tr> 
                                  <td align="left" valign="top"><table width="100%" border="0" cellspacing="2" cellpadding="2">
                                      <tr> 
                                        <td align="left" valign="middle" class="textBold"><%=(rsMsgDetail.Fields.Item("DAT_USER").Value)%></td>
                                      </tr>
                                      <tr> 
                                        <td align="left" valign="middle" class="text"><strong>From:</strong> 
                                          <%=(rsMsgDetail.Fields.Item("U_COUNTRY").Value)%> </td>
                                      </tr>
                                      <tr> 
                                        <td align="left" valign="middle" class="text"><strong>Posts:</strong> 
                                          <%=(rsMsgDetail.Fields.Item("POST_COUNT").Value)%> </td>
                                      </tr>
                                      <tr>
                                        <td align="left" valign="middle" class="text"><strong>Since:</strong> 
                                          <%=(rsMsgDetail.Fields.Item("U_DATED").Value)%> 
                                        </td>
                                      </tr>
                                    </table></td>
                                  <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                                  <td align="left" valign="top"><table width="100%" border="0" cellspacing="2" cellpadding="2">
                                      <tr> 
                                        <td align="left" valign="middle" class="text"><font color="#999999">Posted 
                                          on <%=(rsMsgDetail.Fields.Item("DAT_DATED").Value)%></font></td>
                                        <td align="right" valign="middle" class="text"><table border="0" cellspacing="2" cellpadding="2">
                                            <tr align="center" valign="middle">
                                              <td><a href="MAILTO: <%=(rsMsgDetail.Fields.Item("U_EMAIL").Value)%>"><img src="../assets/icon_email.gif" alt="EMAIL <%=UCASE(rsMsgDetail.Fields.Item("U_ID").Value)%>" hspace="5" vspace="0" border="0" align="absmiddle"></a></td>
                                              <td><a href="../home/post.asp?iData=<%=(rsMsgDetail.Fields.Item("DAT_ID").Value)%>&iCat=<%=(rsMsgDetail.Fields.Item("CAT_ID").Value)%>&iChannel=<%=(rsMsgDetail.Fields.Item("CHA_ID").Value)%>"><img src="../assets/icon_topic_reply.gif" alt="REPLY THIS TOPIC" hspace="5" vspace="0" border="0" align="absmiddle"></a></td>
                                              <td><a href="../home/post.asp?iData=0&iCat=<%=(rsMsgDetail.Fields.Item("CAT_ID").Value)%>&iChannel=<%=(rsMsgDetail.Fields.Item("CHA_ID").Value)%>"><img src="../assets/icon_topic_new.gif" alt="NEW TOPIC" hspace="5" vspace="0" border="0" align="absmiddle"></a></td>
                                            </tr>
                                          </table></td>
                                      </tr>
                                      <tr> 
                                        <td colspan="2" align="left" valign="top" class="text"><%=TrimBody(rsMsgDetail.Fields.Item("DAT_DESCRIPTION").Value)%></td>
                                      </tr>
                                    </table></td>
                                </tr>
								<% 
While ((Repeat1__numRows <> 0) AND (NOT rsRepDetail.EOF)) 
%>
                                <tr> 
                                  <td height="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                                  <td height="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                                  <td height="1" class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                                </tr>
                                
                                <tr> 
                                  <td align="left" valign="top"><table width="100%" border="0" cellspacing="2" cellpadding="2">
                                      <tr> 
                                        <td align="left" valign="middle" class="textBold"><%=(rsRepDetail.Fields.Item("DAT_USER").Value)%></td>
                                      </tr>
                                      <tr> 
                                        <td align="left" valign="middle" class="text"><strong>From:</strong> 
                                          <%=(rsRepDetail.Fields.Item("U_COUNTRY").Value)%> </td>
                                      </tr>
                                      <tr> 
                                        <td align="left" valign="middle" class="text"><strong>Posts:</strong> 
                                          <%=(rsRepDetail.Fields.Item("POST_COUNT").Value)%> </td>
                                      </tr>
                                      <tr> 
                                        <td align="left" valign="middle" class="text"><strong>Since:</strong> 
                                          <%=(rsRepDetail.Fields.Item("U_DATED").Value)%> </td>
                                      </tr>
                                    </table></td>
                                  <td class="bgTableBorder"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                                  <td align="left" valign="top"><table width="100%" border="0" cellspacing="2" cellpadding="2">
                                      <tr> 
                                        <td align="left" valign="middle" class="text"><font color="#999999">Replied 
                                          on <%=(rsRepDetail.Fields.Item("DAT_DATED").Value)%></font></td>
                                        <td align="right" valign="middle" class="text"><table border="0" cellspacing="2" cellpadding="2">
                                            <tr align="center" valign="middle"> 
                                              <td><a href="MAILTO: <%=(rsRepDetail.Fields.Item("U_EMAIL").Value)%>"><img src="../assets/icon_email.gif" alt="EMAIL <%=UCASE(rsRepDetail.Fields.Item("U_ID").Value)%>" hspace="5" vspace="0" border="0" align="absmiddle"></a></td>
                                              <td><a href="../home/post.asp?iData=<%=(rsRepDetail.Fields.Item("DAT_PARENT").Value)%>&iCat=<%=(rsRepDetail.Fields.Item("CAT_ID").Value)%>&iChannel=<%=(rsRepDetail.Fields.Item("CHA_ID").Value)%>"><img src="../assets/icon_topic_reply.gif" alt="REPLY THIS TOPIC" hspace="5" vspace="0" border="0" align="absmiddle"></a></td>
                                              <td><a href="../home/post.asp?iData=0&iCat=<%=(rsRepDetail.Fields.Item("CAT_ID").Value)%>&iChannel=<%=(rsRepDetail.Fields.Item("CHA_ID").Value)%>"><img src="../assets/icon_topic_new.gif" alt="NEW TOPIC" hspace="5" vspace="0" border="0" align="absmiddle"></a></td>
                                            </tr>
                                          </table></td>
                                      </tr>
                                      <tr> 
                                        <td colspan="2" align="left" valign="top" class="text"><%=TrimBody(rsRepDetail.Fields.Item("DAT_DESCRIPTION").Value)%></td>
                                      </tr>
                                    </table></td>
                                </tr>
                                <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsRepDetail.MoveNext()
Wend
%>
                              </table> </td>
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
rsMsgDetail.Close()
%>

<%
rsRepDetail.Close()
%>