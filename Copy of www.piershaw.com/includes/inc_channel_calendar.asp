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
set rsEventsDay = Server.CreateObject("ADODB.Recordset")
rsEventsDay.ActiveConnection = MM_connDUportal_STRING
rsEventsDay.Source = "SELECT * FROM DATAS, CATEGORIES, CHANNELS WHERE DAT_CATEGORY = CAT_ID AND CAT_CHANNEL = CHA_ID AND CHA_NAME = 'EVENTS' ORDER BY DAT_DATED"
rsEventsDay.CursorType = 0
rsEventsDay.CursorLocation = 2
rsEventsDay.LockType = 3
rsEventsDay.Open()
rsEventsDay_numRows = 0
%>
<%
Dim rsEventsDay__numRows
Dim rsEventsDay__index

rsEventsDay__numRows = -1
rsEventsDay__index = 0
rsEventsDay_numRows = rsEventsDay_numRows + rsEventsDay__numRows
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
If Request("iDate") <> "" Then
   iDate = DateValue(Request("iDate"))
Else
   iDate = date
End if

CurrentMonth = Month(iDate)
CurrentMonthName = MonthName(CurrentMonth)
CurrentYear = Year(iDate)

FirstDayDate = DateSerial(CurrentYear, CurrentMonth, 1)
FirstDay = WeekDay(FirstDayDate, 0)
CurrentDay = FirstDayDate

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
                      <td align="left" valign="middle" class="textBoldColor"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr align="center" valign="middle"> 
                            <td width="33%" align = "left" class="bgHeader"> <div class="textBoldColor">&laquo; 
                                <A Href="channel.asp?iChannel=<%=Request.QueryString("iChannel")%>&iDate=<%= Server.URLEncode(DateAdd("m",-1, iDate))%>&nChannel=Events"><%= UCASE(MonthName(Month(DateAdd("m", -1, CurrentDay)))) %></a></div></td>
                            <td width="33%" class="textBoldColor"><%= UCASE(CurrentMonthName & " " & CurrentYear) %></td>
                            <td width="33%" align = "right" class="bgHeader"><div class="textBoldColor"><A Href="channel.asp?iChannel=<%=Request.QueryString("iChannel")%>&iDate=<%= Server.URLEncode(DateAdd("m",1,iDate))%>&nChannel=Events"><%= UCASE(MonthName(Month(DateAdd("m", 1, CurrentDay)))) %></a> 
                                &raquo;</div></td>
                          </tr>
                        </table></td>
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
                      <td align="left" valign="top"><table width="100%" border="0" cellpadding="3" cellspacing="1" class="bgTableBorder">
                          <tr> 
                            <% For DayLoop = 1 to 7 %>
                            <td width="14%"  align="center" valign="middle" class="bgDayListing"><%= WeekDayName(Dayloop, True, 0)%></td>
                            <% Next%>
                          </tr>
                          <tr align="center" valign="middle"> 
                            <%
If FirstDay <> 1 Then
i = FirstDay - 1 
Do While i > 0 
%>
                            <TD height="100" align="right" valign="top" class="bgGrey"><%=Day(CurrentDay - i)%></td>
                            <% 
i = i - 1 
Loop
%>
                            <%
End if
DayCounter = FirstDay
CorrectMonth = True

Do While CorrectMonth = True

If CurrentDay = iDate Then %>
                            <TD height="100" align="right" valign="top" class="bgCurrentDay"> 
                              <% Else %>
                            <TD height="100" align="right" valign="top" class="bgWhite" onmouseover="this.className='bgCurrentDay';" onmouseout="this.className='bgWhite';"> 
                              <% End if %>
                              <%=Day(CurrentDay)%> 
                              <% rsEventsDay.Filter = "DAT_DATED = #" & CurrentDay & "#" %>
                              <% If not rsEventsDay.EOF Or not rsEventsDay.BOF Then %>
                              <% 
While ((rsEventsDay__numRows <> 0) AND (NOT rsEventsDay.EOF)) 
%>
                              <table width="100%" border="0" cellspacing="1" cellpadding="0">
                                <tr> 
                                  <td align="left" valign="middle" class="text"> 
                                    - <a href="detail.asp?iData=<%= (rsEventsDay.Fields.Item("DAT_ID").Value)%>&iCat=<%= (rsEventsDay.Fields.Item("DAT_CATEGORY").Value)%>&iChannel=<%= (rsEventsDay.Fields.Item("CAT_CHANNEL").Value)%>&nChannel=Events"> 
                                    <% =(DoTrimProperly((rsEventsDay.Fields.Item("DAT_NAME").Value), 20, 0, 1, "...")) %>
                                    </a> </td>
                                </tr>
                              </table>
                              <% 
  rsEventsDay__index=rsEventsDay__index+1
  rsEventsDay__numRows=rsEventsDay__numRows-1
  rsEventsDay.MoveNext()
Wend
%>
                              <% End If %>
                            </TD>
                            <%
DayCounter = DayCounter + 1
If DayCounter > 7 then
   DayCounter = 1 %>
                          </TR>
                          <tr align="center" valign="middle"> 
                            <%
End if

CurrentDay = DateAdd("d", 1, CurrentDay)

If Month(CurrentDay) <> CurrentMonth then
   CorrectMonth = False
End if
Loop
%>
                            <% 
If DayCounter <> 1 Then
i = 0
Do While i < 8 - DayCounter
%>
                            <TD height="100" align="right" valign="top" class="bgGrey"><%= i + 1 %></TD>
                            <%
i = i + 1
Loop
End if
%>
                          </TR>
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
