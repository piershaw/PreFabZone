<%
'****************************************************************************************
'**  Copyright Notice                                                               
'**  Copyright 2003 DUware All Rights Reserved.                                
'**  This program is a commercial software; you can modify (at your own risk) any part of it 
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
'****************************************************************************************
%>
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
                      <td align="left" valign="middle" class="textBoldColor"> 
                        CALENDAR</td>
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
              
                  
                <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="bgTable">
                    <tr> 
                      <td align="left" valign="top"><table width="100%" border="0" cellpadding="2" cellspacing="1">
                          <tr> 
                            <% For DayLoop = 1 to 7 %>
                            <td  align="center" valign="middle" class="textBold"><%= WeekDayName(Dayloop, True, 0)%></td>
                            <% Next%>
                          </tr>
                          <tr align="center" valign="middle"> 
                            <%
If FirstDay <> 1 Then
i = FirstDay - 1 
Do While i > 0 
%>
                            <TD align="right" valign="top" class="textGrey"><%=Day(CurrentDay - i)%></td>
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
                            <TD align="right" valign="top" class="textRed"> 
                              <% Else %>
                            <TD align="right" valign="top" class="text"> 
                              <% End if %>
                              <%=Day(CurrentDay)%> 
                              
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
                            <TD align="right" valign="top" class="textGrey"><%= i + 1 %></TD>
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
