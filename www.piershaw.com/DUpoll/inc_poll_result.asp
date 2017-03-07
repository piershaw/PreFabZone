<!--#include file="../Connections/connDUportal.asp" -->
<%
Dim rsPollResult__varQUEST
rsPollResult__varQUEST = "999"
if (Request.QueryString("id")  <> "") then rsPollResult__varQUEST = Request.QueryString("id") 
%>
<%
set rsPollResult = Server.CreateObject("ADODB.Recordset")
rsPollResult.ActiveConnection = MM_connDUportal_STRING
rsPollResult.Source = "SELECT QUESTIONS.QUEST_ID, QUESTION, QUEST_DATED, TOTAL_VOTES, ANSWERS, VOTES, LAST_VOTE, ((VOTES/TOTAL_VOTES)*100) AS PER, QUEST_DESCRIPTION  FROM QUESTIONS INNER JOIN ANSWERS ON QUESTIONS.QUEST_ID = ANSWERS.QUEST_ID  WHERE QUESTIONS.QUEST_ID = " + Replace(rsPollResult__varQUEST, "'", "''") + ""
rsPollResult.CursorType = 0
rsPollResult.CursorLocation = 2
rsPollResult.LockType = 3
rsPollResult.Open()
rsPollResult_numRows = 0
%>
<%
Dim RepeatPollResult__numRows
RepeatPollResult__numRows = -1
Dim RepeatPollResult__index
RepeatPollResult__index = 0
rsPollResult_numRows = rsPollResult_numRows + RepeatPollResult__numRows
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="middle" align="left" height="20" class = "bg_navigator"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>&nbsp;POLL 
      RESULT</b></font></td>
  </tr>
  <tr>
    <td valign="top" align="left" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr>
    <td valign="top" align="left">
      <TABLE border="0" cellpadding="0" width="100%" cellspacing="0">
        <TR> 
          <TD align="left" valign="top"> 
            <table border="0" cellspacing="3" cellpadding="1" width="100%">
              <tr align="LEFT" valign="MIDDLE"> 
          <td colspan="3"> 
            <font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#009999"><b><a href="pollDetail.asp?id=<%=(rsPollResult.Fields.Item("QUEST_ID").Value)%>"><font size="1" ><%=(rsPollResult.Fields.Item("QUESTION").Value)%></font></a></b></font>
          </td>
        </tr>
        <tr align="LEFT" valign="MIDDLE"> 
          <td colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>From</b> 
            <font color="#FF0000"><%=(rsPollResult.Fields.Item("QUEST_DATED").Value)%></font> <b>to</b> <font color="#FF0000"><%=(rsPollResult.Fields.Item("LAST_VOTE").Value)%></font></font></td>
          <td align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Number 
            of votes:</b> <font color="#FF0000"><%=(rsPollResult.Fields.Item("TOTAL_VOTES").Value)%> votes </font></font></td>
        </tr>
        <tr align="LEFT" valign="MIDDLE"> 
          <td valign="middle" colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Description:</b> 
            <%=(rsPollResult.Fields.Item("QUEST_DESCRIPTION").Value)%></font></td>
        </tr>
        <tr align="right" valign="MIDDLE"> 
          <td valign="middle" colspan="3"> 
            <% 
While ((RepeatPollResult__numRows <> 0) AND (NOT rsPollResult.EOF)) 
%>
            <table width="100%" border="0" cellspacing="3" cellpadding="2">
              <tr> 
                <td align="right" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><font color="#0000FF"><%=(rsPollResult.Fields.Item("ANSWERS").Value)%></font></b></font></td>
                <td align="left" valign="middle" width="60%"> 
                  <table width="<%=(rsPollResult.Fields.Item("PER").Value)%>%" border="0" cellspacing="0" cellpadding="0" background="../assets/rainbowBar.gif">
                    <tr> 
                      <td align="right" valign="middle" height="19"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" ><i><b><%= FormatNumber((rsPollResult.Fields.Item("PER").Value), 0, -2, -2, -2) %>%</b></i></font></td>
                    </tr>
                  </table>
                </td>
                <td align="right" valign="middle" width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><font color="#0000FF"><%=(rsPollResult.Fields.Item("VOTES").Value)%> votes</font><font color="#009900"> </font></b></font></td>
              </tr>
            </table>
            <% 
  RepeatPollResult__index=RepeatPollResult__index+1
  RepeatPollResult__numRows=RepeatPollResult__numRows-1
  rsPollResult.MoveNext()
Wend
%>
            </td>
        </tr>
      </table>
    </TD>
  </TR>
</TABLE></td>
  </tr>
  <tr>
    <td valign="top" align="left" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
</table>
<%
rsPollResult.Close()
%>
