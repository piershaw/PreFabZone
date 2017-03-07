
<!--#include file="../Connections/connDUportal.asp" -->
<%
Dim rsHaveVoted__varQUEST
rsHaveVoted__varQUEST = "999"
if (Request.QueryString("id")  <> "") then rsHaveVoted__varQUEST = Request.QueryString("id") 
%>
<%
set rsHaveVoted = Server.CreateObject("ADODB.Recordset")
rsHaveVoted.ActiveConnection = MM_connDUportal_STRING
rsHaveVoted.Source = "SELECT QUESTIONS.QUEST_ID, QUESTION, QUEST_DATED, TOTAL_VOTES, ANSWERS, VOTES, LAST_VOTE, ((VOTES/TOTAL_VOTES)*100) AS PER, QUEST_DESCRIPTION  FROM QUESTIONS INNER JOIN ANSWERS ON QUESTIONS.QUEST_ID = ANSWERS.QUEST_ID  WHERE QUESTIONS.QUEST_ID = " + Replace(rsHaveVoted__varQUEST, "'", "''") + ""
rsHaveVoted.CursorType = 0
rsHaveVoted.CursorLocation = 2
rsHaveVoted.LockType = 3
rsHaveVoted.Open()
rsHaveVoted_numRows = 0
%>
<%
Dim rsVoted__MMColParam
rsVoted__MMColParam = "999"
if (Request.Cookies(rsHaveVoted__varQUEST) <> "") then rsVoted__MMColParam = Request.Cookies(rsHaveVoted__varQUEST)
%>
<%
set rsVoted = Server.CreateObject("ADODB.Recordset")
rsVoted.ActiveConnection = MM_connDUportal_STRING
rsVoted.Source = "SELECT ANSWERS FROM ANSWERS WHERE ANS_ID = " + Replace(rsVoted__MMColParam, "'", "''") + ""
rsVoted.CursorType = 0
rsVoted.CursorLocation = 2
rsVoted.LockType = 3
rsVoted.Open()
rsVoted_numRows = 0
%>
<%
Dim RepeatHaveVoted__numRows
RepeatHaveVoted__numRows = -1
Dim RepeatHaveVoted__index
RepeatHaveVoted__index = 0
rsHaveVoted_numRows = rsHaveVoted_numRows + RepeatHaveVoted__numRows
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="left" valign="middle" height="20" class = "bg_navigator"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>&nbsp;YOU 
      HAVE VOTED ON THIS POLL</b></font></td>
  </tr>
  <tr>
    <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr>
    <td align="left" valign="top">
      <TABLE border="0" cellspacing="0" cellpadding="0" width="100%">
        <TR> 
          <TD align="left" valign="top"> 
            <table border="0" cellspacing="3" cellpadding="1" width="100%">
              <tr align="LEFT" valign="MIDDLE"> 
          <td colspan="3"> 
            <font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#009999"><b><a href="pollDetail.asp?id=<%=(rsHaveVoted.Fields.Item("QUEST_ID").Value)%>"><font size="1"><%=(rsHaveVoted.Fields.Item("QUESTION").Value)%></font></a></b></font>
          </td>
        </tr>
        <tr align="LEFT" valign="MIDDLE"> 
          <td colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>From</b> 
            <font color="#FF0000"><%=(rsHaveVoted.Fields.Item("QUEST_DATED").Value)%></font> <b>to</b> <font color="#FF0000"><%=(rsHaveVoted.Fields.Item("LAST_VOTE").Value)%></font></font></td>
          <td align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Number 
            of votes:</b> <font color="#FF0000"><%=(rsHaveVoted.Fields.Item("TOTAL_VOTES").Value)%> votes </font></font></td>
        </tr>
        <tr align="LEFT" valign="MIDDLE"> 
          <td valign="middle" colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Description:</b> 
            <%=(rsHaveVoted.Fields.Item("QUEST_DESCRIPTION").Value)%></font></td>
        </tr>
        <tr align="left" valign="MIDDLE"> 
          <% If Not rsVoted.EOF Or Not rsVoted.BOF Then %>
          <td valign="middle" colspan="3" height="25"><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Your 
            vote was not accepted because you have voted on this poll before (you 
            chose <font color="#FF0000"><%=(rsVoted.Fields.Item("ANSWERS").Value)%></font>). Below is the current the result.</font></td>
          <% End If ' end Not rsVoted.EOF Or NOT rsVoted.BOF %>
        </tr>
        <tr align="right" valign="MIDDLE"> 
          <td valign="middle" colspan="3"> 
            <% 
While ((RepeatHaveVoted__numRows <> 0) AND (NOT rsHaveVoted.EOF)) 
%>
            <table width="100%" border="0" cellspacing="3" cellpadding="2">
              <tr> 
                <td align="right" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><font color="#0000FF"><%=(rsHaveVoted.Fields.Item("ANSWERS").Value)%></font></b></font></td>
                <td align="left" valign="middle" width="60%"> 
                  <table width="<%=(rsHaveVoted.Fields.Item("PER").Value)%>%" border="0" cellspacing="0" cellpadding="0" background="../assets/rainbowBar.gif">
                    <tr> 
                      <td align="right" valign="middle" height="19"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" ><i><b><%= FormatNumber((rsHaveVoted.Fields.Item("PER").Value), 0, -2, -2, -2) %>%</b></i></font></td>
                    </tr>
                  </table>
                </td>
                <td align="right" valign="middle" width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><font color="#0000FF"><%=(rsHaveVoted.Fields.Item("VOTES").Value)%> votes</font><font color="#009900"> </font></b></font></td>
              </tr>
            </table>
            <% 
  RepeatHaveVoted__index=RepeatHaveVoted__index+1
  RepeatHaveVoted__numRows=RepeatHaveVoted__numRows-1
  rsHaveVoted.MoveNext()
Wend
%>
          </td>
        </tr>
      </table>
    </TD>
  </TR>
</TABLE>
</td>
  </tr>
  <tr>
    <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
</table>
<%
rsHaveVoted.Close()
%>
<%
rsVoted.Close()
%>
