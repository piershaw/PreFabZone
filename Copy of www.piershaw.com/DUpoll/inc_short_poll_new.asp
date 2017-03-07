<!--#include file="../Connections/connDUportal.asp" -->
<%
set rsNewPoll = Server.CreateObject("ADODB.Recordset")
rsNewPoll.ActiveConnection = MM_connDUportal_STRING
rsNewPoll.Source = "SELECT *  FROM QUESTIONS INNER JOIN ANSWERS ON QUESTIONS.QUEST_ID = ANSWERS.QUEST_ID  WHERE QUEST_ACTIVE = 1  ORDER BY QUEST_DATED, ANSWERS"
rsNewPoll.CursorType = 0
rsNewPoll.CursorLocation = 2
rsNewPoll.LockType = 3
rsNewPoll.Open()
rsNewPoll_numRows = 0
%>
<%
Dim RepeatNewPoll__numRows
RepeatNewPoll__numRows = -1
Dim RepeatNewPoll__index
RepeatNewPoll__index = 0
rsNewPoll_numRows = rsNewPoll_numRows + RepeatNewPoll__numRows
%>
<script language="JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr valign="middle" class = "bg_navigator"> 
    <td align="left" height="20">&nbsp;<font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font size="1">POLLS</font></b></font></td>
    <td align="right" valign="middle">&nbsp; </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr> 
    <td align="center" valign="top" colspan="2"> 
      <table border="0" cellspacing="2" cellpadding="2" width="180">
        <TR> 
          <FORM name="QUESTION" action="../DUpoll/pollVoting.asp" method="get">
            <TD align="left" valign="top"> 
              <TABLE border="0" cellspacing="1" cellpadding="1">
                <TR> 
                  <TD align="left" valign="middle" colspan="2"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsNewPoll.Fields.Item("QUEST_DESCRIPTION").Value)%> <i>(<%=(rsNewPoll.Fields.Item("TOTAL_VOTES").Value)%> votes)</i></font></b></TD>
                </TR>
                <TR> 
                  <TD align="left" valign="middle" colspan="2"><div class = "links"> 
                    <INPUT type="hidden" name="QUEST_ID" value="<%=(rsNewPoll.Fields.Item("QUEST_ID").Value)%>">
                    <font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsNewPoll.Fields.Item("QUESTION").Value)%></font> <a href="../DUpoll/pollResult.asp?id=<%=(rsNewPoll.Fields.Item("QUEST_ID").Value)%>"><font face="Verdana, Arial, Helvetica, sans-serif" size="1">(View Result)</font></a></div></TD>
                </TR>
                <% 
While ((RepeatNewPoll__numRows <> 0) AND (NOT rsNewPoll.EOF)) 
%>
                <TR> 
                  <TD align="left" valign="middle"> 
                    <INPUT type="radio" name="ANS_ID" value="<%=(rsNewPoll.Fields.Item("ANS_ID").Value)%>" checked>
                    <FONT face="Verdana, Arial, Helvetica, sans-serif" size="2"><I></I></FONT> 
                  </TD>
                  <TD align="left" valign="middle" width = "175"><FONT size="1" face="Verdana, Arial, Helvetica, sans-serif"><I><%=(rsNewPoll.Fields.Item("ANSWERS").Value)%></I></FONT></TD>
                </TR>
                <% 
  RepeatNewPoll__index=RepeatNewPoll__index+1
  RepeatNewPoll__numRows=RepeatNewPoll__numRows-1
  rsNewPoll.MoveNext()
Wend
%>
                <TR align="right"> 
                  <TD valign="middle" align="center">&nbsp; </TD>
                  <TD valign="middle"> 
                    <input type="submit" name="Submit" value="Vote" class = "buttons">
                   </TD>
                </TR>
              </TABLE>
            </TD>
          </form>
        </TR>
      </TABLE>
    </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
</table>
<%
rsNewPoll.Close()
%>
