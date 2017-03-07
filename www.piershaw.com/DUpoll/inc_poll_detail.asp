<!--#include file="../Connections/connDUportal.asp" -->
<%
Dim rsPollDetail__id
rsPollDetail__id = "9999"
if (Request.QueryString("id") <> "") then rsPollDetail__id = Request.QueryString("id")
%>
<%
set rsPollDetail = Server.CreateObject("ADODB.Recordset")
rsPollDetail.ActiveConnection = MM_connDUportal_STRING
rsPollDetail.Source = "SELECT *  FROM QUESTIONS INNER JOIN ANSWERS ON QUESTIONS.QUEST_ID = ANSWERS.QUEST_ID  WHERE QUESTIONS.QUEST_ID = " + Replace(rsPollDetail__id, "'", "''") + " ORDER BY ANSWERS"
rsPollDetail.CursorType = 0
rsPollDetail.CursorLocation = 2
rsPollDetail.LockType = 3
rsPollDetail.Open()
rsPollDetail_numRows = 0
%>
<%
Dim RepeatPollDetail__numRows
RepeatPollDetail__numRows = -1
Dim RepeatPollDetail__index
RepeatPollDetail__index = 0
rsPollDetail_numRows = rsPollDetail_numRows + RepeatPollDetail__numRows
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
    <td align="left" height="20">&nbsp;<font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font size="1">ARCHIVED POLL</font></b></font></td>
    <td align="right" valign="middle">&nbsp; </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr align="left"> 
    <td valign="top" colspan="2"> 
      <table border="0" cellspacing="2" cellpadding="2">
        <TR> 
          <FORM name="QUESTION" action="pollVoting.asp" method="get">
            <TD align="left" valign="top"> 
              <TABLE border="0" cellspacing="1" cellpadding="1">
                <TR> 
                  <TD align="left" valign="middle" colspan="2"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsPollDetail.Fields.Item("QUEST_DESCRIPTION").Value)%> <i>(<%=(rsPollDetail.Fields.Item("TOTAL_VOTES").Value)%> votes)</i></font></b></TD>
                </TR>
                <TR> 
                  <TD align="left" valign="middle" colspan="2"> 
                    <div class = "links"> 
                      <INPUT type="hidden" name="QUEST_ID" value="<%=(rsPollDetail.Fields.Item("QUEST_ID").Value)%>">
                      <font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsPollDetail.Fields.Item("QUESTION").Value)%></font> <a href="../DUpoll/pollResult.asp?id=<%=(rsPollDetail.Fields.Item("QUEST_ID").Value)%>"><font face="Verdana, Arial, Helvetica, sans-serif" size="1">(View 
                      Result)</font></a></div>
                  </TD>
                </TR>
                <% 
While ((RepeatPollDetail__numRows <> 0) AND (NOT rsPollDetail.EOF)) 
%>
                <TR> 
                  <TD align="left" valign="middle"> 
                    <INPUT type="radio" name="ANS_ID" value="<%=(rsPollDetail.Fields.Item("ANS_ID").Value)%>" checked>
                    <FONT face="Verdana, Arial, Helvetica, sans-serif" size="2"><I></I></FONT> 
                  </TD>
                  <TD align="left" valign="middle" width = "175"><FONT size="1" face="Verdana, Arial, Helvetica, sans-serif"><I><%=(rsPollDetail.Fields.Item("ANSWERS").Value)%></I></FONT></TD>
                </TR>
                <% 
  RepeatPollDetail__index=RepeatPollDetail__index+1
  RepeatPollDetail__numRows=RepeatPollDetail__numRows-1
  rsPollDetail.MoveNext()
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
rsPollDetail.Close()
%>
