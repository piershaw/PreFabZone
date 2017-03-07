<!--#include file="../Connections/connDUportal.asp" -->
<%
set rsHotlinks = Server.CreateObject("ADODB.Recordset")
rsHotlinks.ActiveConnection = MM_connDUportal_STRING
rsHotlinks.Source = "SELECT *, (LINK_RATE/NO_RATES) AS RATING  FROM LINKS  WHERE LINK_APPROVED = Yes  ORDER BY NO_HITS DESC"
rsHotlinks.CursorType = 0
rsHotlinks.CursorLocation = 2
rsHotlinks.LockType = 3
rsHotlinks.Open()
rsHotlinks_numRows = 0
%>
<%
Dim RepeatHotLinks__numRows
RepeatHotLinks__numRows = 5
Dim RepeatHotLinks__index
RepeatHotLinks__index = 0
rsHotlinks_numRows = rsHotlinks_numRows + RepeatHotLinks__numRows
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="middle" class = "bg_navigator" height="20"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;<font size="1">HOT 
      LINKS </font></font></b></td>
    <td align="right" valign="middle" class = "bg_navigator" height="20">&nbsp; </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr> 
    <td align="left" valign="top" colspan="2"> 
      <% 
While ((RepeatHotLinks__numRows <> 0) AND (NOT rsHotlinks.EOF)) 
%>
      <table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr> 
          <td align="left" valign="middle">
            <div class = "links"><img src="../assets/bullet.gif" width="11" height="11" align="absmiddle"> 
              <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font size="1"><a href="../DUdirectory/dirHitting.asp?id=<%=(rsHotlinks.Fields.Item("LINK_ID").Value)%>&url=<%=(rsHotlinks.Fields.Item("LINK_URL").Value)%>" target="_blank" onclick="window.location.reload(true);"><%=(rsHotlinks.Fields.Item("LINK_NAME").Value)%></a></font></b></font> 
              <font size="1"><i>(<%=(rsHotlinks.Fields.Item("LINK_URL").Value)%>)</i></font></div>
          </td>
          <td align="right" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b> Rated:</b> 
            <img src="../assets/<%= FormatNumber((rsHotlinks.Fields.Item("RATING").Value), 1, -2, -2, -2) %>.gif" align="absmiddle"> 
            <b>Hit:</b> <%=(rsHotlinks.Fields.Item("NO_HITS").Value)%></font></td>
        </tr>
        <tr> 
          <td align="left" valign="top" colspan="2"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="14">&nbsp;</td>
                <td align="left" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                  <% =(DoTrimProperly((rsHotlinks.Fields.Item("LINK_DESC").Value), 100, 1, 1, " ...")) %>
                  </font></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <% 
  RepeatHotLinks__index=RepeatHotLinks__index+1
  RepeatHotLinks__numRows=RepeatHotLinks__numRows-1
  rsHotlinks.MoveNext()
Wend
%>
    </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
</table>
<%
rsHotlinks.Close()
%>
