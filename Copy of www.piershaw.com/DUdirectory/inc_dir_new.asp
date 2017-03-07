<!--#include file="../Connections/connDUportal.asp" -->
<%
set rsNewLinks = Server.CreateObject("ADODB.Recordset")
rsNewLinks.ActiveConnection = MM_connDUportal_STRING
rsNewLinks.Source = "SELECT * FROM LINKS WHERE LINK_APPROVED = Yes ORDER BY LINK_DATE DESC"
rsNewLinks.CursorType = 0
rsNewLinks.CursorLocation = 2
rsNewLinks.LockType = 3
rsNewLinks.Open()
rsNewLinks_numRows = 0
%>
<%
Dim RepeatNewLinks__numRows
RepeatNewLinks__numRows = 5
Dim RepeatNewLinks__index
RepeatNewLinks__index = 0
rsNewLinks_numRows = rsNewLinks_numRows + RepeatNewLinks__numRows
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="middle" class = "bg_navigator" height="20"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;<font size="1">NEW 
      LINKS </font></font></b></td>
    <td align="right" valign="middle" class = "bg_navigator" height="20">&nbsp; </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr> 
    <td align="left" valign="top" colspan="2"> 
      <% 
While ((RepeatNewLinks__numRows <> 0) AND (NOT rsNewLinks.EOF)) 
%>
      <table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr> 
          <td align="left" valign="middle">
            <div class = "links"><img src="../assets/bullet.gif" width="11" height="11" align="absmiddle"> 
              <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font size="1"><a href="../DUdirectory/dirHitting.asp?id=<%=(rsNewLinks.Fields.Item("LINK_ID").Value)%>&url=<%=(rsNewLinks.Fields.Item("LINK_URL").Value)%>" target="_blank" onclick="window.location.reload(true);"><%=(rsNewLinks.Fields.Item("LINK_NAME").Value)%></a></font></b></font> 
              <font size="1"><i>(<%=(rsNewLinks.Fields.Item("LINK_URL").Value)%>)</i></font></div>
          </td>
          <td align="right" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Dated:</b> <%= (rsNewLinks.Fields.Item("LINK_DATE").Value)%></font></td>
        </tr>
        <tr> 
          <td align="left" valign="top" colspan="2"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="14">&nbsp;</td>
                <td align="left" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><% =(DoTrimProperly((rsNewlinks.Fields.Item("LINK_DESC").Value), 100, 1, 1, " ...")) %></font></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <% 
  RepeatNewLinks__index=RepeatNewLinks__index+1
  RepeatNewLinks__numRows=RepeatNewLinks__numRows-1
  rsNewLinks.MoveNext()
Wend
%>
    </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
</table>
<%
rsNewLinks.Close()
%>
