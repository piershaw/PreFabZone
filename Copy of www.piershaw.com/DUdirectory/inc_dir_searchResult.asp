<!--#include file="../Connections/connDUportal.asp" -->
<%
Dim rsSearchLinks__keyword
rsSearchLinks__keyword = "1"
if (Request.QueryString("key") <> "") then rsSearchLinks__keyword = Request.QueryString("key")
%>
<%
set rsSearchLinks = Server.CreateObject("ADODB.Recordset")
rsSearchLinks.ActiveConnection = MM_connDUportal_STRING
rsSearchLinks.Source = "SELECT *  FROM LINKS  WHERE LINK_APPROVED = Yes and (LINK_NAME LIKE '%" + Replace(rsSearchLinks__keyword, "'", "''") + "%' OR LINK_DESC LIKE '%" + Replace(rsSearchLinks__keyword, "'", "''") + "%')  ORDER BY LINK_DATE DESC"
rsSearchLinks.CursorType = 0
rsSearchLinks.CursorLocation = 2
rsSearchLinks.LockType = 3
rsSearchLinks.Open()
rsSearchLinks_numRows = 0
%>
<%
Dim rsSearchLinks__numRows
rsSearchLinks__numRows = 30
Dim rsSearchLinks__index
rsSearchLinks__index = 0
rsSearchLinks_numRows = rsSearchLinks_numRows + rsSearchLinks__numRows
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="middle" class = "bg_navigator" height="20"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;<font size="1">SEARCH 
      RESULT </font></font></b></td>
    <td align="right" valign="middle" class = "bg_navigator" height="20">&nbsp; </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr> 
    <td align="left" valign="top" colspan="2"> 
      <% 
While ((rsSearchLinks__numRows <> 0) AND (NOT rsSearchLinks.EOF)) 
%>
      <table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr> 
          <td align="left" valign="middle"> 
            <div class = "links"><img src="../assets/bullet.gif" width="11" height="11" align="absmiddle"> 
              <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font size="1"><a href="../DUdirectory/dirHitting.asp?id=<%=(rsSearchLinks.Fields.Item("LINK_ID").Value)%>&url=<%=(rsSearchLinks.Fields.Item("LINK_URL").Value)%>" target="_blank" onclick="window.location.reload(true);"><%=(rsSearchLinks.Fields.Item("LINK_NAME").Value)%></a></font></b></font> <font size="1"><i>(<%=(rsSearchLinks.Fields.Item("LINK_URL").Value)%>)</i></font></div>
          </td>
          <td align="right" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Dated:</b> 
            <%= (rsSearchLinks.Fields.Item("LINK_DATE").Value)%></font></td>
        </tr>
        <tr> 
          <td align="left" valign="top" colspan="2"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="14">&nbsp;</td>
                <td align="left" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                  <% =(DoTrimProperly((rsSearchLinks.Fields.Item("LINK_DESC").Value), 100, 1, 1, " ...")) %>
                  </font></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <% 
  rsSearchLinks__index=rsSearchLinks__index+1
  rsSearchLinks__numRows=rsSearchLinks__numRows-1
  rsSearchLinks.MoveNext()
Wend
%>
    </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
</table>
<%
rsSearchLinks.Close()
%>
