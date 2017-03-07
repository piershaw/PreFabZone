<!--#include file="../Connections/connDUportal.asp" -->
<%
set rsCounter = Server.CreateObject("ADODB.Recordset")
rsCounter.ActiveConnection = MM_connDUportal_STRING
rsCounter.Source = "SELECT DISTINCT  (SELECT COUNT (*) FROM USERS) AS MEMBERS,  (SELECT COUNT(*) FROM LINKS) AS LINKS,  (SELECT COUNT (*) FROM QUESTIONS) AS POLLS,  (SELECT COUNT (*) FROM MESSAGES) AS TOPICS, (SELECT COUNT (*) FROM NEWS) AS NEWS, (SELECT COUNT(*) FROM REPLIES) AS REPLIES  FROM BANNERS"
rsCounter.CursorType = 0
rsCounter.CursorLocation = 2
rsCounter.LockType = 3
rsCounter.Open()
rsCounter_numRows = 0
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="middle" height="20" class = "bg_navigator">&nbsp;<font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>PORTAL 
      COUNTER</b></font></td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr> 
    <td align="left" valign="top"> 
      <table border="0" cellspacing="2" cellpadding="5" width="180">
        <tr align="left" valign="middle"> 
          <td align="left"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">MEMBERS:</font></b></td>
          <td><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsCounter.Fields.Item("MEMBERS").Value)%></font></b></td>
        </tr>
        <tr align="left" valign="middle"> 
          <td align="left"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">LINKS:</font></b></td>
          <td><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsCounter.Fields.Item("LINKS").Value)%></font></b></td>
        </tr>
        <tr align="left" valign="middle"> 
          <td align="left"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">NEWS:</font></b></td>
          <td><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsCounter.Fields.Item("NEWS").Value)%></font></b></td>
        </tr>
        <tr align="left" valign="middle"> 
          <td align="left"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">POLLS:</font></b></td>
          <td><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsCounter.Fields.Item("POLLS").Value)%></font></b></td>
        </tr>
        <tr align="left" valign="middle"> 
          <td align="left"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">TOPICS:</font></b></td>
          <td><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsCounter.Fields.Item("TOPICS").Value)%></font></b></td>
        </tr>
        <tr align="left" valign="middle"> 
          <td align="left"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">REPLIES:</font></b></td>
          <td><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsCounter.Fields.Item("REPLIES").Value)%></font></b></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
</table>
<%
rsCounter.Close()
%>
