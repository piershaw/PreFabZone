<!--#include file="../Connections/connDUportal.asp" -->
<%
set rsLongHotNews = Server.CreateObject("ADODB.Recordset")
rsLongHotNews.ActiveConnection = MM_connDUportal_STRING
rsLongHotNews.Source = "SELECT * FROM NEWS WHERE NEWS_APPROVED = Yes ORDER BY NEWS_HITS DESC"
rsLongHotNews.CursorType = 0
rsLongHotNews.CursorLocation = 2
rsLongHotNews.LockType = 3
rsLongHotNews.Open()
rsLongHotNews_numRows = 0
%>
<%
Dim RepeatLongHotNews__numRows
RepeatLongHotNews__numRows = 5
Dim RepeatLongHotNews__index
RepeatLongHotNews__index = 0
rsLongHotNews_numRows = rsLongHotNews_numRows + RepeatLongHotNews__numRows
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  
  <tr> 
    <td align="left" valign="middle" class = "bg_navigator" height="20"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;<font size="1">HOT 
      NEWS </font></font></b></td>
    <td align="right" valign="middle" class = "bg_navigator" height="20">&nbsp; </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr> 
    <td align="left" valign="top" colspan="2"> 
      <% 
While ((RepeatLongHotNews__numRows <> 0) AND (NOT rsLongHotNews.EOF)) 
%>
      <table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr> 
          <td align="left" valign="middle"> 
            <div class = "links"><img src="../assets/bullet.gif" width="11" height="11" align="absmiddle"> 
              <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font size="1"><A HREF="newsDetail.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "id=" & rsLongHotNews.Fields.Item("NEWS_ID").Value %>"><%=(rsLongHotNews.Fields.Item("NEWS_TITLE").Value)%></A></font></b></font> <font size="2"><i>(<%=(rsLongHotNews.Fields.Item("NEWS_SOURCE").Value)%>)</i></font></div>
          </td>
          <td align="right" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
            <b>Read:</b> <%=(rsLongHotNews.Fields.Item("NEWS_HITS").Value)%></font></td>
        </tr>
        <tr> 
          <td align="left" valign="top" colspan="2"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="14">&nbsp;</td>
                <td align="left" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                  <% =(DoTrimProperly((rsLongHotNews.Fields.Item("NEWS_DESC").Value), 250, 1, 1, " ...")) %>
                  </font></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <% 
  RepeatLongHotNews__index=RepeatLongHotNews__index+1
  RepeatLongHotNews__numRows=RepeatLongHotNews__numRows-1
  rsLongHotNews.MoveNext()
Wend
%>
    </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
</table>
<%
rsLongHotNews.Close()
%>
