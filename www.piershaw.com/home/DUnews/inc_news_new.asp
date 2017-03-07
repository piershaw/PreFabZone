<!--#include file="../Connections/connDUportal.asp" -->
<%
set rsLongNewNews = Server.CreateObject("ADODB.Recordset")
rsLongNewNews.ActiveConnection = MM_connDUportal_STRING
rsLongNewNews.Source = "SELECT * FROM NEWS WHERE NEWS_APPROVED = Yes ORDER BY NEWS_DATE DESC"
rsLongNewNews.CursorType = 0
rsLongNewNews.CursorLocation = 2
rsLongNewNews.LockType = 3
rsLongNewNews.Open()
rsLongNewNews_numRows = 0
%>
<%
Dim RepeatLongNewNews__numRows
RepeatLongNewNews__numRows = 5
Dim RepeatLongNewNews__index
RepeatLongNewNews__index = 0
rsLongNewNews_numRows = rsLongNewNews_numRows + RepeatLongNewNews__numRows
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
    <td align="left" valign="middle" class = "bg_navigator" height="20"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;<font size="1">LATEST 
      NEWS </font></font></b></td>
    <td align="right" valign="middle" class = "bg_navigator" height="20">&nbsp; </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr> 
    <td align="left" valign="top" colspan="2"> 
      <% 
While ((RepeatLongNewNews__numRows <> 0) AND (NOT rsLongNewNews.EOF)) 
%>
      <table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr> 
          <td align="left" valign="middle"> 
            <div class = "links"><img src="../assets/bullet.gif" width="11" height="11" align="absmiddle"> 
              <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font size="1"><A HREF="newsDetail.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "id=" & rsLongNewNews.Fields.Item("NEWS_ID").Value %>"><%=(rsLongNewNews.Fields.Item("NEWS_TITLE").Value)%></A></font></b></font> <font size="2"><i>(<%=(rsLongNewNews.Fields.Item("NEWS_SOURCE").Value)%>)</i></font></div>
          </td>
          <td align="right" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
            <b>Dated:</b> <%=(rsLongNewNews.Fields.Item("NEWS_DATE").Value)%></font></td>
        </tr>
        <tr> 
          <td align="left" valign="top" colspan="2"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="14">&nbsp;</td>
                <td align="left" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                  <% =(DoTrimProperly((rsLongNewNews.Fields.Item("NEWS_DESC").Value), 250, 1, 1, " ...")) %>
                  </font></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <% 
  RepeatLongNewNews__index=RepeatLongNewNews__index+1
  RepeatLongNewNews__numRows=RepeatLongNewNews__numRows-1
  rsLongNewNews.MoveNext()
Wend
%>
    </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
</table>
<%
rsLongNewNews.Close()
%>
