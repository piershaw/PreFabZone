<!--#include file="../Connections/connDUportal.asp" -->
<%
set rsShortHotNews = Server.CreateObject("ADODB.Recordset")
rsShortHotNews.ActiveConnection = MM_connDUportal_STRING
rsShortHotNews.Source = "SELECT * FROM NEWS WHERE NEWS_APPROVED = Yes ORDER BY NEWS_HITS DESC"
rsShortHotNews.CursorType = 0
rsShortHotNews.CursorLocation = 2
rsShortHotNews.LockType = 3
rsShortHotNews.Open()
rsShortHotNews_numRows = 0
%>
<%
Dim RepeatShortHotNews__numRows
RepeatShortHotNews__numRows = 5
Dim RepeatShortHotNews__index
RepeatShortHotNews__index = 0
rsShortHotNews_numRows = rsShortHotNews_numRows + RepeatShortHotNews__numRows
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
  <tr valign="middle" class = "bg_navigator"> 
    <td align="left" height="20">&nbsp;<font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font size="1">NEWS</font></b></font></td>
    <td align="right" valign="middle">&nbsp; </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr> 
    <td align="left" valign="top" colspan="2"> 
      <% 
While ((RepeatShortHotNews__numRows <> 0) AND (NOT rsShortHotNews.EOF)) 
%>
      <table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr align="left" valign="middle"> 
          <td> 
            <div class = "links"><img src="../assets/bullet.gif" align="absmiddle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FF0000">&nbsp;<A HREF="../DUnews/newsDetail.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "id=" & rsShortHotNews.Fields.Item("NEWS_ID").Value %>"><%=(rsShortHotNews.Fields.Item("NEWS_TITLE").Value)%></A></font></div>
          </td>
        </tr>
      </table>
      <% 
  RepeatShortHotNews__index=RepeatShortHotNews__index+1
  RepeatShortHotNews__numRows=RepeatShortHotNews__numRows-1
  rsShortHotNews.MoveNext()
Wend
%>
    </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
</table>
<%
rsShortHotNews.Close()
%>
