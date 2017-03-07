<!--#include file="../Connections/connDUportal.asp" -->
<%
Dim rsSearchNews__keyword
rsSearchNews__keyword = "1"
if (Request.QueryString("key") <> "") then rsSearchNews__keyword = Request.QueryString("key")
%>
<%
set rsSearchNews = Server.CreateObject("ADODB.Recordset")
rsSearchNews.ActiveConnection = MM_connDUportal_STRING
rsSearchNews.Source = "SELECT *  FROM NEWS  WHERE NEWS_APPROVED = Yes AND (NEWS_TITLE LIKE '%" + Replace(rsSearchNews__keyword, "'", "''") + "%' or NEWS_DESC LIKE '%" + Replace(rsSearchNews__keyword, "'", "''") + "%')  ORDER BY NEWS_DATE DESC"
rsSearchNews.CursorType = 0
rsSearchNews.CursorLocation = 2
rsSearchNews.LockType = 3
rsSearchNews.Open()
rsSearchNews_numRows = 0
%>
<%
Dim rsSearchNews__numRows
rsSearchNews__numRows = 30
Dim rsSearchNews__index
rsSearchNews__index = 0
rsSearchNews_numRows = rsSearchNews_numRows + rsSearchNews__numRows
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
    <td align="left" valign="middle" class = "bg_navigator" height="20"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;<font size="1">SEARCH 
      RESULT</font></font></b></td>
    <td align="right" valign="middle" class = "bg_navigator" height="20">&nbsp; </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr> 
    <td align="left" valign="top" colspan="2"> 
      <% 
While ((rsSearchNews__numRows <> 0) AND (NOT rsSearchNews.EOF)) 
%>
      <table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr> 
          <td align="left" valign="middle"> 
            <div class = "links"><img src="../assets/bullet.gif" width="11" height="11" align="absmiddle"> 
              <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font size="1"><A HREF="newsDetail.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "id=" & rsSearchNews.Fields.Item("NEWS_ID").Value %>"><%=(rsSearchNews.Fields.Item("NEWS_TITLE").Value)%></A></font></b></font> <font size="2"><i>(<%=(rsSearchNews.Fields.Item("NEWS_SOURCE").Value)%>)</i></font></div>
          </td>
          <td align="right" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
            <b>Dated:</b> <%=(rsSearchNews.Fields.Item("NEWS_DATE").Value)%></font></td>
        </tr>
        <tr> 
          <td align="left" valign="top" colspan="2"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="14">&nbsp;</td>
                <td align="left" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                  <% =(DoTrimProperly((rsSearchNews.Fields.Item("NEWS_DESC").Value), 250, 1, 1, " ...")) %>
                  </font></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <% 
  rsSearchNews__index=rsSearchNews__index+1
  rsSearchNews__numRows=rsSearchNews__numRows-1
  rsSearchNews.MoveNext()
Wend
%>
    </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
</table>
<%
rsSearchNews.Close()
%>
