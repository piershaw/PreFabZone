<!--#include file="../Connections/connDUportal.asp" -->
<%
set rsHotTopic = Server.CreateObject("ADODB.Recordset")
rsHotTopic.ActiveConnection = MM_connDUportal_STRING
rsHotTopic.Source = "SELECT *, FOR_NAME  FROM MESSAGES INNER JOIN FORUMS ON MESSAGES.FOR_ID = FORUMS.FOR_ID  ORDER BY MSG_REPLY_COUNT DESC"
rsHotTopic.CursorType = 0
rsHotTopic.CursorLocation = 2
rsHotTopic.LockType = 3
rsHotTopic.Open()
rsHotTopic_numRows = 0
%>
<%
Dim RepeatHotTopic__numRows
RepeatHotTopic__numRows = 5
Dim RepeatHotTopic__index
RepeatHotTopic__index = 0
rsHotTopic_numRows = rsHotTopic_numRows + RepeatHotTopic__numRows
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
    <td align="left" height="20">&nbsp;<font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font size="1">TOPICS</font></b></font></td>
    <td align="right" valign = "middle">&nbsp; </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr> 
    <td align="left" valign="top" colspan="2"> 
      <% 
While ((RepeatHotTopic__numRows <> 0) AND (NOT rsHotTopic.EOF)) 
%>
      <table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr align="left" valign="middle"> 
          <td> 
            <div class = "links"><img src="../assets/bullet.gif" align="absmiddle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FF0000">&nbsp;<A HREF="../DUforum/msgDetail.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "msg_id=" & rsHotTopic.Fields.Item("MSG_ID").Value & "&" & "for_id=" & rsHotTopic.Fields.Item("FOR_ID").Value %>"><%=(rsHotTopic.Fields.Item("MSG_SUBJECT").Value)%></A></font></div>
          </td>
        </tr>
      </table>
      <% 
  RepeatHotTopic__index=RepeatHotTopic__index+1
  RepeatHotTopic__numRows=RepeatHotTopic__numRows-1
  rsHotTopic.MoveNext()
Wend
%>
    </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
</table>
<%
rsHotTopic.Close()
%>
