<!--#include file="../Connections/connDUportal.asp" -->
<%
set rsReplyNew = Server.CreateObject("ADODB.Recordset")
rsReplyNew.ActiveConnection = MM_connDUportal_STRING
rsReplyNew.Source = "SELECT *, U_EMAIL  FROM (REPLIES INNER JOIN MESSAGES ON MESSAGES.MSG_ID = REPLIES.MSG_ID) INNER JOIN FORUMS ON FORUMS.FOR_ID = MESSAGES.FOR_ID, USERS  WHERE U_ID = REP_AUTHOR    ORDER BY REP_DATE DESC"
rsReplyNew.CursorType = 0
rsReplyNew.CursorLocation = 2
rsReplyNew.LockType = 3
rsReplyNew.Open()
rsReplyNew_numRows = 0
%>
<%
Dim RepeatLongReplyNew__numRows
RepeatLongReplyNew__numRows = 15
Dim RepeatLongReplyNew__index
RepeatLongReplyNew__index = 0
rsReplyNew_numRows = rsReplyNew_numRows + RepeatLongReplyNew__numRows
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
    <td align="left" valign="middle" class = "bg_navigator" height="20"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;<font size="1">NEW 
      REPLIES</font></font></b></td>
    <td align="right" valign="middle" class = "bg_navigator" height="20">&nbsp; </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr valign="top" align="center"> 
    <td colspan="2" align="left"> 
      <div class = "links"> 
        <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#999999">
          <tr align="center" valign="middle" class = "bg_login"> 
            <td align="left" height="20"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Topic</font></b></td>
            <td width=""><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Forum</font></b></td>
            <td width="60"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Dated</font></b></td>
            <td width="40"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Replies</font></b></td>
            <td width="40"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Reads</font></b></td>
            <td width="130"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Reply 
              Date</font></b></td>
          </tr>
          <% 
While ((RepeatLongReplyNew__numRows <> 0) AND (NOT rsReplyNew.EOF)) 
%>
          <tr align="center" valign="middle" bgcolor="#FFFFFF"> 
            <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><a href="msgDetail.asp?msg_id=<%=(rsReplyNew.Fields.Item("MSG_ID").Value)%>&for_id=<%=(rsReplyNew.Fields.Item("FOR_ID").Value)%>">Re: 
              <%=(rsReplyNew.Fields.Item("MSG_SUBJECT").Value)%></a></b> by <a href="mailto:<%=(rsReplyNew.Fields.Item("U_EMAIL").Value)%>"><%=(rsReplyNew.Fields.Item("REP_AUTHOR").Value)%></a> </font></td>
            <td><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><A HREF="messages.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "for_id=" & rsReplyNew.Fields.Item("FOR_ID").Value %>"><%=(rsReplyNew.Fields.Item("FOR_NAME").Value)%></A></font></td>
            <td><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsReplyNew.Fields.Item("MSG_DATE").Value)%></font></td>
            <td><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsReplyNew.Fields.Item("MSG_REPLY_COUNT").Value)%></font></td>
            <td><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsReplyNew.Fields.Item("MSG_READ_COUNT").Value)%></font></td>
            <td><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsReplyNew.Fields.Item("REP_DATE").Value)%></font></td>
          </tr>
          <% 
  RepeatLongReplyNew__index=RepeatLongReplyNew__index+1
  RepeatLongReplyNew__numRows=RepeatLongReplyNew__numRows-1
  rsReplyNew.MoveNext()
Wend
%>
        </table>
      </div>
    </td>
  </tr>
  <tr valign="top" align="center" bgcolor="#000000"> 
    <td colspan="2" align="left"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
</table>
<%
rsReplyNew.Close()
%>
