<!--#include file="../Connections/connDUportal.asp" -->
<%
set rsLinkCat = Server.CreateObject("ADODB.Recordset")
rsLinkCat.ActiveConnection = MM_connDUportal_STRING
rsLinkCat.Source = "SELECT *, (SELECT COUNT (*)  FROM LINKS  WHERE LINKS.CAT_ID = LINK_CATS.CAT_ID AND LINK_APPROVED = Yes)  AS LINK_COUNT  FROM LINK_CATS  ORDER BY CAT_NAME ASC"
rsLinkCat.CursorType = 0
rsLinkCat.CursorLocation = 2
rsLinkCat.LockType = 3
rsLinkCat.Open()
rsLinkCat_numRows = 0
%>
<%
set rsLinkSub = Server.CreateObject("ADODB.Recordset")
rsLinkSub.ActiveConnection = MM_connDUportal_STRING
rsLinkSub.Source = "SELECT *  FROM LINK_SUBS  ORDER BY SUB_NAME ASC"
rsLinkSub.CursorType = 0
rsLinkSub.CursorLocation = 2
rsLinkSub.LockType = 3
rsLinkSub.Open()
rsLinkSub_numRows = 0
%>
<%
Dim HLooperLink__numRows
HLooperLink__numRows = -2
Dim HLooperLink__index
HLooperLink__index = 0
rsLinkCat_numRows = rsLinkCat_numRows + HLooperLink__numRows
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
    <td align="left" valign="middle" class = "bg_navigator" height="20"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;<font size="1">LINKS 
      DIRECTORY</font></font></b></td>
    <td align="right" valign="middle" class = "bg_navigator" height="20">&nbsp; </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr> 
    <td align="center" valign="top" colspan="2"> 
      <table width = "100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="center" valign="middle"> 
            <table width="99%">
              <%
startrw = 0
endrw = HLooperLink__index
numberColumns = 2
numrows = -1
while((numrows <> 0) AND (Not rsLinkCat.EOF))
	startrw = endrw + 1
	endrw = endrw + numberColumns
 %>
              <tr align="center" valign="top"> 
                <%
While ((startrw <= endrw) AND (Not rsLinkCat.EOF))
%>
                <td align="left" valign="middle" width = "50%"> 
                  <table border="0" cellspacing="2" cellpadding="1">
                    <tr align="left" valign="top"> 
                      <td valign="middle"> 
                        <div class = "links"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><a href="../DUdirectory/dirCat.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "id=" & rsLinkCat.Fields.Item("CAT_ID").Value %>"><%=(rsLinkCat.Fields.Item("CAT_NAME").Value)%></a></b> 
                          <i><font size="2">(<%=(rsLinkCat.Fields.Item("LINK_COUNT").Value)%>)</font></i></font><br>
                          <%
Dim RepeatDirectory__numRows
RepeatDirectory__numRows = 3
Dim RepeatDirectory__index
RepeatDirectory__index = 0
rsLinkSub_numRows = rsLinkSub_numRows + RepeatDirectory__numRows
varID = rsLinkCat.Fields.Item("CAT_ID").Value
rsLinkSub.Filter = "CAT_ID = " & varID 
%>
                          <% While ((RepeatDirectory__numRows <> 0) AND (NOT rsLinkSub.EOF)) %>
                          <font face="Verdana, Arial, Helvetica, sans-serif"> 
                          <font size="1"><a href="../DUdirectory/dirSub.asp?catid=<%=(rsLinkSub.Fields.Item("CAT_ID").Value)%>&subid=<%=(rsLinkSub.Fields.Item("SUB_ID").Value)%>"><%=(rsLinkSub.Fields.Item("SUB_NAME").Value)%></a></font></font> 
                          <% If RepeatDirectory__index - 1 > 0 Then %>
                          <font face="Verdana, Arial, Helvetica, sans-serif"> 
                          <i><font size="1">...</font></i></font> 
                          <% Else %>
                          <font face="Verdana, Arial, Helvetica, sans-serif"> 
                          <i><font size="1">, </font></i></font> 
                          <% End If %>
                          <%
  RepeatDirectory__index=RepeatDirectory__index+1
  RepeatDirectory__numRows=RepeatDirectory__numRows-1
  rsLinkSub.MoveNext()
Wend
%>
                        </div>
                      </td>
                    </tr>
                  </table>
                </td>
                <%
	startrw = startrw + 1
	rsLinkCat.MoveNext()
	Wend
	%>
              </tr>
              <%
 numrows=numrows-1
 Wend
 %>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
</table>
<%
rsLinkCat.Close()
%>
<%
rsLinkSub.Close()
%>
