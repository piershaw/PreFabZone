<!--#include file="../Connections/connDUportal.asp" -->
<%
set rsNewsType = Server.CreateObject("ADODB.Recordset")
rsNewsType.ActiveConnection = MM_connDUportal_STRING
rsNewsType.Source = "SELECT *, (SELECT COUNT(*) FROM NEWS WHERE NEWS_TYPE = TYPE_ID) AS NEWS_COUNT  FROM NEWS_TYPES  ORDER BY TYPE_NAME ASC"
rsNewsType.CursorType = 0
rsNewsType.CursorLocation = 2
rsNewsType.LockType = 3
rsNewsType.Open()
rsNewsType_numRows = 0
%>
<%
Dim HLooper1__numRows
HLooper1__numRows = -3
Dim HLooper1__index
HLooper1__index = 0
rsNewsType_numRows = rsNewsType_numRows + HLooper1__numRows
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
    <td align="left" valign="middle" class = "bg_navigator" height="20"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>&nbsp;NEWS</b></font></td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr> 
    <td align="left" valign="top"> 
      <table>
        <%
startrw = 0
endrw = HLooper1__index
numberColumns = 3
numrows = -1
while((numrows <> 0) AND (Not rsNewsType.EOF))
	startrw = endrw + 1
	endrw = endrw + numberColumns
 %>
        <tr align="center" valign="top"> 
          <%
While ((startrw <= endrw) AND (Not rsNewsType.EOF))
%>
          <td align="left" > 
            <table border="0" cellspacing="5" cellpadding="5" width="100%">
              <tr> 
                <td align="left" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><A HREF="news.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "type_id=" & rsNewsType.Fields.Item("TYPE_ID").Value %>"><%=(rsNewsType.Fields.Item("TYPE_NAME").Value)%></A></b> <font size="1"><i>(<%=(rsNewsType.Fields.Item("NEWS_COUNT").Value)%>)</i></font></font></td>
              </tr>
            </table>
          </td>
          <%
	startrw = startrw + 1
	rsNewsType.MoveNext()
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
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
</table>
<%
rsNewsType.Close()
%>
