<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../Connections/connDUportal.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="admin"
MM_authFailedURL="../DUhome/default.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (false Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<%
set rsAdminBanners = Server.CreateObject("ADODB.Recordset")
rsAdminBanners.ActiveConnection = MM_connDUportal_STRING
rsAdminBanners.Source = "SELECT * FROM BANNERS ORDER BY B_CLICKED_TOTAL DESC"
rsAdminBanners.CursorType = 0
rsAdminBanners.CursorLocation = 2
rsAdminBanners.LockType = 3
rsAdminBanners.Open()
rsAdminBanners_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = 5
Dim Repeat1__index
Repeat1__index = 0
rsAdminBanners_numRows = rsAdminBanners_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
rsAdminBanners_total = rsAdminBanners.RecordCount

' set the number of rows displayed on this page
If (rsAdminBanners_numRows < 0) Then
  rsAdminBanners_numRows = rsAdminBanners_total
Elseif (rsAdminBanners_numRows = 0) Then
  rsAdminBanners_numRows = 1
End If

' set the first and last displayed record
rsAdminBanners_first = 1
rsAdminBanners_last  = rsAdminBanners_first + rsAdminBanners_numRows - 1

' if we have the correct record count, check the other stats
If (rsAdminBanners_total <> -1) Then
  If (rsAdminBanners_first > rsAdminBanners_total) Then rsAdminBanners_first = rsAdminBanners_total
  If (rsAdminBanners_last > rsAdminBanners_total) Then rsAdminBanners_last = rsAdminBanners_total
  If (rsAdminBanners_numRows > rsAdminBanners_total) Then rsAdminBanners_numRows = rsAdminBanners_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsAdminBanners_total = -1) Then

  ' count the total records by iterating through the recordset
  rsAdminBanners_total=0
  While (Not rsAdminBanners.EOF)
    rsAdminBanners_total = rsAdminBanners_total + 1
    rsAdminBanners.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsAdminBanners.CursorType > 0) Then
    rsAdminBanners.MoveFirst
  Else
    rsAdminBanners.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsAdminBanners_numRows < 0 Or rsAdminBanners_numRows > rsAdminBanners_total) Then
    rsAdminBanners_numRows = rsAdminBanners_total
  End If

  ' set the first and last displayed record
  rsAdminBanners_first = 1
  rsAdminBanners_last = rsAdminBanners_first + rsAdminBanners_numRows - 1
  If (rsAdminBanners_first > rsAdminBanners_total) Then rsAdminBanners_first = rsAdminBanners_total
  If (rsAdminBanners_last > rsAdminBanners_total) Then rsAdminBanners_last = rsAdminBanners_total

End If
%>
<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = rsAdminBanners
MM_rsCount   = rsAdminBanners_total
MM_size      = rsAdminBanners_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  r = Request.QueryString("index")
  If r = "" Then r = Request.QueryString("offset")
  If r <> "" Then MM_offset = Int(r)

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  i = 0
  While ((Not MM_rs.EOF) And (i < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    i = i + 1
  Wend
  If (MM_rs.EOF) Then MM_offset = i  ' set MM_offset to the last possible record

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  i = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or i < MM_offset + MM_size))
    MM_rs.MoveNext
    i = i + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = i
    If (MM_size < 0 Or MM_size > MM_rsCount) Then MM_size = MM_rsCount
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  i = 0
  While (Not MM_rs.EOF And i < MM_offset)
    MM_rs.MoveNext
    i = i + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsAdminBanners_first = MM_offset + 1
rsAdminBanners_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (rsAdminBanners_first > MM_rsCount) Then rsAdminBanners_first = MM_rsCount
  If (rsAdminBanners_last > MM_rsCount) Then rsAdminBanners_last = MM_rsCount
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
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
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 0) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    params = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For i = 0 To UBound(params)
      nextItem = Left(params(i), InStr(params(i),"=") - 1)
      If (StrComp(nextItem,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & params(i)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then MM_keepMove = MM_keepMove & "&"
urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="
MM_moveFirst = urlStr & "0"
MM_moveLast  = urlStr & "-1"
MM_moveNext  = urlStr & Cstr(MM_offset + MM_size)
prev = MM_offset - MM_size
If (prev < 0) Then prev = 0
MM_movePrev  = urlStr & Cstr(prev)
%>
<html>
<head>
<title>DUportal</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/default.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="top" height="60" bgcolor="#009999"><img src="../assets/DUportalAdmin.gif" width="268" height="60"></td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr> 
    <td align="left" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="200" align="left" valign="top"> 
            <!--#include file="inc_left.asp" -->
          </td>
          <td width="1" bgcolor="#000000"><img src="../assets/verticalBar.gif" width="1" height="5"></td>
          <td align="left" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td align="left" valign="middle" bgcolor="#00CC99" height="20" colspan="2"> 
                  <div class = "login">&nbsp;<b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="whatsnew.asp">HOME</a> 
                    | <a href="users.asp">USERS</a> | <a href="banners.asp">BANNERS</a> 
                    | <a href="links.asp">LINKS</a> | <a href="forums.asp">FORUMS</a> 
                    | <a href="news.asp">NEWS</a> | <a href="polls.asp">POLLS</a></font></b></div>
                </td>
              </tr>
              <tr> 
                <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              
              
              
              
              <tr bgcolor="#CCCCCC" valign="middle"> 
                <td align="left" height="20"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp; 
                  <a href="bannersEdit.asp"> EDITING BANNERS</a> | <a href="banners.asp">ADD?</a></font></b></td>
                <td align="right" height="20"> <font face="Verdana, Arial, Helvetica, sans-serif"> 
                  <font size="1"> SHOWING 
                  <%
For i = 1 to rsAdminBanners_total Step MM_size
TM_endCount = i + MM_size - 1
if TM_endCount > rsAdminBanners_total Then TM_endCount = rsAdminBanners_total
if i <> MM_offset + 1 Then
Response.Write("<a href=""" & Request.ServerVariables("URL") & "?" & MM_keepMove & "offset=" & i-1 & """>")
Response.Write(i & "-" & TM_endCount & "</a>")
else
Response.Write("<b>" & i & "-" & TM_endCount & "</b>")
End if
if(TM_endCount <> rsAdminBanners_total) then Response.Write(" | ")
next
 %>
                  </font> </font> &nbsp;</td>
              </tr>
              <tr bgcolor="#000000"> 
                <td align="left" valign="top" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="top" colspan="2"> 
                  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsAdminBanners.EOF)) 
%>
                  <table width="100%" border="0" cellspacing="0" cellpadding="4">
                    <tr> 
                      <td align="center" valign="middle"> 
                        <form name="BANNER" method="get" action="bannerEditing.asp">
                          <table width="99%" border="0" cellspacing="0" cellpadding="2">
                            <tr align="left" valign="middle" bgcolor="#FFFFFF"> 
                              <td><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                                Name: 
                                <input type="text" name="B_NAME" size="30" value="<%=(rsAdminBanners.Fields.Item("B_NAME").Value)%>">
                                </font></td>
                              <td><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                                URL 
                                <input type="text" name="B_URL" size="30" value="<%=(rsAdminBanners.Fields.Item("B_URL").Value)%>">
                                </font></td>
                            </tr>
                            <tr align="left" valign="middle" bgcolor="#FFFFFF"> 
                              <td><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                                Image 
                                <input type="text" name="B_IMAGE" size="30" value="<%=(rsAdminBanners.Fields.Item("B_IMAGE").Value)%>">
                                </font></td>
                              <td> <font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                                ALT 
                                <input type="text" name="B_ALT" size="30" value="<%=(rsAdminBanners.Fields.Item("B_ALT").Value)%>">
                                </font></td>
                            </tr>
                            <tr align="center" valign="middle" bgcolor="#FFFFFF"> 
                              <td colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="<%=(rsAdminBanners.Fields.Item("B_URL").Value)%>" target="_blank"><img src="../DUbanner/images/<%=(rsAdminBanners.Fields.Item("B_IMAGE").Value)%>" border="0" alt="<%=(rsAdminBanners.Fields.Item("B_ALT").Value)%>"></a></font></td>
                            </tr>
                            <tr align="left" valign="middle" bgcolor="#FFFFFF"> 
                              <td><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Total 
                                Clicks:</b> <%=(rsAdminBanners.Fields.Item("B_CLICKED_TOTAL").Value)%> <b>From</b> <%=(rsAdminBanners.Fields.Item("B_ADDED_DATE").Value)%> <b>to</b> <%=(rsAdminBanners.Fields.Item("B_CLICKED_DATE").Value)%> 
                                <input type="hidden" name="B_ID" value="<%=(rsAdminBanners.Fields.Item("B_ID").Value)%>">
                                </font></td>
                              <td align="right"> <font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                                <input type="submit" name="Submit" value="Save">
                                <input type="submit" name="Submit" value="Delete">
                                <br>
                                For demo purpose, the buttons above are disable 
                                to prevent DUresources users deleting or making 
                                changes to these real banners.</font></td>
                            </tr>
                          </table>
                        </form>
                      </td>
                    </tr>
                    <tr> 
                      <td align="center" valign="middle"> 
                        <hr size="1" width="100%" color= "black">
                      </td>
                    </tr>
                  </table>
                  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsAdminBanners.MoveNext()
Wend
%>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
<%
rsAdminBanners.Close()
%>
