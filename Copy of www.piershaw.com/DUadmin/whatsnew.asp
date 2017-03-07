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
Dim rsAdminNewLinks__MMColParam
rsAdminNewLinks__MMColParam = "No"
if (Request("MM_EmptyValue") <> "") then rsAdminNewLinks__MMColParam = Request("MM_EmptyValue")
%>
<%
set rsAdminNewLinks = Server.CreateObject("ADODB.Recordset")
rsAdminNewLinks.ActiveConnection = MM_connDUportal_STRING
rsAdminNewLinks.Source = "SELECT * FROM LINKS WHERE LINK_APPROVED = " + Replace(rsAdminNewLinks__MMColParam, "'", "''") + " ORDER BY LINK_ID ASC"
rsAdminNewLinks.CursorType = 0
rsAdminNewLinks.CursorLocation = 2
rsAdminNewLinks.LockType = 3
rsAdminNewLinks.Open()
rsAdminNewLinks_numRows = 0
%>
<%
Dim rsAdminNewNews__MMColParam
rsAdminNewNews__MMColParam = "No"
if (Request("MM_EmptyValue") <> "") then rsAdminNewNews__MMColParam = Request("MM_EmptyValue")
%>
<%
set rsAdminNewNews = Server.CreateObject("ADODB.Recordset")
rsAdminNewNews.ActiveConnection = MM_connDUportal_STRING
rsAdminNewNews.Source = "SELECT * FROM NEWS WHERE NEWS_APPROVED = " + Replace(rsAdminNewNews__MMColParam, "'", "''") + " ORDER BY NEWS_ID ASC"
rsAdminNewNews.CursorType = 0
rsAdminNewNews.CursorLocation = 2
rsAdminNewNews.LockType = 3
rsAdminNewNews.Open()
rsAdminNewNews_numRows = 0
%>
<%
Dim RepeatAdminNewLinks__numRows
RepeatAdminNewLinks__numRows = -1
Dim RepeatAdminNewLinks__index
RepeatAdminNewLinks__index = 0
rsAdminNewLinks_numRows = rsAdminNewLinks_numRows + RepeatAdminNewLinks__numRows
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsAdminNewNews_numRows = rsAdminNewNews_numRows + Repeat1__numRows
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
                <td align="left" valign="middle" bgcolor="#00CC99" height="20">
                  <div class = "login">&nbsp;<b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="users.asp">USERS</a> 
                    | <a href="banners.asp">BANNERS</a> | <a href="links.asp">LINKS</a> 
                    | <a href="forums.asp">FORUMS</a> | <a href="news.asp">NEWS</a> 
                    | <a href="polls.asp">POLLS</a></font></b></div>
                </td>
              </tr>
              <tr> 
                <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="middle" height="20" bgcolor="#CCCCCC"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp; 
                  LINKS FOR APPROVAL</font></b></td>
              </tr>
              <tr> 
                <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="top"> 
                  <form name="LINKS" method="get" action="linkApproving.asp">
                    <table width="100%" border="0" cellspacing="0" cellpadding="5">
                      <tr> 
                        <td align="left" valign="top" colspan="2"> 
                          <% 
While ((RepeatAdminNewLinks__numRows <> 0) AND (NOT rsAdminNewLinks.EOF)) 
%>
                          <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr> 
                              <td align="left" valign="middle"> 
                                <div class = "links"> 
                                  <input type="checkbox" name="LINK_ID" value="<%=(rsAdminNewLinks.Fields.Item("LINK_ID").Value)%>">
                                  <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font size="1"><a href="../DUdirectory/dirHitting.asp?id=<%=(rsAdminNewLinks.Fields.Item("LINK_ID").Value)%>&url=<%=(rsAdminNewLinks.Fields.Item("LINK_URL").Value)%>" target="_blank" onClick="window.location.reload(true);"><%=(rsAdminNewLinks.Fields.Item("LINK_NAME").Value)%></a></font></b></font> <font size="1"><i>(<%=(rsAdminNewLinks.Fields.Item("LINK_URL").Value)%>)</i></font></div>
                              </td>
                              <td align="right" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Dated:</b> 
                                <%= (rsAdminNewLinks.Fields.Item("LINK_DATE").Value)%></font></td>
                            </tr>
                            <tr> 
                              <td align="left" valign="top" colspan="2"> 
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td width="14">&nbsp;</td>
                                    <td align="left" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                                      <%=(rsAdminNewLinks.Fields.Item("LINK_DESC").Value)%></font></td>
                                  </tr>
                                </table>
                              </td>
                            </tr>
                          </table>
                          <% 
  RepeatAdminNewLinks__index=RepeatAdminNewLinks__index+1
  RepeatAdminNewLinks__numRows=RepeatAdminNewLinks__numRows-1
  rsAdminNewLinks.MoveNext()
Wend
%>
                        </td>
                      </tr>
                      <tr> 
                        <td align="left" valign="top" colspan="2"> 
                          <input type="submit" name="Submit" value="Approve">
                          <input type="submit" name="Submit" value="Delete">
                        </td>
                      </tr>
                    </table>
                  </form>
                </td>
              </tr>
              <tr> 
                <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="middle" bgcolor="#CCCCCC" height="20"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp; 
                  NEWS FOR APPROVAL</font></b></td>
              </tr>
              <tr> 
                <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="top"> 
                  <form name="NEWS" method="get" action="newsApproving.asp">
                    <table width="100%" border="0" cellspacing="0" cellpadding="5">
                      <tr> 
                        <td align="left" valign="top" colspan="2"> 
                          <% 
While ((Repeat1__numRows <> 0) AND (NOT rsAdminNewNews.EOF)) 
%>
                          <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr> 
                              <td align="left" valign="middle"> 
                                <div class = "links"> 
                                  <input type="checkbox" name="NEWS_ID" value="<%=(rsAdminNewNews.Fields.Item("NEWS_ID").Value)%>">
                                  <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font size="1"><a href="<%=(rsAdminNewNews.Fields.Item("NEWS_URL").Value)%>"><%=(rsAdminNewNews.Fields.Item("NEWS_TITLE").Value)%></a></font></b></font> <font size="1"><i>(<%=(rsAdminNewNews.Fields.Item("NEWS_URL").Value)%>)</i></font></div>
                              </td>
                              <td align="right" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Dated:</b> 
                                <%=(rsAdminNewNews.Fields.Item("NEWS_DATE").Value)%></font></td>
                            </tr>
                            <tr> 
                              <td align="left" valign="top" colspan="2"> 
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td width="14">&nbsp;</td>
                                    <td align="left" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                                      <%=(rsAdminNewNews.Fields.Item("NEWS_DESC").Value)%></font></td>
                                  </tr>
                                </table>
                              </td>
                            </tr>
                          </table>
                          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsAdminNewNews.MoveNext()
Wend
%>
                        </td>
                      </tr>
                      <tr> 
                        <td align="left" valign="top" colspan="2"> 
                          <input type="submit" name="Submit" value="Approve">
                          <input type="submit" name="Submit" value="Delete">
                        </td>
                      </tr>
                    </table>
                  </form>
                </td>
              </tr>
              <tr> 
                <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
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
rsAdminNewLinks.Close()
%>
<%
rsAdminNewNews.Close()
%>
