<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_AdminUser")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "default.asp"
  ' redirect with URL parameters (remove the "MM_Logoutnow" query param).
  if (MM_logoutRedirectPage = "") Then MM_logoutRedirectPage = CStr(Request.ServerVariables("URL"))
  If (InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
    MM_newQS = "?"
    For Each Item In Request.QueryString
      If (Item <> "MM_Logoutnow") Then
        If (Len(MM_newQS) > 1) Then MM_newQS = MM_newQS & "&"
        MM_newQS = MM_newQS & Item & "=" & Server.URLencode(Request.QueryString(Item))
      End If
    Next
    if (Len(MM_newQS) > 1) Then MM_logoutRedirectPage = MM_logoutRedirectPage & MM_newQS
  End If
  Response.Redirect(MM_logoutRedirectPage)
End If
%>
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="default.asp"
MM_grantAccess=false
If Session("MM_AdminUser") = "admin" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
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
 <link href="../assets/DUportal.css" rel="stylesheet" type="text/css"> 
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#003399">
              <tr> 
                <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../assets/bg_header.gif">
                    <tr> 
                      <td width="10"><img src="../assets/header_end_left.gif"></td>
                      <td align="left" valign="middle" class="textBoldColor">ADMIN 
                        MENU</td>
                      <td width="28" align="right" valign="middle"><img src="../assets/header_end_right.gif"></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif"></td>
                <td align="left" valign="top" class="bgTable"> <table width="100%" border="0" cellpadding="3" cellspacing="0">
                    <tr align="left" valign="middle" style=cursor:hand;  onmouseover="this.className='bgMouseOver';" onmouseout="this.className='bgMouseOff';"> 
                      <td width="5" ><img src="../assets/icon_cross.gif" hspace="0" vspace="2"></td>
                      <td class="textBoldColor"><a href="approve.asp">APPROVE</a></td>
                    </tr>
                    <tr align="left" valign="middle" style=cursor:hand;  onmouseover="this.className='bgMouseOver';" onmouseout="this.className='bgMouseOff';"> 
                      <td ><img src="../assets/icon_cross.gif" hspace="0" vspace="2"></td>
                      <td class="textBoldColor"><a href="submit.asp">ADD NEW</a></td>
                    </tr>
                    <tr align="left" valign="middle" style=cursor:hand;  onmouseover="this.className='bgMouseOver';" onmouseout="this.className='bgMouseOff';"> 
                      <td ><img src="../assets/icon_cross.gif" hspace="0" vspace="2"></td>
                      <td class="textBoldColor"><a href="users.asp">USERS</a></td>
                    </tr>
                    <tr align="left" valign="middle" style=cursor:hand;  onmouseover="this.className='bgMouseOver';" onmouseout="this.className='bgMouseOff';"> 
                      <td ><img src="../assets/icon_cross.gif" hspace="0" vspace="2"></td>
                      <td class="textBoldColor"><a href="polls.asp">POLLS</a></td>
                    </tr>
                    <tr align="left" valign="middle" style=cursor:hand;  onmouseover="this.className='bgMouseOver';" onmouseout="this.className='bgMouseOff';"> 
                      <td ><img src="../assets/icon_cross.gif" hspace="0" vspace="2"></td>
                      <td class="textBoldColor"><a href="datas.asp">DATA</a></td>
                    </tr>
                    <tr align="left" valign="middle" style=cursor:hand;  onmouseover="this.className='bgMouseOver';" onmouseout="this.className='bgMouseOff';"> 
                      <td ><img src="../assets/icon_cross.gif" hspace="0" vspace="2"></td>
                      <td class="textBoldColor"><a href="channels.asp">CHANNELS</a></td>
                    </tr>
                    <tr align="left" valign="middle" style=cursor:hand;  onmouseover="this.className='bgMouseOver';" onmouseout="this.className='bgMouseOff';"> 
                      <td ><img src="../assets/icon_cross.gif" hspace="0" vspace="2"></td>
                      <td class="textBoldColor"><a href="types.asp">CATEGORIES</a></td>
                    </tr>
                    <tr align="left" valign="middle" style=cursor:hand;  onmouseover="this.className='bgMouseOver';" onmouseout="this.className='bgMouseOff';"> 
                      <td ><img src="../assets/icon_cross.gif" hspace="0" vspace="2"></td>
                      <td class="textBoldColor"><a href="config.asp">CONFIGURATION</a></td>
                    </tr>
                    <tr align="left" valign="middle" style=cursor:hand;  onmouseover="this.className='bgMouseOver';" onmouseout="this.className='bgMouseOff';"> 
                      <td ><img src="../assets/icon_cross.gif" hspace="0" vspace="2"></td>
                      <td class="textBoldColor"><a href="../home/" target="_blank">VIEW 
                        PORTAL</a></td>
                    </tr>
                    <tr align="left" valign="middle" style=cursor:hand;  onmouseover="this.className='bgMouseOver';" onmouseout="this.className='bgMouseOff';"> 
                      <td ><img src="../assets/icon_cross.gif" hspace="0" vspace="2"></td>
                      <td class="textBoldColor"><a href="<%= MM_Logout %>">LOG 
                        OUT</a></td>
                    </tr>
                  </table></td>
                <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif"></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td align="center" valign="top" background="../assets/bg_header_bottom.gif"><table border="0" cellpadding="0" cellspacing="0" class="bgTable" >
              <tr> 
                <td><img src="../assets/header_bottom.gif"></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="5" align="left" valign="top"><img src="../assets/_spacer.gif" width="1" height="1"></td>
  </tr>
</table>
