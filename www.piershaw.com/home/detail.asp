<!--#include file="../includes/inc_config.asp" -->
<html>
<head>
<title><%= strPageTitle %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../assets/DUportal.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div class="links">
<table width="<%= strPageSize %>" border="0" cellspacing="2" cellpadding="2">
  <tr align="left" valign="top"> 
      <td colspan="3">
        <!--#include file="../includes/inc_header.asp" -->
      </td>
  </tr>
  <tr align="left" valign="top"> 
    <td width="<%= strLeftSize %>">
      <!--#include file="../includes/inc_left.asp" -->
    </td>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td align="left" valign="top">
			
			 <% If Request.QueryString("nChannel") = "Topics" Then %>
			<!--#include file="../includes/inc_detail_forum.asp" --> 
			 <% Else %>
			 <!--#include file="../includes/inc_detail.asp" -->
			  <!--#include file="../includes/inc_detail_related.asp" -->					
			 <% End If %>
			 
			 <% If Request.QueryString("nChannel") = "Ads" Then %>
			<!--#include file="../includes/inc_detail_contact.asp" --> 					
			 <% End If %>
			 
			 
			
			</td>
          </tr>
         
          <tr>
            <td align="left" valign="top">
			
			 <!--#include file="../includes/inc_footer.asp" -->
			 
			 </td>
          </tr>
        </table>
		
					
       
       
      </td>
    <% If Request.QueryString("nChannel") <> "Topics" Then %>
    <td width="<%= strRightSize %>">
      <!--#include file="../includes/inc_right.asp" -->
    </td>
	<% End If %>
  </tr>
</table>
</div>
</body>
</html>
