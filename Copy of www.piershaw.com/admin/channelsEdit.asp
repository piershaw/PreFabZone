<!--#include file="../includes/inc_config.asp" -->
<html>
<head>
<title><%= strPageTitle %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../assets/DUportal.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div class="links">
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr align="left" valign="top"> 
      <td colspan="3">
        <!--#include file="../includes/inc_header.asp" -->
      </td>
  </tr>
  <tr align="left" valign="top"> 
    <td width="<%= strLeftSize %>">
      <!--#include file="inc_menu.asp" -->
    </td>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          
          <tr> 
            <td align="left" valign="top"> <!--#include file="inc_channel_edit.asp" --></td>
          </tr>
		   <tr> 
            <td align="left" valign="top"> <!--#include file="inc_channel_listing.asp" --></td>
          </tr>
          
        </table>
       
       
      </td>
    
  </tr>
</table>
</div>
</body>
</html>
