<!--#include file="../Connections/connDUportal.asp" -->
<%
'****************************************************************************************
'**  Copyright Notice                                                               
'**  Copyright 2003 DUware All Rights Reserved.                                
'**  This program is free software; you can modify (at your own risk) any part of it 
'**  under the terms of the License that accompanies this software and use it both 
'**  privately and commercially.
'**  All copyright notices must remain in tacked in the scripts and the 
'**  outputted HTML.
'**  You may use parts of this program in your own private work, but you may NOT
'**  redistribute, repackage, or sell the whole or any part of this program even 
'**  if it is modified or reverse engineered in whole or in part without express 
'**  permission from the author.
'**  You may not pass the whole or any part of this application off as your own work.
'**  All links to DUware and powered by logo's must remain unchanged and in place
'**  and must remain visible when the pages are viewed unless permission is first granted
'**  by the copyright holder.
'**  This program is distributed in the hope that it will be useful,
'**  but WITHOUT ANY WARRANTY; without even the implied warranty of
'**  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE OR ANY OTHER 
'**  WARRANTIES WHETHER EXPRESSED OR IMPLIED.
'**  No official support is available for this program but you may post support questions at: -
'**  http://www.duware.com/support
'****************************************************************************************
%>

<%
if(Request.QueryString("iData") <> "") then cmdHits__varData = Request.QueryString("iData")
%>
<%

set cmdHits = Server.CreateObject("ADODB.Command")
cmdHits.ActiveConnection = MM_connDUportal_STRING
cmdHits.CommandText = "UPDATE DATAS  SET DAT_HITS = DAT_HITS + 1  WHERE DAT_ID = " + Replace(cmdHits__varData, "'", "''") + ""
cmdHits.CommandType = 1
cmdHits.CommandTimeout = 0
cmdHits.Prepared = true
cmdHits.Execute()

%>

<%
Dim rsDetail__MMColParam
rsDetail__MMColParam = "0"
if (Request.QueryString("iData") <> "") then rsDetail__MMColParam = Request.QueryString("iData")
%>
<%
set rsDetail = Server.CreateObject("ADODB.Recordset")
rsDetail.ActiveConnection = MM_connDUportal_STRING
rsDetail.Source = "SELECT *  FROM DATAS,  CATEGORIES, CHANNELS  WHERE DAT_ID = " + Replace(rsDetail__MMColParam, "'", "''") + " AND DAT_CATEGORY = CAT_ID AND CAT_CHANNEL = CHA_ID"
rsDetail.CursorType = 0
rsDetail.CursorLocation = 2
rsDetail.LockType = 3
rsDetail.Open()
rsDetail_numRows = 0
%>
              <%
Dim dat_rated
Dim dat_rate_count 
Dim dat_rate_value
dat_rate_count = rsDetail.Fields.Item("DAT_RATES").Value
dat_rate_value = rsDetail.Fields.Item("DAT_RATED").Value
If dat_rate_count > 0 Then 
dat_rated = (dat_rate_value/dat_rate_count)
else
dat_rated = 0
end if
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
                      <td align="left" valign="middle" class="textBoldColor"><a href="default.asp">HOME 
                        </a> &raquo; <a href="channel.asp?iChannel=<%=(rsDetail.Fields.Item("CAT_CHANNEL").Value)%>&nChannel=<%=(rsDetail.Fields.Item("CHA_NAME").Value)%>"><%=UCASE(rsDetail.Fields.Item("CHA_NAME").Value)%></a> 
                        &raquo; <a href="type.asp?iCat=<%=(rsDetail.Fields.Item("DAT_CATEGORY").Value)%>&iChannel=<%=(rsDetail.Fields.Item("CAT_CHANNEL").Value)%>&nChannel=<%=(rsDetail.Fields.Item("CHA_NAME").Value)%>"><%=UCASE(rsDetail.Fields.Item("CAT_NAME").Value)%></a> 
                        &raquo; DETAIL</td>
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
                <td align="left" valign="top" class="bgTable"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                   
                    <tr> 
                      <td colspan="2" align="left" valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td align="left" valign="top" bgcolor="#000000"><img src="../assets/_spacer.gif" width="1" height="1"></td>
                          </tr>
                          <tr> 
                            <td align="left" valign="top"> <table width="100%" border="0" cellspacing="2" cellpadding="2">
                                <tr valign="middle"> 
                                  <td colspan="2" align="left" class="text"><b>Name:</b> 
                                    <%=(rsDetail.Fields.Item("DAT_NAME").Value)%></td>
                                </tr>
                                <tr valign="middle"> 
                                  <td align="left" class="text"><strong>Dated:</strong> 
                                    <%=(rsDetail.Fields.Item("DAT_DATED").Value)%> </td>
                                  <% If Request.Cookies(Request.QueryString("iData")) = "" Then %>
                                  <form action="../includes/inc_rating.asp" method="get" name="rate" id="rate">
                                    <td align="left" class="text"> <input name="iChannel" type="hidden" id="iChannel" value="<%=(rsDetail.Fields.Item("CAT_CHANNEL").Value)%>"> 
                                      <input name="iCat" type="hidden" id="iCat" value="<%=(rsDetail.Fields.Item("DAT_CATEGORY").Value)%>"> 
                                      <input name="iData" type="hidden" id="iData" value="<%=(rsDetail.Fields.Item("DAT_ID").Value)%>"> 
				      <input name="nChannel" type="hidden" id="nChannel" value="<%=(rsDetail.Fields.Item("CHA_NAME").Value)%>"> 
                                      <select name="iRate" class="form">
                                        <option value="5" selected>* * * * *</option>
                                        <option value="4">* * * *</option>
                                        <option value="3">* * *</option>
                                        <option value="2">* *</option>
                                        <option value="1">*</option>
                                      </select> <input name="rate" type="submit" class="button" id="Rate" value="Rate"> 
                                    </td>
                                  </form>
                                  <% End If %>
                                </tr>
                                <tr valign="middle"> 
                                  <td width="50%" align="left" class="text"><strong>Submited 
                                    By:</strong> <%=(rsDetail.Fields.Item("DAT_USER").Value)%></td>
                                  <td width="50%" align="left" class="text"><strong>Views:</strong> 
                                    <%=(rsDetail.Fields.Item("DAT_HITS").Value)%></td>
                                </tr>
                                <tr valign="middle"> 
                                  <td width="50%" align="left" class="text"><strong>Rating:</strong> 
                                    <img src="../assets/<%= FormatNumber(dat_rated, 1, -2, -2, -2) %>.gif" align="absmiddle"> 
                                    (<%= FormatNumber(dat_rated, 1, -2, -2, -2) %>)</td>
                                  <td width="50%" align="left" class="text"><strong>By: 
                                    </strong><%=(rsDetail.Fields.Item("DAT_RATES").Value)%> users</td>
                                </tr>
                                <% If Request.QueryString("nChannel") = "Products" Then %>
                                <tr valign="middle"> 
                                  <td align="left" class="text"><strong>Product 
                                    No:</strong> <%=(rsDetail.Fields.Item("DAT_SKU").Value)%></td>
                                  <td align="left" class="text"><strong>Brand:</strong> 
                                    <%=(rsDetail.Fields.Item("DAT_BRAND").Value)%></td>
                                </tr>
                                <% End If %>
                                <% If Request.QueryString("nChannel") = "Products" OR Request.QueryString("nChannel") = "Ads" Then %>
                                <tr valign="middle"> 
                                  <td align="left" class="text"><strong>Price:</strong> 
                                    <%= myPaypalCurrencySign %><%=(rsDetail.Fields.Item("DAT_PRICE").Value)%></td>
                                  <td align="left" class="text"><strong>Shipping 
                                    Cost:</strong> <%= myPaypalCurrencySign %><%=(rsDetail.Fields.Item("DAT_SHIP").Value)%></td>
                                </tr>
                                <% End If %>
                                <% If Request.QueryString("nChannel") = "Ads" Then %>
                                <tr valign="middle"> 
                                  <td align="left" class="text"><strong>Quantity:</strong> 
                                    <%=(rsDetail.Fields.Item("DAT_QUANTITY").Value)%></td>
                                  <td align="left" class="text"><strong>Expired:</strong> 
                                    <%=(rsDetail.Fields.Item("DAT_EXPIRED").Value)%></td>
                                </tr>
                                <% End If %>

				<% If Request.QueryString("nChannel") = "Events" Or Request.QueryString("nChannel") = "Businesses" Or  Request.QueryString("nChannel") = "Ads" Then %>
                                <tr valign="middle"> 
                                  
                                  <td  colspan="2" align="left" class="text"><strong>Location:</strong> 
                                    <%=(rsDetail.Fields.Item("DAT_LOCATION").Value)%>
</td>
                                </tr>
                                <% End If %>



                                <tr valign="middle"> 
                                  <td colspan="2" align="left" class="text"><b>Description:</b> 
                                    <% = TrimBody(rsDetail.Fields.Item("DAT_DESCRIPTION").Value) %> </td>
                                </tr>
                                <% If Request.QueryString("nChannel") = "Products" Then %>
                                <tr valign="top"> 
                                  <td align="left" class="text"> <table border="0" cellspacing="0" cellpadding="0">
                                      <tr align="left" valign="middle"> 
                                        <form target="paypal" action="https://www.paypal.com/cgi-bin/webscr" method="post">
                                          <td> <input type="hidden" name="cmd" value="_cart"> 
                                            <input type="hidden" name="business" value="<%= myPaypalID %>"> 
                                            <input type="hidden" name="currency_code" value="<%= myPaypalCurrency %>"> 
                                            <input type="hidden" name="return" value="<%= myReturnURL %>"> 
                                            <input type="hidden" name="cancel_return" value="<%= myCancelURL %>"> 
                                            <input type="hidden" name="item_name" value="<%=(rsDetail.Fields.Item("DAT_SKU").Value)%> - <%=(rsDetail.Fields.Item("DAT_NAME").Value)%>"> 
                                            <input type="hidden" name="item_number" value="<%=(rsDetail.Fields.Item("DAT_ID").Value)%>"> 
                                            <input type="hidden" name="amount" value="<%=(rsDetail.Fields.Item("DAT_PRICE").Value)%>"> 
                                            <input type="hidden" name="shipping" value="<%=(rsDetail.Fields.Item("DAT_SHIP").Value)%>"> 
                                            <input type="hidden" name="add" value="1"> 
                                            <input name="addtocart" type="submit" class="form" value="Add to Cart"> 
                                          </td>
                                        </form>
                                      </tr>
                                    </table></td>
                                  <td align="left" class="text"> <table border="0" cellspacing="0" cellpadding="0">
                                      <tr align="left" valign="middle"> 
                                        <form target="paypal" action="https://www.paypal.com/cgi-bin/webscr" method="post">
                                          <td> <input type="hidden" name="cmd" value="_cart"> 
                                            <input type="hidden" name="business" value="<%= myPaypalID %>"> 
                                            <input type="hidden" name="display" value="1"> 
                                            <input name="viewcart" type="submit" class="form" value="View Cart"> 
                                          </td>
                                        </form>
                                      </tr>
                                    </table></td>
                                </tr>
                                
                                <% End If %>
								 <% If LEN(rsDetail.Fields.Item("DAT_URL").Value) > "10" Then %>
								<tr align="right" valign="middle"> 
                                  <td colspan="2" class="text"><a href="<%=(rsDetail.Fields.Item("DAT_URL").Value)%>" target="_blank">more ...</a> </td>
                                </tr>
								 <% End If %>
                                <% If rsDetail.Fields.Item("DAT_PICTURE").Value <> "" Then %>
                                <tr align="center" valign="middle"> 
                                  <td colspan="2" class="text"><img src="../pictures/<%=(rsDetail.Fields.Item("DAT_PICTURE").Value)%>" border="0"></td>
                                </tr>
                                <% End If %>
                              </table></td>
                          </tr>
                         
                          
                        </table></td>
                    </tr>
                  </table></td>
                <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif"></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td align="center" valign="top" background="../assets/bg_header_bottom.gif"></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="7" align="left" valign="top"><img src="../assets/_spacer.gif" width="1" height="1"></td>
  </tr>
</table>
<%
rsDetail.Close()
%>
