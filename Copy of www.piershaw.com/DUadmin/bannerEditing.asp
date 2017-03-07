<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/connDUportal.asp" -->


<%
if(Request.QueryString("B_ID") <> "") then spBannerDeleting__varID = Request.QueryString("B_ID")
if(Request.QueryString("B_NAME") <> "") then cmdSaveBanner__varName = Request.QueryString("B_NAME")
if(Request.QueryString("B_URL") <> "") then cmdSaveBanner__varUrl = Request.QueryString("B_URL")
if(Request.QueryString("B_IMAGE") <> "") then cmdSaveBanner__varImage = Request.QueryString("B_IMAGE")
if(Request.QueryString("B_ALT") <> "") then cmdSaveBanner__varAlt = Request.QueryString("B_ALT")
%>





<% If Request.QueryString("Submit") = "Delete" Then %>
<%
set spBannerDeleting = Server.CreateObject("ADODB.Command")
spBannerDeleting.ActiveConnection = MM_connDUportal_STRING
spBannerDeleting.CommandText = "DELETE FROM BANNERS  WHERE B_ID = " + Replace(spBannerDeleting__varID, "'", "''") + ""
spBannerDeleting.CommandType = 1
spBannerDeleting.CommandTimeout = 0
spBannerDeleting.Prepared = true
spBannerDeleting.Execute()
%>
<% End If %>





<% If Request.QueryString("Submit") = "Save" Then %>
<%
set cmdSaveBanner = Server.CreateObject("ADODB.Command")
cmdSaveBanner.ActiveConnection = MM_connDUportal_STRING
cmdSaveBanner.CommandText = "UPDATE BANNERS  SET B_NAME = '" + Replace(cmdSaveBanner__varName, "'", "''") + "', B_URL = '" + Replace(cmdSaveBanner__varUrl, "'", "''") + "', B_IMAGE = '" + Replace(cmdSaveBanner__varImage, "'", "''") + "', B_ALT = '" + Replace(cmdSaveBanner__varAlt, "'", "''") + "' WHERE B_ID = " + Replace(spBannerDeleting__varID, "'", "''") + ""
cmdSaveBanner.CommandType = 1
cmdSaveBanner.CommandTimeout = 0
cmdSaveBanner.Prepared = true
cmdSaveBanner.Execute()
%>
<% End If %>



<%
Response.Redirect("../DUadmin/banners.asp")
%>















