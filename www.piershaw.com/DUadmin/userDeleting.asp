
<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/connDUportal.asp" -->
<%

if(Request.QueryString("U_ID") <> "") then spUserDeleting__varID = Request.QueryString("U_ID")

%>
<%

set spUserDeleting = Server.CreateObject("ADODB.Command")
spUserDeleting.ActiveConnection = MM_connDUportal_STRING
spUserDeleting.CommandText = "DELETE FROM USERS  WHERE U_ID IN ('" + Replace(spUserDeleting__varID, "'", "''") + "') "
spUserDeleting.CommandType = 1
spUserDeleting.CommandTimeout = 0
spUserDeleting.Prepared = true
spUserDeleting.Execute()
Response.Redirect("../DUadmin/users.asp")
%>


