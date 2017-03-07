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
' *** Edit Operations: declare variables

MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""

' Add your variables here!
MM_AutoNum = -1
MM_errorURL = "error.asp"

%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insertAuto")) <> "") Then


  MM_editConnection = MM_connDUportal_STRING
  MM_editTable = "QUESTIONS"
  MM_editColumn = "QUEST_ID"
  'MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "pollsAns.asp"
  MM_fieldsStr  = "QUESTION|value|DESCRIPTION|value"
  MM_columnsStr = "QUESTION|',none,''|QUEST_DESCRIPTION|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(i+1) = CStr(Request.Form(MM_fields(i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Insert Record: construct a sql insert statement and execute it

If (CStr(Request("MM_insertAuto")) <> "") Then


  If (Not MM_abortEdit) Then
    'Execute the insert
    'Retrieve record (assume a 1-1 relationship)
    Set MM_rs = Server.CreateObject("ADODB.Recordset")
    MM_rs.ActiveConnection = MM_editConnection
    MM_rs.Source = "SELECT * FROM " & MM_editTable
    MM_rs.CursorType = 1
    MM_rs.CursorLocation = 2
    MM_rs.LockType = 3
    MM_rs.Open()
    MM_rs_numRows = 0


    MM_rs.AddNew
    
    ' Fill in the fields on the form
    Dim varTemp	
	For i = LBound(MM_fields) To UBound(MM_fields) Step 2
        varTemp = MM_fields(i+1)		
        If (varTemp <> "NULL") And (varTemp <> "") Then    
            MM_rs.Fields(MM_columns(i)).value = varTemp
        End if	
    Next

    MM_rs.Update
    MM_rs.MoveLast

    MM_AutoNum = MM_rs.Fields(MM_editColumn).value
    MM_rs.Close

    ' Modify the following to save/check the Auto-Number field
    '<----------------------------MODIFY------------------------------>

    If (MM_AutoNum <> -1) then
        Session("ID") = MM_AutoNum
    else
        Response.Redirect(MM_errorURL)
    end if

    '<--------------------------END MODIFY---------------------------->
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
set rsAdminPolls = Server.CreateObject("ADODB.Recordset")
rsAdminPolls.ActiveConnection = MM_connDUportal_STRING
rsAdminPolls.Source = "SELECT * FROM QUESTIONS ORDER BY QUEST_ID ASC"
rsAdminPolls.CursorType = 0
rsAdminPolls.CursorLocation = 2
rsAdminPolls.LockType = 3
rsAdminPolls.Open()
rsAdminPolls_numRows = 0
%>
<html>
<head>
<title>DUportal</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/default.css" type="text/css">
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') {
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (val<min || max<val) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="top" height="60" bgcolor="#009999"><img src="../assets/DUportalAdmin.gif" width="268" height="60"></td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
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
              <tr> 
                <td align="left" valign="middle" height="20" bgcolor="#CCCCCC"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp;MANAGING 
                  POLLS</font></b></td>
                <td align="right" valign="middle" height="20" bgcolor="#CCCCCC"> 
                  <font face="Verdana, Arial, Helvetica, sans-serif"> <font size="1"> 
                  &nbsp; </font> </font> </td>
              </tr>
              <tr> 
                <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="top" colspan="2"> 
                  <table width="100%" border="0" cellspacing="5" cellpadding="5">
                    <tr> 
                      <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#FF0000">Add 
                        new poll:</font></b></font></td>
                    </tr>
                    <tr> 
                      <td> 
                        <form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
                          <table width="100%" border="0" cellspacing="5" cellpadding="5">
                            <tr> 
                              <td align="right">Poll Question:</td>
                              <td> 
                                <input type="text" name="QUESTION" size="50">
                              </td>
                            </tr>
                            <tr> 
                              <td align="right">Poll Description</td>
                              <td> 
                                <textarea name="DESCRIPTION" cols="40" rows="3"></textarea>
                              </td>
                            </tr>
                            <tr> 
                              <td align="right">&nbsp;</td>
                              <td> 
                                <input type="submit" name="Submit" value="AddQuestion">
                              </td>
                            </tr>
                          </table>
                          <input type="hidden" name="MM_insertQuestion" value="true">
                          <input type="hidden" name="MM_insertAuto" value="true">
                        </form>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr bgcolor="#000000"> 
                <td align="left" valign="top" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="top" colspan="2"> 
                  <table width="100%" border="0" cellspacing="5" cellpadding="5">
                    <tr> 
                      <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#FF0000">Activate/Deactivate 
                        Polls:</font></b></font></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="top"> 
                        <form name="form2" method="get" action="pollActivating.asp">
                          <font size="2" face="Verdana, Arial, Helvetica, sans-serif">Select 
                          a poll to activate; current active poll will be deactivated</font><br>
                          <select name="ID">
                            <%
While (NOT rsAdminPolls.EOF)
%>
                            <option value="<%=(rsAdminPolls.Fields.Item("QUEST_ID").Value)%>" <%if (CStr(rsAdminPolls.Fields.Item("QUEST_ID").Value) = CStr(rsAdminPolls.Fields.Item("QUESTION").Value)) then Response.Write("SELECTED") : Response.Write("")%> ><%=(rsAdminPolls.Fields.Item("QUESTION").Value)%></option>
                            <%
  rsAdminPolls.MoveNext()
Wend
If (rsAdminPolls.CursorType > 0) Then
  rsAdminPolls.MoveFirst
Else
  rsAdminPolls.Requery
End If
%>
                          </select>
                          <input type="submit" name="Submit2" value="Go">
                        </form>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr bgcolor="#000000"> 
                <td align="left" valign="top" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="top" colspan="2"> 
                  <table width="100%" border="0" cellspacing="5" cellpadding="5">
                    <tr> 
                      <td>&nbsp;</td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr bgcolor="#000000"> 
                <td align="left" valign="top" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="top" colspan="2">&nbsp;</td>
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
rsAdminPolls.Close()
%>

