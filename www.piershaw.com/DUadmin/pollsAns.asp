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
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then

  MM_editConnection = MM_connDUportal_STRING
  MM_editTable = "ANSWERS"
  MM_editRedirectUrl = "pollsAns.asp"
  MM_fieldsStr  = "ANSWER|value|QUEST_ID|value"
  MM_columnsStr = "ANSWERS|',none,''|QUEST_ID|none,none,NULL"

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

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    FormVal = MM_fields(i+1)
    MM_typeArray = Split(MM_columns(i+1),",")
    Delim = MM_typeArray(0)
    If (Delim = "none") Then Delim = ""
    AltVal = MM_typeArray(1)
    If (AltVal = "none") Then AltVal = ""
    EmptyVal = MM_typeArray(2)
    If (EmptyVal = "none") Then EmptyVal = ""
    If (FormVal = "") Then
      FormVal = EmptyVal
    Else
      If (AltVal <> "") Then
        FormVal = AltVal
      ElseIf (Delim = "'") Then  ' escape quotes
        FormVal = "'" & Replace(FormVal,"'","''") & "'"
      Else
        FormVal = Delim + FormVal + Delim
      End If
    End If
    If (i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End if
    MM_tableValues = MM_tableValues & MM_columns(i)
    MM_dbValues = MM_dbValues & FormVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim rsID__MMColParam
rsID__MMColParam = "1"
if (Session("ID") <> "") then rsID__MMColParam = Session("ID")
%>
<%
set rsID = Server.CreateObject("ADODB.Recordset")
rsID.ActiveConnection = MM_connDUportal_STRING
rsID.Source = "SELECT * FROM QUESTIONS WHERE QUEST_ID = " + Replace(rsID__MMColParam, "'", "''") + ""
rsID.CursorType = 0
rsID.CursorLocation = 2
rsID.LockType = 3
rsID.Open()
rsID_numRows = 0
%>
<%
Dim rsAns__MMColParam
rsAns__MMColParam = "1"
if (Session("ID") <> "") then rsAns__MMColParam = Session("ID")
%>
<%
set rsAns = Server.CreateObject("ADODB.Recordset")
rsAns.ActiveConnection = MM_connDUportal_STRING
rsAns.Source = "SELECT * FROM ANSWERS WHERE QUEST_ID = " + Replace(rsAns__MMColParam, "'", "''") + " ORDER BY ANS_ID DESC"
rsAns.CursorType = 0
rsAns.CursorLocation = 2
rsAns.LockType = 3
rsAns.Open()
rsAns_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsAns_numRows = rsAns_numRows + Repeat1__numRows
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

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
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
                        answers:</font></b></font></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="top"> 
                        <form name="form1" method="POST" action="<%=MM_editAction%>">
                          <table width="100%" border="0" cellspacing="5" cellpadding="5">
                            <tr> 
                              <td align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Question:</font></b></td>
                              <td><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=(rsID.Fields.Item("QUESTION").Value)%></font></b></td>
                            </tr>
                            <tr> 
                              <td align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Description:</font></b></td>
                              <td><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=(rsID.Fields.Item("QUEST_DESCRIPTION").Value)%></font></b></td>
                            </tr>
                            <% If Not rsAns.EOF Or Not rsAns.BOF Then %>
                            <tr> 
                              <td align="right" valign="top"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Answer</font></b></td>
                              <td> 
                                <% 
While ((Repeat1__numRows <> 0) AND (NOT rsAns.EOF)) 
%>
                                <table width="100%" border="0" cellspacing="2" cellpadding="2">
                                  <tr> 
                                    <td align="left" valign="top"><font size="2"><i><b><font color="#009999" face="Verdana, Arial, Helvetica, sans-serif"><%=(rsAns.Fields.Item("ANSWERS").Value)%></font></b></i></font></td>
                                  </tr>
                                </table>
                                <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsAns.MoveNext()
Wend
%>
                              </td>
                            </tr>
                            <% End If ' end Not rsAns.EOF Or NOT rsAns.BOF %>
                            <tr> 
                              <td align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Add 
                                an answer: </font></b></td>
                              <td> 
                                <input type="text" name="ANSWER" size="40" maxlength="40">
                                <input type="hidden" name="QUEST_ID" value="<%=(rsID.Fields.Item("QUEST_ID").Value)%>">
                                <input type="submit" name="Submit" value="Add">
                                <input type="button" name="Submit2" value="Finish" onClick="MM_goToURL('parent','polls.asp');return document.MM_returnValue">
                              </td>
                            </tr>
                          </table>
                          <input type="hidden" name="MM_insert" value="true">
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
rsID.Close()
%>
<%
rsAns.Close()
%>
