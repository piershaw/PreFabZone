<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../Connections/connDUportal.asp" -->
<%
Dim rsMyMsg__MMColParam
rsMyMsg__MMColParam = "1"
if (Session("MM_Username") <> "") then rsMyMsg__MMColParam = Session("MM_Username")
%>
<%
set rsMyMsg = Server.CreateObject("ADODB.Recordset")
rsMyMsg.ActiveConnection = MM_connDUportal_STRING
rsMyMsg.Source = "SELECT * FROM MESSAGES WHERE MSG_AUTHOR = '" + Replace(rsMyMsg__MMColParam, "'", "''") + "' ORDER BY MSG_LAST_POST DESC"
rsMyMsg.CursorType = 0
rsMyMsg.CursorLocation = 2
rsMyMsg.LockType = 3
rsMyMsg.Open()
rsMyMsg_numRows = 0
%>
<%
Dim rsMyLinks__MMColParam
rsMyLinks__MMColParam = "1"
if (Session("MM_Username") <> "") then rsMyLinks__MMColParam = Session("MM_Username")
%>
<%
set rsMyLinks = Server.CreateObject("ADODB.Recordset")
rsMyLinks.ActiveConnection = MM_connDUportal_STRING
rsMyLinks.Source = "SELECT * FROM LINKS WHERE LINK_ADDER = '" + Replace(rsMyLinks__MMColParam, "'", "''") + "' ORDER BY LINK_ID DESC"
rsMyLinks.CursorType = 0
rsMyLinks.CursorLocation = 2
rsMyLinks.LockType = 3
rsMyLinks.Open()
rsMyLinks_numRows = 0
%>
<%
Dim rsMyNews__MMColParam
rsMyNews__MMColParam = "1"
if (Session("MM_Username") <> "") then rsMyNews__MMColParam = Session("MM_Username")
%>
<%
set rsMyNews = Server.CreateObject("ADODB.Recordset")
rsMyNews.ActiveConnection = MM_connDUportal_STRING
rsMyNews.Source = "SELECT * FROM NEWS WHERE NEWS_ADDER = '" + Replace(rsMyNews__MMColParam, "'", "''") + "' ORDER BY NEWS_DATE DESC"
rsMyNews.CursorType = 0
rsMyNews.CursorLocation = 2
rsMyNews.LockType = 3
rsMyNews.Open()
rsMyNews_numRows = 0
%>
<%
Dim RepeatMyMsg__numRows
RepeatMyMsg__numRows = -1
Dim RepeatMyMsg__index
RepeatMyMsg__index = 0
rsMyMsg_numRows = rsMyMsg_numRows + RepeatMyMsg__numRows
%>
<%
Dim RepeatMyNews__numRows
RepeatMyNews__numRows = -1
Dim RepeatMyNews__index
RepeatMyNews__index = 0
rsMyNews_numRows = rsMyNews_numRows + RepeatMyNews__numRows
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
Dim RepeatMyLinks__numRows
RepeatMyLinks__numRows = -1
Dim RepeatMyLinks__index
RepeatMyLinks__index = 0
rsMyLinks_numRows = rsMyLinks_numRows + RepeatMyLinks__numRows
%>
<% Response.Buffer = "true" %>
<html>
<head>
<title>DUportal</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/default.css" type="text/css">
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
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
  <tr valign="top"> 
    <td align="left" class = "bg_banner" height="62" valign="middle"> 
      <!--#include file="../includes/inc_header.asp" -->
    </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr valign="middle"> 
    <td align="left" class = "bg_navigator" height="20"> 
      <!--#include file="../includes/inc_navigator.asp" -->
    </td>
  </tr>
  <tr> 
    <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
  </tr>
  <tr> 
    <td align="left" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr align="left" valign="top"> 
          <td width="200"> 
            <!--#include file="inc_left.asp" -->
          </td>
          <td bgcolor="#000000" width="1"><img src="../assets/verticalBar.gif" width="1" height="5"></td>
          <td> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td align="left" valign="top" class = "bg_login" height="30"> 
                  <!--#include file="../includes/inc_login.asp" -->
                </td>
              </tr>
              <tr> 
                <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="top"> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td align="left" valign="middle" height="20" class = "bg_navigator"><font face="Verdana, Arial, Helvetica, sans-serif" size="3"><b><font size="1">&nbsp;MY 
                        LINKS </font></b></font></td>
						
                      <td align="right" valign="middle" height="20" class = "bg_navigator"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="../DUdirectory/dirAdd.asp?action=cat">ADD 
                        LINK</a>&nbsp;</font></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="top" colspan="2"> 
                        <% 
While ((RepeatMyLinks__numRows <> 0) AND (NOT rsMyLinks.EOF)) 
%>
                        <table width="100%" border="0" cellspacing="0" cellpadding="3">
                          <tr> 
                            <td align="left" valign="middle"> 
                              <div class = "links"><img src="../assets/bullet.gif" width="11" height="11" align="absmiddle"> 
                                <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font size="1"><a href="../DUdirectory/dirHitting.asp?id=<%=(rsMyLinks.Fields.Item("LINK_ID").Value)%>&url=<%=(rsMyLinks.Fields.Item("LINK_URL").Value)%>" target="_blank" onClick="window.location.reload(true);"><%=(rsMyLinks.Fields.Item("LINK_NAME").Value)%></a></font></b></font> <font size="1"><i>(<%=(rsMyLinks.Fields.Item("LINK_URL").Value)%>)</i></font></div>
                            </td>
                            <td align="right" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Dated:</b> 
                              <%= (rsMyLinks.Fields.Item("LINK_DATE").Value)%></font></td>
                          </tr>
                        </table>
                        <% 
  RepeatMyLinks__index=RepeatMyLinks__index+1
  RepeatMyLinks__numRows=RepeatMyLinks__numRows-1
  rsMyLinks.MoveNext()
Wend
%>
                      </td>
                    </tr>
                    <tr> 
                      <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="middle" height="20" class = "bg_navigator"><font face="Verdana, Arial, Helvetica, sans-serif" size="3"><b><font size="1">&nbsp;MY 
                        MESSAGES</font></b></font></td>
                      <td align="right" valign="middle" height="20" class = "bg_navigator"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="../DUforum/msgPost.asp">POST 
                        MESSAGE</a>&nbsp;</font></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                    </tr>
                    <tr valign="middle"> 
                      <td align="left" colspan="2"> 
                        <% 
While ((RepeatMyMsg__numRows <> 0) AND (NOT rsMyMsg.EOF)) 
%>
                        <table width="100%" border="0" cellspacing="3" cellpadding="3">
                          <tr> 
                            <td align="left" valign="top"> 
                              <div class = "links"><img src="../assets/bullet.gif" width="11" height="11" align="absmiddle"> 
                                <font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="../DUforum/msgDetail.asp?msg_id=<%=(rsMyMsg.Fields.Item("MSG_ID").Value)%>&for_id=<%=(rsMyMsg.Fields.Item("FOR_ID").Value)%>"><b><%=(rsMyMsg.Fields.Item("MSG_SUBJECT").Value)%></b></a> (Last Posted on <%=(rsMyMsg.Fields.Item("MSG_LAST_POST").Value)%>- Reads: <%=(rsMyMsg.Fields.Item("MSG_READ_COUNT").Value)%> - Replies: <%=(rsMyMsg.Fields.Item("MSG_REPLY_COUNT").Value)%>)</font></div>
                            </td>
                          </tr>
                        </table>
                        <% 
  RepeatMyMsg__index=RepeatMyMsg__index+1
  RepeatMyMsg__numRows=RepeatMyMsg__numRows-1
  rsMyMsg.MoveNext()
Wend
%>
                      </td>
                    </tr>
                    <tr> 
                      <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="middle" height="20" class = "bg_navigator"><font face="Verdana, Arial, Helvetica, sans-serif" size="3"><b><font size="1">&nbsp;MY 
                        NEWS</font></b></font></td>
                      <td align="right" valign="middle" height="20" class = "bg_navigator"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="../DUnews/newsAdd.asp">ADD 
                        NEWS</a>&nbsp; </font></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="top" colspan="2"> 
                        <% 
While ((RepeatMyNews__numRows <> 0) AND (NOT rsMyNews.EOF)) 
%>
                        <table width="100%" border="0" cellspacing="0" cellpadding="3">
                          <tr> 
                            <td align="left" valign="middle"> 
                              <div class = "links"><img src="../assets/bullet.gif" width="11" height="11" align="absmiddle"> 
                                <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font size="1"><a href="../DUnews/newsDetail.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "id=" & rsMyNews.Fields.Item("NEWS_ID").Value %>"><%=(rsMyNews.Fields.Item("NEWS_TITLE").Value)%></a></font></b></font> <font size="2"><i>(<%=(rsMyNews.Fields.Item("NEWS_SOURCE").Value)%>)</i></font></div>
                            </td>
                            <td align="right" valign="middle"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                              <b>Dated:</b> <%=(rsMyNews.Fields.Item("NEWS_DATE").Value)%></font></td>
                          </tr>
                        </table>
                        <% 
  RepeatMyNews__index=RepeatMyNews__index+1
  RepeatMyNews__numRows=RepeatMyNews__numRows-1
  rsMyNews.MoveNext()
Wend
%>
                      </td>
                    </tr>
                    <tr> 
                      <td align="left" valign="top" bgcolor="#000000" colspan="2"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td height="40"> 
      <!--#include file="../includes/inc_footer.asp" -->
    </td>
  </tr>
</table>
</body>
</html>
<%
rsMyMsg.Close()
%>
<%
rsMyLinks.Close()
%>
<%
rsMyNews.Close()
%>
