<%@LANGUAGE="VBSCRIPT"%>
<% Response.Buffer = "true" %>
<!--#include file="../Connections/connDUportal.asp" -->
<%
if(Request.QueryString("for_id") <> "") then cmdFor_Read_Counter__varFOR_ID = Request.QueryString("for_id")
%>
<%
if(Request.QueryString("msg_id") <> "") then cmdMsg_Read_Counter__varMSG_ID = Request.QueryString("msg_id")
%>
<%
Dim rsMsg__varID
rsMsg__varID = "83"
if (Request.QueryString("msg_id")    <> "") then rsMsg__varID = Request.QueryString("msg_id")   
%>
<%
set rsMsg = Server.CreateObject("ADODB.Recordset")
rsMsg.ActiveConnection = MM_connDUportal_STRING
rsMsg.Source = "SELECT MSG_REPLY_COUNT, MSG_DATE, MSG_AUTHOR, MSG_SUBJECT, MSG_BODY, FORUMS.FOR_NAME,FORUMS.FOR_ID,  U_EMAIL, (SELECT COUNT(*) FROM MESSAGES WHERE MSG_AUTHOR = U_ID) AS TOPIC_COUNT, (SELECT COUNT(*) FROM REPLIES WHERE REP_AUTHOR = MSG_AUTHOR)AS REPLY_COUNT  FROM MESSAGES INNER JOIN FORUMS ON FORUMS.FOR_ID = MESSAGES.FOR_ID, USERS  WHERE U_ID = MSG_AUTHOR AND MSG_ID = " + Replace(rsMsg__varID, "'", "''") + ""
rsMsg.CursorType = 0
rsMsg.CursorLocation = 2
rsMsg.LockType = 3
rsMsg.Open()
rsMsg_numRows = 0
%>
<%
Dim rsReplier__MMColParam
rsReplier__MMColParam = "1"
if (Session("MM_Username") <> "") then rsReplier__MMColParam = Session("MM_Username")
%>
<%
set rsReplier = Server.CreateObject("ADODB.Recordset")
rsReplier.ActiveConnection = MM_connDUportal_STRING
rsReplier.Source = "SELECT * FROM USERS WHERE U_ID = '" + Replace(rsReplier__MMColParam, "'", "''") + "'"
rsReplier.CursorType = 0
rsReplier.CursorLocation = 2
rsReplier.LockType = 3
rsReplier.Open()
rsReplier_numRows = 0
%>
<%
Dim rsForum__MMColParam
rsForum__MMColParam = "1"
if (Request.QueryString("for_id") <> "") then rsForum__MMColParam = Request.QueryString("for_id")
%>
<%
set rsForum = Server.CreateObject("ADODB.Recordset")
rsForum.ActiveConnection = MM_connDUportal_STRING
rsForum.Source = "SELECT * FROM FORUMS WHERE FOR_ID = " + Replace(rsForum__MMColParam, "'", "''") + ""
rsForum.CursorType = 0
rsForum.CursorLocation = 2
rsForum.LockType = 3
rsForum.Open()
rsForum_numRows = 0
%>
<%
Dim rsRepForm__varID
rsRepForm__varID = "9999"
if (Request.QueryString("msg_id")   <> "") then rsRepForm__varID = Request.QueryString("msg_id")  
%>
<%
set rsRepForm = Server.CreateObject("ADODB.Recordset")
rsRepForm.ActiveConnection = MM_connDUportal_STRING
rsRepForm.Source = "SELECT *  FROM MESSAGES INNER JOIN FORUMS ON FORUMS.FOR_ID = MESSAGES.FOR_ID  WHERE MSG_ID = " + Replace(rsRepForm__varID, "'", "''") + ""
rsRepForm.CursorType = 0
rsRepForm.CursorLocation = 2
rsRepForm.LockType = 3
rsRepForm.Open()
rsRepForm_numRows = 0
%>
<%
Dim rsRep__varID
rsRep__varID = "11"
if (Request.QueryString("msg_id")     <> "") then rsRep__varID = Request.QueryString("msg_id")    
%>
<%
set rsRep = Server.CreateObject("ADODB.Recordset")
rsRep.ActiveConnection = MM_connDUportal_STRING
rsRep.Source = "SELECT REP_DATE, REP_AUTHOR, REP_BODY, U_EMAIL, (SELECT COUNT(*) FROM MESSAGES WHERE MSG_AUTHOR = REP_AUTHOR) AS TOPIC_COUNT,  (SELECT COUNT(*) FROM REPLIES WHERE REP_AUTHOR = U_ID ) AS REPLY_COUNT  FROM REPLIES, USERS  WHERE U_ID = REP_AUTHOR AND MSG_ID = " + Replace(rsRep__varID, "'", "''") + ""
rsRep.CursorType = 0
rsRep.CursorLocation = 2
rsRep.LockType = 3
rsRep.Open()
rsRep_numRows = 0
%>
<% ' this is to increment the read counter in FORUMS table
set cmdFor_Read_Counter = Server.CreateObject("ADODB.Command")
cmdFor_Read_Counter.ActiveConnection = MM_connDUportal_STRING
cmdFor_Read_Counter.CommandText = "UPDATE FORUMS  SET FOR_READ_COUNT  = FOR_READ_COUNT + 1  WHERE FOR_ID  = " + Replace(cmdFor_Read_Counter__varFOR_ID, "'", "''") + ""
cmdFor_Read_Counter.CommandType = 1
cmdFor_Read_Counter.CommandTimeout = 0
cmdFor_Read_Counter.Prepared = true
cmdFor_Read_Counter.Execute()
%>

<% ' this is to increment the read counter in MESSAGES table
set cmdMsg_Read_Counter = Server.CreateObject("ADODB.Command")
cmdMsg_Read_Counter.ActiveConnection = MM_connDUportal_STRING
cmdMsg_Read_Counter.CommandText = "UPDATE MESSAGES  SET MSG_READ_COUNT  = MSG_READ_COUNT + 1  WHERE MSG_ID  = " + Replace(cmdMsg_Read_Counter__varMSG_ID, "'", "''") + ""
cmdMsg_Read_Counter.CommandType = 1
cmdMsg_Read_Counter.CommandTimeout = 0
cmdMsg_Read_Counter.Prepared = true
cmdMsg_Read_Counter.Execute()
%>


<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsRep_numRows = rsRep_numRows + Repeat1__numRows
%>
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>	
function DoTrimProperly(str, nNamedFormat, properly, pointed, points)
  dim strRet
  strRet = Server.HTMLEncode(str)
  strRet = replace(strRet, vbcrlf,"")
  strRet = replace(strRet, vbtab,"")
  If (LEN(strRet) > nNamedFormat) Then
    strRet = LEFT(strRet, nNamedFormat)			
    If (properly = 1) Then					
      Dim TempArray								
      TempArray = split(strRet, " ")	
      Dim n
      strRet = ""
      for n = 0 to Ubound(TempArray) - 1
        strRet = strRet & " " & TempArray(n)
      next
    End If
    If (pointed = 1) Then
      strRet = strRet & points
    End If
  End If
  DoTrimProperly = strRet
End Function
</SCRIPT>
<%
Function DoWhiteSpace(str)
DoWhiteSpace = (Replace(str, vbCrlf, "<br>"))
End Function
%>
<%
Public Function LinkURLs(byVal strIn)
	'############################################
	'# - Creates an HTML link for any           #
	'#   HTTP, HTTPS, FTP, FTPS, NEWS or MAILTO #
	'#   path in a string.                      #
	'# - REQUIRES VBSCRIPT SCRIPTING ENGINE 5.5 #
	'############################################
	dim re, sOut

	set re = New RegExp
	re.global = true
	re.ignorecase = true
	're.multiline = true

	' pattern that parses and finds any Internet URLs:
	re.pattern = _
		"((mailto\:|(news|(ht|f)tp(s?))\://){1}\S+)"

	' replace method of RegExp object uses a remembered
	' pattern denoted as $1 to link the found URL(s)
	sOut = re.replace( strIn, _
		"<A HREF=""$1"" TARGET=""_new"">$1</A>")

	set re = Nothing
	LinkURLs = sOut
End Function
%>
<html>
<head>
<title>DUportal</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/default.css" type="text/css">
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
          <td bgcolor="#000000" width="1"><img src="../assets/verticalBar.gif"></td>
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
                      <td align="left" valign="middle"  height="20"> 
                        <div class = "links">&nbsp;<b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="default.asp">MESSAGES 
                          BOARDS</a> &gt; <a href="messages.asp?for_id=<%=(rsMsg.Fields.Item("FOR_ID").Value)%>"><%= UCase((rsMsg.Fields.Item("FOR_NAME").Value)) %></a> &gt; <%= UCase((rsMsg.Fields.Item("MSG_SUBJECT").Value)) %></font></b></div>
                      </td>
                    </tr>
                    <tr> 
                      <td align="left" valign="middle" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="middle" class = "bg_navigator" height="20"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                        &nbsp;</font></b><font size="1"><font face="Verdana, Arial, Helvetica, sans-serif"><b>TOPIC 
                        DETAIL </b></font></font></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                    </tr>
                    <tr valign="top" align="center"> 
                      <td align="left"> 
                        <div class = "links"> 
                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td align="left" valign="top"> 
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr align="left" valign="top"> 
                                    <td width="150"> 
                                      <li><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="mailto:<%=(rsMsg.Fields.Item("U_EMAIL").Value)%>"><%=(rsMsg.Fields.Item("MSG_AUTHOR").Value)%></a></font></li>
                                      <li><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsMsg.Fields.Item("TOPIC_COUNT").Value)%> topics</font></li>
                                      <li><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsMsg.Fields.Item("REPLY_COUNT").Value)%> replies</font></li>
                                    </td>
                                    <td bgcolor="#000000" width="1"><img src="../assets/verticalBar.gif" width="1" height="5"></td>
                                    <td> 
                                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                        <tr> 
                                          <td align="left" valign="middle" height="20"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><i>&nbsp;Posted 
                                            by <%=(rsMsg.Fields.Item("MSG_AUTHOR").Value)%> on <%=(rsMsg.Fields.Item("MSG_DATE").Value)%> with <%=(rsMsg.Fields.Item("MSG_REPLY_COUNT").Value)%> replies</i></font></td>
                                        </tr>
                                        <tr> 
                                          <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                                        </tr>
                                        <tr> 
                                          <td align="right" valign="top"> 
                                            <table width="98%" border="0" cellspacing="3" cellpadding="3">
                                              <tr> 
                                                <td align="left" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=DoWhiteSpace(LinkURLs(rsMsg.Fields.Item("MSG_BODY").Value))%></font></td>
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
                              <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                            </tr>
                            <tr> 
                              <td align="left" valign="top">&nbsp;</td>
                            </tr>
                            <tr> 
                              <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                            </tr>
                            <tr> 
                              <td align="left" valign="top"> 
                                <% 
While ((Repeat1__numRows <> 0) AND (NOT rsRep.EOF)) 
%>
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr align="left" valign="top"> 
                                    <td width="150"> 
                                      <li><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="mailto:<%=(rsRep.Fields.Item("U_EMAIL").Value)%>"><%=(rsRep.Fields.Item("REP_AUTHOR").Value)%></a></font></li>
                                      <li><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsRep.Fields.Item("TOPIC_COUNT").Value)%> topics</font></li>
                                      <li><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsRep.Fields.Item("REPLY_COUNT").Value)%> replies</font></li>
                                    </td>
                                    <td bgcolor="#000000" width="1"><img src="../assets/verticalBar.gif" width="1" height="5"></td>
                                    <td> 
                                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                        <tr> 
                                          <td align="left" valign="middle" height="20"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><i>&nbsp;Replied 
                                            by <%=(rsRep.Fields.Item("REP_AUTHOR").Value)%> on <%=(rsRep.Fields.Item("REP_DATE").Value)%></i></font></td>
                                        </tr>
                                        <tr> 
                                          <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                                        </tr>
                                        <tr> 
                                          <td align="right" valign="top"> 
                                            <table width="98%" border="0" cellspacing="3" cellpadding="3">
                                              <tr> 
                                                <td align="left" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=DoWhiteSpace(LinkURLs(rsRep.Fields.Item("REP_BODY").Value))%></font></td>
                                              </tr>
                                            </table>
                                          </td>
                                        </tr>
                                      </table>
                                    </td>
                                  </tr>
                                  <tr align="left" valign="top" bgcolor="#000000"> 
                                    <td colspan="3"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                                  </tr>
                                </table>
                                <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsRep.MoveNext()
Wend
%>
                              </td>
                            </tr>
                            <tr> 
                              <td align="left" valign="top">&nbsp;</td>
                            </tr>
                          </table>
                        </div>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td align="left" valign="top"> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="middle" class = "bg_navigator" height="20"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp;REPLY 
                        THIS TOPIC</font></b></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                    </tr>
                    <tr> 
                      <form name="POST" method="get" action="msgAdding.asp">
                        <td align="left" valign="top"> 
                          <table width="100%" border="0" cellspacing="5" cellpadding="5">
                            <tr align="left" valign="middle"> 
                              <td align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Replier:</font></b></td>
                              <td> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                                <% If Not rsReplier.EOF Or Not rsReplier.BOF Then %>
                                <input type="hidden" name="REP_AUTHOR" value="<%=(rsReplier.Fields.Item("U_ID").Value)%>">
                                <%=(rsReplier.Fields.Item("U_ID").Value)%> 
                                <% Else %>
                                <font color = "ff0000"> To reply, please 
                                login or register first.</font> 
                                <% End If ' end Not rsReplier.EOF Or NOT rsReplier.BOF %>
                                </font></b> </td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Forum:</font></b></td>
                              <td> 
                                <div class = "links"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                                  <a href="messages.asp?for_id=<%=(rsForum.Fields.Item("FOR_ID").Value)%>"><%=(rsRepForm.Fields.Item("FOR_NAME").Value)%></a></font></b></div>
                              </td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Re:</font></b></td>
                              <td><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=(rsRepForm.Fields.Item("MSG_SUBJECT").Value)%></font></b></td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td align="right" valign="top" rowspan="2"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Message:</font></b></td>
                              <td valign="top"><font size="1"><i><font face="Verdana, Arial, Helvetica, sans-serif">If 
                                your message contains HTML or ASP codes, please 
                                replace &lt; with [ and &gt; with ]. If not, your 
                                codes won't display correctly. </font> </i> </font></td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td> 
                                <textarea name="REP_BODY" cols="60" rows="10" class = "fields"></textarea>
                              </td>
                            </tr>
                            <tr align="left" valign="middle"> 
                              <td align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
                                <input type="hidden" name="MSG_ID" value="<%=(rsRepForm.Fields.Item("MSG_ID").Value)%>">
								<input type="hidden" name="FOR_ID" value="<%=(rsRepForm.Fields.Item("FOR_ID").Value)%>">
                                </font></b></td>
                              <td> 
                                <% If Not rsReplier.EOF Or Not rsReplier.BOF Then %>
                                <input type="submit" name="SUBMIT" value="REPLY" class = "buttons">
                                <% End If %>
                              </td>
                            </tr>
                          </table>
                        </td>
                      </form>
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
rsMsg.Close()
%>
<%
rsReplier.Close()
%>
<%
rsForum.Close()
%>
<%
rsRepForm.Close()
%>
<%
rsRep.Close()
%>

