<!--#include file="../Connections/connDUportal.asp" -->
<!--#include file="inc_restriction.asp" -->

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
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
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

If (CStr(Request("MM_insert")) = "addAnswer") Then

  MM_editConnection = MM_connDUportal_STRING
  MM_editTable = "DATAS"
  MM_editRedirectUrl = "pollsAnswers.asp"
  MM_fieldsStr  = "DAT_NAME|value|DAT_PARENT|value|DAT_CATEGORY|value|DAT_APPROVED|value|DAT_DATED|value|DAT_EXPIRED|value|DAT_USER|value"
  MM_columnsStr = "DAT_NAME|',none,''|DAT_PARENT|none,none,NULL|DAT_CATEGORY|none,none,NULL|DAT_APPROVED|none,none,NULL|DAT_DATED|',none,NULL|DAT_EXPIRED|',none,NULL|DAT_USER|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
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

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
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
set rsQuestion = Server.CreateObject("ADODB.Recordset")
rsQuestion.ActiveConnection = MM_connDUportal_STRING
rsQuestion.Source = "SELECT * FROM DATAS, CATEGORIES, CHANNELS WHERE DAT_CATEGORY = CAT_ID AND CAT_CHANNEL = CHA_ID AND CHA_NAME ='POLLS' AND DAT_PARENT=0 ORDER BY DAT_ID DESC"
rsQuestion.CursorType = 0
rsQuestion.CursorLocation = 2
rsQuestion.LockType = 3
rsQuestion.Open()
rsQuestion_numRows = 0
Session("poll_id") = rsQuestion.Fields.Item("DAT_ID").Value
%>

<%
Dim rsAnswer__MMColParam
rsAnswer__MMColParam = "1"
If (Session("poll_id") <> "") Then 
  rsAnswer__MMColParam = Session("poll_id")
End If
%>
<%
set rsAnswer = Server.CreateObject("ADODB.Recordset")
rsAnswer.ActiveConnection = MM_connDUportal_STRING
rsAnswer.Source = "SELECT * FROM DATAS WHERE DAT_PARENT = " + Replace(rsAnswer__MMColParam, "'", "''") + " ORDER BY DAT_ID ASC"
rsAnswer.CursorType = 0
rsAnswer.CursorLocation = 2
rsAnswer.LockType = 3
rsAnswer.Open()
rsAnswer_numRows = 0
%>

<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsAnswer_numRows = rsAnswer_numRows + Repeat1__numRows
%>
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
                      <td align="left" valign="middle" class="textBoldColor">CHOICES</td>
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
                      <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="bgTable">
                          <tr> 
                            <td align="left" valign="top"><table width="100%" border="0" cellspacing="5" cellpadding="5">
                                <tr> 
                                  <td class="textBold">Add Choices:</td>
                                </tr>
                                <tr> 
                                  <td align="left" valign="top"> <table width="100%" border="0" cellpadding="5" cellspacing="5" class="text">
                                      <tr> 
                                        <td align="right"><strong>Question:</strong></td>
                                        <td><%=(rsQuestion.Fields.Item("DAT_NAME").Value)%></td>
                                      </tr>
                                      <% If Not rsAnswer.EOF Or Not rsAnswer.BOF Then %>
                                      <tr> 
                                        <td align="right" valign="top"><strong>Choies: 
                                          </strong></td>
                                        <td> 
                                          <% 
While ((Repeat1__numRows <> 0) AND (NOT rsAnswer.EOF)) 
%>
                                          <table width="100%" border="0" cellspacing="2" cellpadding="2">
                                            <tr> 
                                              <td align="left" valign="top" class="text"><%=(rsAnswer.Fields.Item("DAT_NAME").Value)%></td>
                                            </tr>
                                          </table>
                                          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsAnswer.MoveNext()
Wend
%> </td>
                                      </tr>
                                      <% End If ' end Not rsAnswer.EOF Or NOT rsAnswer.BOF %>
                                      <tr> 
                                        <form ACTION="<%=MM_editAction%>" METHOD="POST" name="addAnswer">
                                          <td align="right"><strong>Add A Choice: 
                                            </strong></td>
                                          <td> <input name="DAT_NAME" type="text" class="form" id="DAT_NAME" size="40" maxlength="40"> 
                                            <input name="DAT_PARENT" type="hidden" id="DAT_PARENT" value="<%=((rsQuestion.Fields.Item("DAT_ID").Value))%>"> 
                                            <input name="Add" type="submit" class="button" value="Add Choice"> 
                                            <input name="Done" type="button" class="button" onClick="MM_goToURL('parent','polls.asp');return document.MM_returnValue" value="Finish"> 
                                            <input name="DAT_CATEGORY" type="hidden" id="DAT_CATEGORY" value="<%=(rsQuestion.Fields.Item("DAT_CATEGORY").Value)%>"> 
                                            <input name="DAT_APPROVED" type="hidden" id="DAT_APPROVED" value="1"> 
                                            <input name="DAT_DATED" type="hidden" id="DAT_DATED" value="<%= date() %>"> 
                                            <input name="DAT_EXPIRED" type="hidden" id="DAT_EXPIRED" value="1/1/3000"> 
                                            <input name="DAT_USER" type="hidden" id="DAT_USER" value="admin"> 
                                          </td>
                                          <input type="hidden" name="MM_insert" value="addAnswer">
                                        </form>
                                      </tr>
                                    </table></td>
                                </tr>
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
rsQuestion.Close()
%>
<%
rsAnswer.Close()
%>