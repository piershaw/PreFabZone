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
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) = "edit_type" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_connDUportal_STRING
  MM_editTable = "CATEGORIES"
  MM_editColumn = "CAT_ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "types.asp"

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
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
  
  
  set cmdDeleteData = Server.CreateObject("ADODB.Command")
cmdDeleteData.ActiveConnection = MM_connDUportal_STRING
cmdDeleteData.CommandText = "DELETE FROM DATAS  WHERE DAT_CATEGORY = " & Request("MM_recordId")
cmdDeleteData.CommandType = 1
cmdDeleteData.CommandTimeout = 0
cmdDeleteData.Prepared = true
cmdDeleteData.Execute()



    ' execute the delete
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
Dim rsChannelListing
Dim rsChannelListing_numRows

Set rsChannelListing = Server.CreateObject("ADODB.Recordset")
rsChannelListing.ActiveConnection = MM_connDUportal_STRING
rsChannelListing.Source = "SELECT * FROM CHANNELS ORDER BY CHA_MENU ASC"
rsChannelListing.CursorType = 0
rsChannelListing.CursorLocation = 2
rsChannelListing.LockType = 1
rsChannelListing.Open()

rsChannelListing_numRows = 0
%>
<%
Dim rsTypeEdit__MMColParam
rsTypeEdit__MMColParam = "1"
If (Request.QueryString("iCat") <> "") Then 
  rsTypeEdit__MMColParam = Request.QueryString("iCat")
End If
%>
<%
Dim rsTypeEdit
Dim rsTypeEdit_numRows

Set rsTypeEdit = Server.CreateObject("ADODB.Recordset")
rsTypeEdit.ActiveConnection = MM_connDUportal_STRING
rsTypeEdit.Source = "SELECT * FROM CATEGORIES WHERE CAT_ID = " + Replace(rsTypeEdit__MMColParam, "'", "''") + ""
rsTypeEdit.CursorType = 0
rsTypeEdit.CursorLocation = 2
rsTypeEdit.LockType = 1
rsTypeEdit.Open()

rsTypeEdit_numRows = 0
%>
 <link href="../assets/DUportal.css" rel="stylesheet" type="text/css"> 
<script language="JavaScript" type="text/JavaScript">
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
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#003399">
              <tr> 
                <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../assets/bg_header.gif">
                    <tr> 
                      <td width="10"><img src="../assets/header_end_left.gif"></td>
                      <td align="left" valign="middle" class="textBoldColor">DELETE 
                        CATEGORY</td>
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
                      <td align="left" valign="top"><table width="100%" border="0" cellspacing="2" cellpadding="2">
                          <tr>
                            <td align="center" valign="middle"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="bgTable">
                                <tr>
                                  <td align="left" valign="top" class="textRed">If 
                                    you delete this type, all of its data records 
                                    will be deleted. Be sure to check if this 
                                    channel has any child records.</td>
                                </tr>
                                <tr> 
                                  <form ACTION="<%=MM_editAction%>" METHOD="POST" name="edit_type">
                                    <td align="left" valign="top"> <table align="center" cellpadding="4" cellspacing="1">
                                        <tr valign="baseline"> 
                                          <td align="right" valign="middle" nowrap class="textBold">CHANNEL:</td>
                                          <td> <select name="CAT_CHANNEL" class="form" id="CAT_CHANNEL">
                                              <%
While (NOT rsChannelListing.EOF)
%>
                                              <option value="<%=(rsChannelListing.Fields.Item("CHA_ID").Value)%>" <%If (Not isNull((rsTypeEdit.Fields.Item("CAT_CHANNEL").Value))) Then If (CStr(rsChannelListing.Fields.Item("CHA_ID").Value) = CStr((rsTypeEdit.Fields.Item("CAT_CHANNEL").Value))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(rsChannelListing.Fields.Item("CHA_MENU").Value)%></option>
                                              <%
  rsChannelListing.MoveNext()
Wend
If (rsChannelListing.CursorType > 0) Then
  rsChannelListing.MoveFirst
Else
  rsChannelListing.Requery
End If
%>
                                            </select> </td>
                                        </tr>
                                        <tr valign="baseline"> 
                                          <td align="right" valign="middle" nowrap class="textBold">CATEGORY 
                                            NAME:</td>
                                          <td> <input name="CAT_NAME" type="text" class="form" value="<%=(rsTypeEdit.Fields.Item("CAT_NAME").Value)%>" size="50" maxlength="50"></td>
                                        </tr>
                                        <tr valign="baseline"> 
                                          <td align="right" valign="middle" nowrap class="textBold">CATEGORY 
                                            DESCRIPTION:</td>
                                          <td> <input name="CAT_DESCRIPTION" type="text" class="form" value="<%=(rsTypeEdit.Fields.Item("CAT_DESCRIPTION").Value)%>" size="50" maxlength="50"></td>
                                        </tr>
                                        <tr valign="baseline"> 
                                          <td nowrap align="right"> </td>
                                          <td> <input type="submit" class="button" onClick="MM_validateForm('CAT_NAME','','R','CAT_DESCRIPTION','','R');return document.MM_returnValue" value="Delete This Type"> 
                                          </td>
                                        </tr>
                                      </table></td>
                                    <input type="hidden" name="MM_delete" value="edit_type">
                                    <input type="hidden" name="MM_recordId" value="<%= rsTypeEdit.Fields.Item("CAT_ID").Value %>">
                                  </form>
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
rsChannelListing.Close()
Set rsChannelListing = Nothing
%>
<%
rsTypeEdit.Close()
Set rsTypeEdit = Nothing
%>
