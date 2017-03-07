
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

If (CStr(Request("MM_insert")) = "Post") Then

  MM_editConnection = MM_connDUportal_STRING
  MM_editTable = "DATAS"
  MM_editRedirectUrl = "../home/type.asp?nChannel=Topics"
  MM_fieldsStr  = "DAT_NAME|value|DAT_DESCRIPTION|value|DAT_CATEGORY|value|DAT_USER|value|DAT_DATED|value|DAT_EXPIRED|value|DAT_LAST|value|DAT_APPROVED|value|DAT_PARENT|value|DAT_COUNT|value"
  MM_columnsStr = "DAT_NAME|',none,''|DAT_DESCRIPTION|',none,''|DAT_CATEGORY|none,none,NULL|DAT_USER|',none,''|DAT_DATED|',none,NULL|DAT_EXPIRED|',none,NULL|DAT_LAST|',none,NULL|DAT_APPROVED|none,none,NULL|DAT_PARENT|none,none,NULL|DAT_COUNT|none,none,NULL"

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
  
  ' Increase DAT_COUNT

set cmdHits = Server.CreateObject("ADODB.Command")
cmdHits.ActiveConnection = MM_connDUportal_STRING
cmdHits.CommandText = "UPDATE DATAS  SET DAT_COUNT = DAT_COUNT + 1, DAT_LAST = DATE()  WHERE DAT_ID = " & Request.QueryString("iData")
cmdHits.CommandType = 1
cmdHits.CommandTimeout = 0
cmdHits.Prepared = true
cmdHits.Execute()

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
Dim rsPost__MMColParam
rsPost__MMColParam = "1"
If (Request.QueryString("iData") <> "") Then 
  rsPost__MMColParam = Request.QueryString("iData")
End If
%>
<%
Dim rsPost
Dim rsPost_numRows

Set rsPost = Server.CreateObject("ADODB.Recordset")
rsPost.ActiveConnection = MM_connDUportal_STRING
rsPost.Source = "SELECT * FROM DATAS WHERE DAT_ID = " + Replace(rsPost__MMColParam, "'", "''") + ""
rsPost.CursorType = 0
rsPost.CursorLocation = 2
rsPost.LockType = 1
rsPost.Open()

rsPost_numRows = 0
%>
<%
Dim rsUser
Dim rsUser_numRows

Set rsUser = Server.CreateObject("ADODB.Recordset")
rsUser.ActiveConnection = MM_connDUportal_STRING
rsUser.Source = "SELECT * FROM USERS WHERE U_ID = '" & Session("MM_Username") & "' OR U_ID = '" & Request.Cookies("DUportalUser") & "'" 
rsUser.CursorType = 0
rsUser.CursorLocation = 2
rsUser.LockType = 1
rsUser.Open()

rsUser_numRows = 0
%>
<%
Dim rsType__MMColParam
rsType__MMColParam = "1"
If (Request.QueryString("iCat") <> "") Then 
  rsType__MMColParam = Request.QueryString("iCat")
End If
%>
<%
Dim rsType
Dim rsType_numRows

Set rsType = Server.CreateObject("ADODB.Recordset")
rsType.ActiveConnection = MM_connDUportal_STRING
rsType.Source = "SELECT * FROM CATEGORIES WHERE CAT_ID = " + Replace(rsType__MMColParam, "'", "''") + ""
rsType.CursorType = 0
rsType.CursorLocation = 2
rsType.LockType = 1
rsType.Open()

rsType_numRows = 0
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
                      <td align="left" valign="middle" class="textBoldColor"><a href="default.asp">HOME</a> 
                        &raquo;FORUMS &raquo; <a href="../home/channel.asp?iChannel=&nChannel="></a><%= UCase((rsType.Fields.Item("CAT_NAME").Value)) %> 
                        &raquo; POST </td>
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
				<form action="<%=MM_editAction%>" method="POST" name="Post" id="Post">
                <td align="left" valign="top" class="bgTable">
                    <table align="center" cellpadding="4" cellspacing="4">
                      <tr valign="baseline"> 
                        <% If rsUser.EOF And rsUser.BOF Then %>
                        <td colspan="2" align="left" valign="top" nowrap class="textRed">Please 
                          login to post a message!</td>
                        <% End If ' end rsUser.EOF And rsUser.BOF %>
                      </tr>
                      <% If Not rsUser.EOF Or Not rsUser.BOF Then %>
                      <tr valign="baseline"> 
                        <td align="right" nowrap class="textBold">TOPIC:</td>
                        <td> 
                          <% If rsPost.EOF And rsPost.BOF Then %>
                          <input name="DAT_NAME" type="text" class="form" value="" size="70" maxlength="150">
                          <% End If ' end rsPost.EOF And rsPost.BOF %> <% If Not rsPost.EOF Or Not rsPost.BOF Then %>
                          <input name="DAT_NAME" type="text" class="form" value="Re: <%=(rsPost.Fields.Item("DAT_NAME").Value)%>" size="70" maxlength="150">
                          <% End If ' end Not rsPost.EOF Or NOT rsPost.BOF %> </td>
                      </tr>
                      <tr align="right" valign="middle" class="textBold"> 
                        <td colspan="2" valign="top" nowrap class="textRed">Use 
                          [ and ] for writing message in HTML</td>
                      </tr>
                      <tr valign="baseline"> 
                        <td align="right" valign="top" nowrap class="textBold">MESSAGE:</td>
                        <td> <textarea name="DAT_DESCRIPTION" cols="70" rows="20" class="form"></textarea> 
                        </td>
                      </tr>
                      <tr valign="baseline"> 
                        <td nowrap align="right">&nbsp;</td>
                        <td> <input name="Post" type="submit" class="button" id="Post" value="Post"> 
                        </td>
                      </tr>
                      <% End If ' end Not rsUser.EOF Or NOT rsUser.BOF %>
                    </table>
                    <input type="hidden" name="DAT_CATEGORY" value="<%= Request.QueryString("iCat") %>">
					 <% If Not rsUser.EOF Or Not rsUser.BOF Then %>
                    <input name="DAT_USER" type="hidden" value="<%=(rsUser.Fields.Item("U_ID").Value)%>">
					<% End If %>
                    <input type="hidden" name="DAT_DATED" value="<%= date() %>">
                    <input type="hidden" name="DAT_EXPIRED" value="1/1/3000">
                    <input type="hidden" name="DAT_LAST" value="<%= date() %>">
                    <input type="hidden" name="DAT_APPROVED" value="1">
                    <input type="hidden" name="DAT_PARENT" value="<%= Request.QueryString("iData") %>">
                    <input type="hidden" name="DAT_COUNT" value="0">
                    <input type="hidden" name="MM_insert" value="Post">
                 
                 </td>
				  </form>
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
rsPost.Close()
Set rsPost = Nothing
%>
<%
rsUser.Close()
Set rsUser = Nothing
%>
<%
rsType.Close()
Set rsType = Nothing
%>
