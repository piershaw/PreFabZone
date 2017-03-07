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
<!--#include file="../ScriptLibrary/incResizeAddOn.asp" -->
<!--#include file="../ScriptLibrary/incPUAddOn.asp" -->
<!--#include file="../ScriptLibrary/incPureUpload.asp" -->

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
'*** Pure ASP File Upload 2.1.3
Dim GP_uploadAction,UploadQueryString
PureUploadSetup
If (CStr(Request.QueryString("GP_upload")) <> "") Then
  Dim pau_thePath,pau_Extensions,pau_Form,pau_Redirect,pau_storeType,pau_sizeLimit,pau_nameConflict,pau_requireUpload,pau_minWidth,pau_minHeight,pau_maxWidth,pau_maxHeight,pau_saveWidth,pau_saveHeight,pau_timeout,pau_progressBar,pau_progressWidth,pau_progressHeight
  pau_thePath = """../pictures"""
  pau_Extensions = "GIF,JPG,JPEG,BMP,PNG"
  pau_Form = "add_data"
  pau_Redirect = ""
  pau_storeType = "file"
  pau_sizeLimit = ""
  pau_nameConflict = "uniq"
  pau_requireUpload = "false"
  pau_minWidth = ""
  pau_minHeight = "" 
  pau_maxWidth = ""
  pau_maxHeight = ""
  pau_saveWidth = ""
  pau_saveHeight = ""
  pau_timeout = "600"
  pau_progressBar = ""
  pau_progressWidth = "300"
  pau_progressHeight = "100"
  
  Dim RequestBin, UploadRequest
  CheckPureUploadVersion 2.13
  ProcessUpload pau_thePath,pau_Extensions,pau_Redirect,pau_storeType,pau_sizeLimit,pau_nameConflict,pau_requireUpload,pau_minWidth,pau_minHeight,pau_maxWidth,pau_maxHeight,pau_saveWidth,pau_saveHeight,pau_timeout
end if
%>
<%
' *** Smart Image Processor 1.1.2
If (CStr(Request.QueryString("GP_upload")) <> "") Then
  Dim RUF_Component, RUF_DotNetResize, RUF_ResizeImages, RUF_AspectImages, RUF_maxWidth, RUF_maxHeight, RUF_Quality, RUF_MakeThumb, RUF_AspectThumb, RUF_Suffix, RUF_maxWidthThumb, RUF_maxHeightThumb, RUF_QualityThumb, RUF_RedirectURL
  RUF_Component = "AUTO"
  RUF_DotNetResize = "../ScriptLibrary/ResizeImage.aspx"
  RUF_ResizeImages = true
  RUF_AspectImages = true
  RUF_maxWidth = "400"
  RUF_maxHeight = "400"  
  RUF_Quality = "80"
  RUF_MakeThumb = true
  RUF_AspectThumb = true
  RUF_Suffix = "_small"
  RUF_maxWidthThumb = "130"
  RUF_maxHeightThumb = "130"
  RUF_QualityThumb = "70"
  RUF_RedirectURL = ""
  if RUF_ResizeImages then
    ResizeUploadedFiles RUF_Component, RUF_DotNetResize, pau_thePath, "", RUF_maxWidth, RUF_maxHeight, RUF_Quality, true, pau_saveWidth, pau_saveHeight, RUF_AspectImages, pau_nameConflict, ""
  end if
  if RUF_MakeThumb then
    ResizeUploadedFiles RUF_Component, RUF_DotNetResize, pau_thePath, RUF_Suffix, RUF_maxWidthThumb, RUF_maxHeightThumb, RUF_QualityThumb, false, pau_saveWidth, pau_saveHeight, RUF_AspectThumb, pau_nameConflict, ""
  end if
  if RUF_RedirectURL <> "" then
    Response.Redirect RUF_RedirectURL
  end if
end if
%>
<%
' *** Edit Operations: (Modified for File Upload) declare variables

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
If (UploadQueryString <> "") Then
  MM_editAction = MM_editAction & "?" & UploadQueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Update Record: (Modified for File Upload) set variables

If (CStr(UploadFormRequest("MM_update")) = "add_data" And CStr(UploadFormRequest("MM_recordId")) <> "") Then

  MM_editConnection = MM_connDUportal_STRING
  MM_editTable = "DATAS"
  MM_editColumn = "DAT_ID"
  MM_recordId = "" + UploadFormRequest("MM_recordId") + ""
  MM_editRedirectUrl = "portal.asp"
  MM_fieldsStr  = "TYPE|value|NAME|value|URL|value|PICTURE|value|LOCATION|value|DATED|value|QUANTITY|value|PRICE|value|SHIP|value|DESCRIPTION|value"
  MM_columnsStr = "DAT_CATEGORY|none,none,NULL|DAT_NAME|',none,''|DAT_URL|',none,''|DAT_PICTURE|',none,''|DAT_LOCATION|',none,''|DAT_DATED|',none,NULL|DAT_QUANTITY|none,none,NULL|DAT_PRICE|none,none,NULL|DAT_SHIP|none,none,NULL|DAT_DESCRIPTION|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_columnsStr = FixColumnsForUpload(MM_fieldsStr,MM_columnsStr)
  MM_fieldsStr = FixFieldsForUpload(MM_fieldsStr,MM_columnsStr)
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(UploadFormRequest(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And UploadQueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And UploadQueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & UploadQueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & UploadQueryString
    End If
  End If

End If
%>
<%
' *** Delete Old Files Before Update Record 1.6.0
If (CStr(UploadFormRequest("MM_update")) <> "" And CStr(UploadFormRequest("MM_recordId")) <> "") Then
  Dim DOF_path, DOF_suffix
  DOF_path = "../pictures"
  DOF_suffix = "_small"
  DeleteFileBeforeUpdate DOF_path,MM_fieldsStr,MM_columnsStr,MM_editConnection,MM_editTable,MM_editColumn,MM_recordId,DOF_suffix
end if
%>
<%
' *** Update Record: (Modified for File Upload) construct a sql update statement and execute it

If (CStr(UploadFormRequest("MM_update")) <> "" And CStr(UploadFormRequest("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
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
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & ", DAT_APPROVED = 0, DAT_USER = '" & rsUser.Fields.Item("U_ID").Value & "'"
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId & " AND DAT_USER = '" & rsUser.Fields.Item("U_ID").Value & "'"

  If (Not MM_abortEdit) Then
    ' execute the update
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
Dim rsTypeListing__MMColParam
rsTypeListing__MMColParam = "0"
If (Request.QueryString("iChannel") <> "") Then 
  rsTypeListing__MMColParam = Request.QueryString("iChannel")
End If
%>
<%
Dim rsTypeListing
Dim rsTypeListing_numRows

Set rsTypeListing = Server.CreateObject("ADODB.Recordset")
rsTypeListing.ActiveConnection = MM_connDUportal_STRING
rsTypeListing.Source = "SELECT * FROM CATEGORIES, CHANNELS WHERE CAT_CHANNEL = CHA_ID AND CAT_CHANNEL = " + Replace(rsTypeListing__MMColParam, "'", "''") + " ORDER BY CAT_NAME ASC"
rsTypeListing.CursorType = 0
rsTypeListing.CursorLocation = 2
rsTypeListing.LockType = 1
rsTypeListing.Open()

rsTypeListing_numRows = 0
%>


<%
Dim rsEdit__MMColParam
rsEdit__MMColParam = "1"
If (Request.QueryString("iData") <> "") Then 
  rsEdit__MMColParam = Request.QueryString("iData")
End If
%>
<%
Dim rsEdit
Dim rsEdit_numRows

Set rsEdit = Server.CreateObject("ADODB.Recordset")
rsEdit.ActiveConnection = MM_connDUportal_STRING
rsEdit.Source = "SELECT * FROM DATAS WHERE DAT_ID = " + Replace(rsEdit__MMColParam, "'", "''") + ""
rsEdit.CursorType = 0
rsEdit.CursorLocation = 2
rsEdit.LockType = 1
rsEdit.Open()

rsEdit_numRows = 0
%>
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
<link href="../assets/DUportal.css" rel="stylesheet" type="text/css">
<script language="javascript" src="../ScriptLibrary/incPureUpload.js"></script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#003399">
              <tr> 
                <td align="left" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../assets/bg_header.gif">
                    <tr> 
                      <td width="10"><img src="../assets/header_end_left.gif"></td>
                      <td align="left" valign="middle" class="textBoldColor">EDIT 
                      </td>
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
                      <td align="left" valign="top">
					  
					  
					  
					  
					  <table width="100%" border="0" cellspacing="2" cellpadding="2">
                          <tr>
                            <form ACTION="<%=MM_editAction%>" METHOD="post" enctype="multipart/form-data" name="add_data" id="add_data" onSubmit="checkFileUpload(this,'GIF,JPG,JPEG,BMP,PNG',false,'','','','','','','');return document.MM_returnValue">
                              
                              <td align="left" valign="top"> <table align="center" cellpadding="3" cellspacing="3" class="textBold">
                                  <% If Not rsTypeListing.EOF Or Not rsTypeListing.BOF Then %>
                                  <tr valign="middle"> 
                                    <td align="right" nowrap>CATEGORY:</td>
                                    <td> 
                                      <select name="TYPE" class="form" id="TYPE">
                                        <%
While (NOT rsTypeListing.EOF)
%>
                                        <option value="<%=(rsTypeListing.Fields.Item("CAT_ID").Value)%>" <%If (Not isNull((rsEdit.Fields.Item("DAT_CATEGORY").Value))) Then If (CStr(rsTypeListing.Fields.Item("CAT_ID").Value) = CStr((rsEdit.Fields.Item("DAT_CATEGORY").Value))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(rsTypeListing.Fields.Item("CAT_NAME").Value)%></option>
                                        <%
  rsTypeListing.MoveNext()
Wend
If (rsTypeListing.CursorType > 0) Then
  rsTypeListing.MoveFirst
Else
  rsTypeListing.Requery
End If
%>
                                      </select> </td>
                                  </tr>
                                  <tr valign="middle"> 
                                    <td align="right" nowrap>NAME:</td>
                                    <td> 
                                      <input name="NAME" type="text" class="form" value="<%=(rsEdit.Fields.Item("DAT_NAME").Value)%>" size="60"> 
                                    </td>
                                  </tr>
                                  <% If Request.QueryString("nChannel") <> "Pictures" And Request.QueryString("nChannel") <> "Topics"Then %>
                                  <tr valign="middle"> 
                                    <td align="right" nowrap>URL:</td>
                                    <td> 
                                      <input name="URL" type="text" class="form" value="<%=(rsEdit.Fields.Item("DAT_URL").Value)%>" size="60"> 
                                    </td>
                                  </tr>
                                  <% Else %>
                                  <input name="URL" type="hidden" value="0">
                                  <% End If %>
                                  <% If Request.QueryString("nChannel") <> "Topics"Then %>
                                  <tr valign="middle"> 
                                    <td colspan="2" align="left" nowrap class="textRed">Leave 
                                      PICTURE blank if keeping the existing one</td>
                                  </tr>
                                  <tr valign="middle"> 
                                    <td align="right" nowrap>PICTURE:</td>
                                    <td> 
                                      <input name="PICTURE" type="file" class="form" id="PICTURE" onChange="checkOneFileUpload(this,'GIF,JPG,JPEG,BMP,PNG',false,'','','','','','','')" value="<%=(rsEdit.Fields.Item("DAT_PICTURE").Value)%>" size="50"></td>
                                  </tr>
                                  <% Else %>
                                  <input name="PICTURE" type="hidden" value="">
                                  <% End If %>
                                  <% If Request.QueryString("nChannel") = "Events" Or Request.QueryString("nChannel") = "Businesses" Or  Request.QueryString("nChannel") = "Ads" Then %>
                                  <tr valign="middle"> 
                                    <td align="right" nowrap> 
                                      LOCATION:</td>
                                    <td> 
                                      <input name="LOCATION" type="text" class="form" value="<%=(rsEdit.Fields.Item("DAT_LOCATION").Value)%>" size="60"> 
                                    </td>
                                  </tr>
                                  <% Else %>
                                  <input name="LOCATION" type="hidden" value="0">
                                  <% End If %>
                                  <tr valign="middle"> 
                                    <td align="right" nowrap> 
                                      DATE:</td>
                                    <td> 
                                      <input name="DATED" type="text" class="form" value="<%=(rsEdit.Fields.Item("DAT_DATED").Value)%>" size="15"> 
                                    </td>
                                  </tr>
                                  <% If Request.QueryString("nChannel") = "Products" OR Request.QueryString("nChannel") = "Ads" Then %>
                                  <tr valign="middle"> 
                                    <td align="right" nowrap>QUANTITY:</td>
                                    <td> 
                                      <input name="QUANTITY" type="text" class="form" id="QUANTITY" value="<%=(rsEdit.Fields.Item("DAT_QUANTITY").Value)%>" size="5"> 
                                    </td>
                                  </tr>
                                  <% Else %>
                                  <input name="QUANTITY" type="hidden" value="0">
                                  <% End If %>
                                  <% If Request.QueryString("nChannel") = "Products" OR Request.QueryString("nChannel") = "Ads"Then %>
                                  <tr valign="middle"> 
                                    <td align="right" nowrap>PRICE:</td>
                                    <td> 
                                      <input name="PRICE" type="text" class="form" value="<%=(rsEdit.Fields.Item("DAT_PRICE").Value)%>" size="30"> 
                                    </td>
                                  </tr>
                                  <tr valign="middle"> 
                                    <td align="right" nowrap>SHIPPING 
                                      COST:</td>
                                    <td> 
                                      <input name="SHIP" type="text" class="form" value="<%=(rsEdit.Fields.Item("DAT_SHIP").Value)%>" size="30"> 
                                    </td>
                                  </tr>
                                  <% Else %>
                                  <input name="PRICE" type="hidden" value="0">
                                  <input name="SHIP" type="hidden" value="0">
                                  <% End If %>
                                  <tr valign="middle"> 
                                    <td align="right" valign="top" nowrap>DESCRIPTION:</td>
                                    <td> 
                                      <textarea name="DESCRIPTION" cols="60" rows="15" class="form"><%=(rsEdit.Fields.Item("DAT_DESCRIPTION").Value)%></textarea> 
                                    </td>
                                  </tr>
                                  <tr valign="middle"> 
                                    <td align="right" nowrap> 
                                     
                                    </td>
                                    <td> 
                                      <input name="Save" type="submit" class="button" id="Save" onClick="MM_validateForm('NAME','','R','URL','','R','LOCATION','','R','DATED','','R','QUANTITY','','RisNum','PRICE','','RisNum','SHIP','','RisNum','DESCRIPTION','','R');return document.MM_returnValue" value="Save Changes"> 
                                    </td>
                                  </tr>
                                  <% End If ' end Not rsTypeListing.EOF Or NOT rsTypeListing.BOF %>
                                </table></td>
                            
                              <input type="hidden" name="MM_update" value="add_data">
                              <input type="hidden" name="MM_recordId" value="<%= rsEdit.Fields.Item("DAT_ID").Value %>">
                            </form>
                             
                          </tr>
                        </table>
						
						
						
						
						
					  
					  </td>
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
rsTypeListing.Close()
Set rsTypeListing = Nothing
%>
<%
rsUser.Close()
Set rsUser = Nothing
%>
<%
rsEdit.Close()
Set rsEdit = Nothing
%>
