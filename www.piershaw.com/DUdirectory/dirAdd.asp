<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/connDUportal.asp" -->
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
  MM_editTable = "LINKS"
  MM_editRedirectUrl = "dirAdded.asp"
  MM_fieldsStr  = "LINK_NAME|value|LINK_URL|value|LINK_DESC|value|SUB_ID|value|CAT_ID|value|LINK_ADDER|value"
  MM_columnsStr = "LINK_NAME|',none,''|LINK_URL|',none,''|LINK_DESC|',none,''|SUB_ID|none,none,NULL|CAT_ID|none,none,NULL|LINK_ADDER|',none,''"

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
set rsDirAddCat = Server.CreateObject("ADODB.Recordset")
rsDirAddCat.ActiveConnection = MM_connDUportal_STRING
rsDirAddCat.Source = "SELECT * FROM LINK_CATS ORDER BY CAT_NAME ASC"
rsDirAddCat.CursorType = 0
rsDirAddCat.CursorLocation = 2
rsDirAddCat.LockType = 3
rsDirAddCat.Open()
rsDirAddCat_numRows = 0
%>
<%
Dim rsDirAddSub__MMColParam
rsDirAddSub__MMColParam = "1"
if (Request.QueryString("cat_id") <> "") then rsDirAddSub__MMColParam = Request.QueryString("cat_id")
%>
<%
set rsDirAddSub = Server.CreateObject("ADODB.Recordset")
rsDirAddSub.ActiveConnection = MM_connDUportal_STRING
rsDirAddSub.Source = "SELECT * FROM LINK_SUBS WHERE CAT_ID = " + Replace(rsDirAddSub__MMColParam, "'", "''") + " ORDER BY SUB_NAME ASC"
rsDirAddSub.CursorType = 0
rsDirAddSub.CursorLocation = 2
rsDirAddSub.LockType = 3
rsDirAddSub.Open()
rsDirAddSub_numRows = 0
%>
<%
Dim rsCatSub__catID
rsCatSub__catID = "999999"
if (Request.QueryString("cat_id")   <> "") then rsCatSub__catID = Request.QueryString("cat_id")  
%>
<%
Dim rsCatSub__subID
rsCatSub__subID = "999999"
if (Request.QueryString("sub_id")   <> "") then rsCatSub__subID = Request.QueryString("sub_id")  
%>
<%
set rsCatSub = Server.CreateObject("ADODB.Recordset")
rsCatSub.ActiveConnection = MM_connDUportal_STRING
rsCatSub.Source = "SELECT CAT_NAME, SUB_NAME  FROM LINK_CATS AS C, LINK_SUBS AS S  WHERE C.CAT_ID = " + Replace(rsCatSub__catID, "'", "''") + " AND S.SUB_ID = " + Replace(rsCatSub__subID, "'", "''") + " AND C.CAT_ID = S.CAT_ID"
rsCatSub.CursorType = 0
rsCatSub.CursorLocation = 2
rsCatSub.LockType = 3
rsCatSub.Open()
rsCatSub_numRows = 0
%>
<%
Dim HLooper1__numRows
HLooper1__numRows = -2
Dim HLooper1__index
HLooper1__index = 0
rsDirAddCat_numRows = rsDirAddCat_numRows + HLooper1__numRows
%>
<%
Dim HLooper2__numRows
HLooper2__numRows = -2
Dim HLooper2__index
HLooper2__index = 0
rsDirAddSub_numRows = rsDirAddSub_numRows + HLooper2__numRows
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
<% Response.Buffer = "true" %>
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
                      <td align="left" valign="middle" height="20" class = "bg_navigator"><font size="1"><b><font face="Verdana, Arial, Helvetica, sans-serif">&nbsp;ADD 
                        NEW LINK</font></b></font></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="top" bgcolor="#000000"><img src="../assets/horizontalBar.gif" width="5" height="1"></td>
                    </tr>
                    <tr> 
                      <td align="left" valign="top"> 
                        <table width="100%" border="0" cellspacing="2" cellpadding="2">
                          <tr> 
                            <td align="center" valign="top"> 
                              <% If Request.QueryString("action") = "cat" Then %>
                              <b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">	
                              <font color="#FF0000">Please select a category for 
                              your site:</font></font></b> <br>
                              <br>
                              <table border="0" cellspacing="5" cellpadding="5">
                                <tr> 
                                  <td align="left" valign="middle"> 
                                    <table cellpadding="5" cellspacing="5">
                                      <%
startrw = 0
endrw = HLooper1__index
numberColumns = 2
numrows = -1
while((numrows <> 0) AND (Not rsDirAddCat.EOF))
	startrw = endrw + 1
	endrw = endrw + numberColumns
 %>
                                      <tr align="center" valign="top"> 
                                        <%
While ((startrw <= endrw) AND (Not rsDirAddCat.EOF))
%>
                                        <td align="left" valign="middle"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><a href="dirAdd.asp?action=sub&cat_id=<%=(rsDirAddCat.Fields.Item("CAT_ID").Value)%>"><%=(rsDirAddCat.Fields.Item("CAT_NAME").Value)%></a></b></font> </td>
                                        <%
	startrw = startrw + 1
	rsDirAddCat.MoveNext()
	Wend
	%>
                                      </tr>
                                      <%
 numrows=numrows-1
 Wend
 %>
                                    </table>
                                  </td>
                                </tr>
                              </table>
                              <% End If %>
                              <br>
                              <% If Request.QueryString("action") = "sub" Then %>
                              <b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FF0000">Now 
                              select a sub-category for your site:</font></b> 
                              <br>
                              <br>
                              <table border="0" cellspacing="2" cellpadding="5">
                                <tr> 
                                  <td align="left" valign="top"> 
                                    <table cellpadding="5" cellspacing="5">
                                      <%
startrw = 0
endrw = HLooper2__index
numberColumns = 2
numrows = -1
while((numrows <> 0) AND (Not rsDirAddSub.EOF))
	startrw = endrw + 1
	endrw = endrw + numberColumns
 %>
                                      <tr align="center" valign="top"> 
                                        <%
While ((startrw <= endrw) AND (Not rsDirAddSub.EOF))
%>
                                        <td align="left" valign="middle"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><a href="dirAdd.asp?action=form&cat_id=<%=(rsDirAddSub.Fields.Item("CAT_ID").Value)%>&sub_id=<%=(rsDirAddSub.Fields.Item("SUB_ID").Value)%>"><%=(rsDirAddSub.Fields.Item("SUB_NAME").Value)%></a></b></font></td>
                                        <%
	startrw = startrw + 1
	rsDirAddSub.MoveNext()
	Wend
	%>
                                      </tr>
                                      <%
 numrows=numrows-1
 Wend
 %>
                                    </table>
                                  </td>
                                </tr>
                              </table>
                              <% End If %>
                              <br>
                              <br>
                              <% If Request.QueryString("action") = "form" Then %>
                              <b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FF0000">Fill 
                              out the form below with your site's information:</font></b> 
                              <form method="POST" action="<%=MM_editAction%>" name="form1">
                                <table align="center" cellpadding="5" cellspacing="5">
                                  <tr valign="baseline"> 
                                    <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">CATEGORY:</font></b></td>
                                    <td><b><font size="2" color="#0000FF" face="Verdana, Arial, Helvetica, sans-serif"><%=(rsCatSub.Fields.Item("CAT_NAME").Value)%></font></b></td>
                                  </tr>
                                  <tr valign="baseline"> 
                                    <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">SUB-CATEGORY:</font></b></td>
                                    <td><b><font size="2" color="#0000FF" face="Verdana, Arial, Helvetica, sans-serif"><%=(rsCatSub.Fields.Item("SUB_NAME").Value)%></font></b></td>
                                  </tr>
                                  <tr valign="baseline"> 
                                    <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">LINK 
                                      NAME:</font></b></td>
                                    <td> 
                                      <input type="text" name="LINK_NAME" value="" size="40" maxlength="40">
                                    </td>
                                  </tr>
                                  <tr valign="baseline"> 
                                    <td nowrap align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">LINK 
                                      URL:</font></b></td>
                                    <td> 
                                      <input type="text" name="LINK_URL" value="" size="45" maxlength="45">
                                    </td>
                                  </tr>
                                  <tr valign="baseline"> 
                                    <td nowrap align="right" valign="top"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">LINK 
                                      DESCRIPTION:</font></b></td>
                                    <td> 
                                      <textarea name="LINK_DESC" cols="45" rows="5"></textarea>
                                    </td>
                                  </tr>
                                  <tr valign="baseline"> 
                                    <td nowrap align="right"> 
                                      <input type="hidden" name="SUB_ID" value="<%= Request.QueryString("sub_id") %>">
                                      <input type="hidden" name="CAT_ID" value="<%= Request.QueryString("cat_id") %>">
                                      <input type="hidden" name="LINK_ADDER" value="<%= Session("MM_Username") %>">
                                    </td>
                                    <td> 
                                      <input type="submit" value="Add Link">
                                    </td>
                                  </tr>
                                </table>
                                <input type="hidden" name="MM_insert" value="true">
                              </form>
                              <% End If %>
                            </td>
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
rsDirAddCat.Close()
%>
<%
rsDirAddSub.Close()
%>
<%
rsCatSub.Close()
%>
