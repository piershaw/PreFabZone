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
Dim rsDropDownsearch
Dim rsDropDownsearch_numRows

Set rsDropDownsearch = Server.CreateObject("ADODB.Recordset")
rsDropDownsearch.ActiveConnection = MM_connDUportal_STRING
rsDropDownsearch.Source = "SELECT DISTINCT *  FROM CHANNELS  WHERE CHA_ACTIVE=1  ORDER BY CHA_MENU ASC"
rsDropDownsearch.CursorType = 0
rsDropDownsearch.CursorLocation = 2
rsDropDownsearch.LockType = 1
rsDropDownsearch.Open()

rsDropDownsearch_numRows = 0
%>
<%
Dim rsDropDownsearch__numRows
Dim rsDropDownsearch__index

rsDropDownsearch__numRows = -1
rsDropDownsearch__index = 0
rsDropDownsearch_numRows = rsDropDownsearch_numRows + rsDropDownsearch__numRows
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
                      <td align="left" valign="middle" class="textBoldColor">SEARCH</td>
                      <td width="28" align="right" valign="middle"><img src="../assets/header_end_right.gif"></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <form name="search" method="get" action="search.asp">
            <td align="left" valign="top" class="bgTable"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif"></td>
                  <td align="left" class="bgTable" valign="top"> <table width="100%" border="0" cellspacing="2" cellpadding="2">
                      <tr align="left" valign="middle"> 
                        <td class="textBold">Keywords:</td>
                        <td><input name="keyword" type="text" class="form" id="keyword" size="14" maxlength="100"></td>
                      </tr>
                      <tr align="left" valign="middle"> 
                        <td class="textBold">Search In: </td>
                        <td><select name="iChannel" class="form" id="iChannel">
                            <option value="All" selected>Entire Site</option>
                            <%
While (NOT rsDropDownsearch.EOF)
%>
                            <option value="<%=(rsDropDownsearch.Fields.Item("CHA_ID").Value)%>"><%=(rsDropDownsearch.Fields.Item("CHA_MENU").Value)%></option>
                            <%
  rsDropDownsearch.MoveNext()
Wend
If (rsDropDownsearch.CursorType > 0) Then
  rsDropDownsearch.MoveFirst
Else
  rsDropDownsearch.Requery
End If
%>
                          </select></td>
                      </tr>
                      <tr align="left" valign="middle"> 
                        <td>&nbsp;</td>
                        <td><input name="Submit" type="submit" class="button" onClick="MM_validateForm('keyword','','R');return document.MM_returnValue" value="Search"></td>
                      </tr>
                    </table></td>
                  <td width="1" class="bgTableBorder"><img src="../assets/_spacer.gif"></td>
                </tr>
              </table></td>
          </form>
        </tr>
        <tr> 
          <td align="center" valign="top" background="../assets/bg_header_bottom.gif"><table border="0" cellpadding="0" cellspacing="0" class="bgTable" >
              <tr> 
                <td><img src="../assets/header_bottom.gif"></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="5"><img src="../assets/_spacer.gif" width="1" height="1"></td>
  </tr>
</table>
<%
rsDropDownsearch.Close()
Set rsDropDownsearch = Nothing
%>
