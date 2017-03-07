<html>
<head>
<title>pick</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"><SCRIPT LANGUAGE="JavaScript">

<!--
function checkOS() {
  if(navigator.userAgent.indexOf('IRIX') != -1)
    { var OpSys = "Irix"; }
  else if((navigator.userAgent.indexOf('Win') != -1) &&
  (navigator.userAgent.indexOf('95') != -1))
    { var OpSys = "Windows95"; }
  else if(navigator.userAgent.indexOf('Win') != -1)
    { var OpSys = "Windows3.1 or NT"; }
  else if(navigator.userAgent.indexOf('Mac') != -1)
    { var OpSys = "Macintosh"; }
  else { var OpSys = "other"; }
  return OpSys;
}
// -->

</SCRIPT>
</head>

<body bgcolor="#00009f" text="#FFCC00" background="zonebk2.jpg"><center>
<table border="0">
  <tr> 
    <td><center>
        <p><b><font size="5">Your Screen Resolution</font></b></p>
        <p><font size="4">NOW IS</font></p>
      </center></td>
  </tr>
  
  <tr> 
    <td> 
      <form method="POST" name="t">
        <table border="0">
          <tr> 
            <td valign="top" width="150"
  bgcolor="#00009f"><strong>width:</strong></td>
            <td> 
              <input type="text" size="20"
name="t1" value="not supported">
            </td>
          </tr>
          <tr> 
            <td valign="top" width="150"
 bgcolor="#00009f"><strong>height:</strong></td>
            <td> 
              <input type="text" size="20"
name="t2" value="not supported">
            </td>
          </tr>
          <tr> 
            <td valign="top" width="150"
bgcolor="#00009f"><strong>colorDepth:</strong></td>
            <td> 
              <input type="text" size="20"
 name="t3" value="not supported">
            </td>
          </tr>
          <tr> 
            <td valign="top" width="150"
bgcolor="#00009f"><strong>pixelDepth: </strong></td>
            <td> 
              <input type="text" size="20"
 name="t4" value="not supported">
            </td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table>
<p align="left">
<script>
<!--
function show(){
if (!document.all&&!document.layers)
return
document.t.t1.value=screen.width
document.t.t2.value=screen.height
document.t.t3.value=screen.colorDepth
document.t.t4.value=screen.pixelDepth
}
show()
//-->
</script>

<center><script>



if (document.all)
var version=/MSIE \d+.\d+/

if (!document.all)
document.write("You are using "+navigator.appName+" "+navigator.userAgent)
else
document.write("You are using "+navigator.appName+" "+navigator.appVersion.match(version))

</script></center>

<center><script>
<!--
var OpSys = checkOS();
document.write(OpSys);
//-->
</script></center>





