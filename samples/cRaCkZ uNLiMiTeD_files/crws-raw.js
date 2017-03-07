if (document.cookie.indexOf("Funraw") == -1) {
  var expdate = new Date((new Date()).getTime() + 360);
  document.cookie="Funraw=POP; expires=" + expdate.toGMTString() + "; path=/;";


//   var w = window.screen.availWidth;
//   var h = window.screen.availHeight;
//   win=window.open("http://www.my-stats.com/ViewSponsorAdt.php?id=1478",'FuN','screenX=0,screenY=0,left=0,top=0,width=' + w + ',height=' + h +',resizable=1,scrollbars=1,resizable=1,status=0,menubar=0');
  
win=window.open("http://www.my-stats.com/ViewSponsorARWN.php?id=1478",'FuN','width=600,height=460,scrollbars=1,resizable=1,status=0,menubar=0');
  if (!win.opener) { win.opener=self; }
  window.focus();
}
