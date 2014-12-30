<%@ Language=VBScript %>
<!--#include file="connect.inc" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate"/>
<meta http-equiv="Pragma" content="no-cache"/>
<meta http-equiv="Expires" content="0"/>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>eBilling website down</title>

<script type="text/JavaScript">

function Redirect(t) {
	setTimeout("location.href = '<%=WebSiteAddress%>';",t);
}

</script>

<style type="text/css">
<!--
.style1 {
	font-size: 14px;
	font-family: Arial, Helvetica, sans-serif;
}
.style2 {
	font-family: Arial, Helvetica, sans-serif;
	color: #990000;
	font-size: 18px;
	font-weight: bold;
}
.style3 {color: #990000}
.style4 {font-size: 14px; font-family: Arial, Helvetica, sans-serif; color: #666666; }
.style5 {font-size: 11px; font-family: Arial, Helvetica, sans-serif; color: #666666; }
-->
</style>
</head>

<body onload="JavaScript:Redirect(15000);">
<p>&nbsp;</p>
<table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="60"><img src="maintenance/seal.png" width="60" height="60" alt="Embassy of United States, Zagreb, Croatia" /></td>
    <td width="340" valign="bottom"><img src="maintenance/zagreb-croatia.PNG" width="301" height="52" alt="Embassy of United States, Zagreb, Croatia" /></td>
    <td width="400">&nbsp;</td>
  </tr>
  <tr>
    <td height="24" colspan="3" background="maintenance/line.PNG">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="3">
		<br />
		<br />
	<p align="center" class="style2">
      eBilling website is currently down for maintenance.</p>
    <p align="center" class="style4">We expect to be back soon. Thanks for your patience.</p>
		<br />
		<br />
    <hr size="0" class="style4" />
	    <br />
    <p align="center" class="style4">For further information please call ISC at #3333.</p>
    </p></td>
  </tr>
  <tr>
    <td height="24" colspan="3" background="maintenance/line.PNG">&nbsp;</td>
  </tr>
    <tr>
    <td height="14" colspan="3" align="right" class="style5">Locked by: <%=filecontent%></td>
  </tr>
</table>
</body>
</html>
