<%@ Language=VBScript %>
<!--#include file="connect.inc" -->

<html>
   <head>
   <script language="vbscript">
       <!--
        Sub btnBack_onclick
           history.back
	End Sub
        Sub btnClose_onclick
		close
	End Sub
       --> 
   </script>
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 
<%
EmpID_ = Request.form("txtEmpID")
EmpType_ = Request.form("txtEmpType")
ReportTo_ = Request.form("cmbReportTo")

%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
<CENTER>
<table border=0 width=100%>
<%


if EmpType_ = "LES" then
       strsql = "Update Local Set ReportTo ='" & ReportTo_ & "' Where FN_EMP_KEY_NUM=" & EmpID_ 
else
       strsql = "Update Amer Set ReportTo ='" & ReportTo_ & "' Where AMER_EMP_KEY_NUM=" & EmpID_ 
end if
	'response.write strsql 
       JakEmpCon.execute strsql
%>               
<tr><td align=center>Your data has been updated.</td></tr>
<tr><td>&nbsp;</td>
<!--
<tr><td align=center> 
<input type="button" value="Close" id="btnclose">
</td></tr>
<tr>
	<td align="center"><br><a href="DefaultAddOn.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
</tr>
-->
</table>

   </body>
</html>