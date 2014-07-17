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
PaymentDueDate_ = request.form("cmbDueDate")
Ceiling_ = request.form("txtCeilingAmount")
DetailRecordAmount_ = request.form("txtDetailRecordAmount")
CashierMinimumAmount_ = request.form("txtCashierMinimumAmount")
%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">

<meta http-equiv="refresh" content="1;url=SetupPaymentDueDate.asp">

<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
   <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">PAYMENT DUE DATE UPDATE</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>

<table border=0 width=100%>
<%
       strsql = "Update PaymentDueDate Set PaymentDueDate=" & PaymentDueDate_ & ", CeilingAmount =" & Ceiling_ & ", DetailRecordAmount=" & DetailRecordAmount_ & ", CashierMinimumAmount=" & CashierMinimumAmount_
       'response.write strsql 
       BillingCon.execute strsql
%>               
<tr><td><br></td></tr>
<tr><td align=center>Your data has already been saved. Thank you.</td></tr>
<tr><td>&nbsp;</td>
<!--
<tr><td align=center> 
<input type="button" value="Close" id="btnclose">
</td></tr>
-->
<tr>
	<td align="center"><br><a href="SetupPaymentDueDate.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
</tr>
</table>

   </body>
</html>