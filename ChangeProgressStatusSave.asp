<%@ Language=VBScript%>
<!--#include file="connect.inc" -->
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"> 

<head>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
<script type="text/javascript">
	function refreshForm()
	{
		opener.location.reload();
		window.close();
	}
</script>
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">CellPhone Billing</TD>
   </TR>
<!--
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
-->
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<%	
	EmpID_ = Request.Form("txtEmpID")
'	response.write CellPhone_ & "<br>"
	Mobilephone_ = Request.Form("txtCellphone")
'	response.write MonthP_ & "<br>"
	MonthP_ = Request.Form("txtMonthP")
'	response.write MonthP_ & "<br>"
	YearP_ = Request.Form("txtYearP")
'	response.write YearP_ & "<br>"
	Status_ = Request.Form("cmbStatus")
'	response.write YearP_ & "<br>"

	'Save Detail
	strsql = "Update MonthlyBilling Set ProgressId=" & Status_ & " Where EmpID='" & EmpID_ & "' And Phonenumber='" & Mobilephone_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ & "'"
	'response.write strsql & "<Br>"  
	BillingCon.execute(strsql) 

%>
<table cellspadding="1" cellspacing="0" width="60%" bgColor="white">
<tr>
	<td><br></td>
</tr>
<tr>
	<td>The data has been updated.</td>
</tr>
<tr>
	<td>
		&nbsp;&nbsp;<input type="button" value="Close" onClick="refreshForm();" </input>
	</td>
</tr>
</table>
</body> 

</html>
