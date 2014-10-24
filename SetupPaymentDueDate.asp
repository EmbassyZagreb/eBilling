
<!--#include file="connect.inc" -->
<html>
<head>

<script language="JavaScript" src="calendar.js"></script>
<script type="text/javascript">
function validate_form()
{
	valid = true;
	msg = ""

	if (document.frmPaymentDueDate.txtCeilingAmount.value == "" )
	{
		msg = msg + "Please fill in Ceiling Amount !!!\n"
		valid = false;
	}
	else
	{
		var myRegExp = new RegExp("^[/+|/-]?[0-9]*[/.]?[0-9]*$");
		if (myRegExp.test(document.frmPaymentDueDate.txtCeilingAmount.value) == false)
		{
			msg = msg + "Invalid data type for Paid Amount !!!\n"
			valid = false;
		}
	}

	if (document.frmPaymentDueDate.txtDetailRecordAmount.value == "" )
	{
		msg = msg + "Please fill in hide detail record value !!!\n"
		valid = false;
	}
	else
	{
		var myRegExp = new RegExp("^[/+|/-]?[0-9]*[/.]?[0-9]*$");
		if (myRegExp.test(document.frmPaymentDueDate.txtDetailRecordAmount.value) == false)
		{
			msg = msg + "Invalid data type for hide detail record value !!!\n"
			valid = false;
		}
	}

	if (document.frmPaymentDueDate.txtCashierMinimumAmount.value == "" )
	{
		msg = msg + "Please fill in accumulated debt record value !!!\n"
		valid = false;
	}
	else
	{
		var myRegExp = new RegExp("^[/+|/-]?[0-9]*[/.]?[0-9]*$");
		if (myRegExp.test(document.frmPaymentDueDate.txtCashierMinimumAmount.value) == false)
		{
			msg = msg + "Invalid data type for accumulated debt record value !!!\n"
			valid = false;
		}
	}

	if (valid == false)
	{
		alert(msg)
	}
	return valid;
}
</script>
<script language="vbscript">
       <!--
        Sub btnCancel_onclick
           history.back
	End Sub

       --> 
   </script>
<% 
 dim user_ 
 dim user1_  

 
 user_ = request.servervariables("remote_user") 
 user1_ = user_  'user1_ = right(user_,len(user_)-4)
'user1_ = "pranataw"
'response.write user1_ & "<br>"

%> 

<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">SETUP PAYMENT DUE DATE</TD>
   </TR>
	<tr>
        	<td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
	</tr>
	<TR>
		<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
	</TR>
  </TABLE>
<form method="post" name="frmPaymentDueDate" id="frmPaymentDueDate" action="SetupPaymentDueDateSave.asp" onSubmit="return validate_form()">
<%  
 dim rst 
 dim strsql
 dim rst1
 
 strsql = "select * from Users where loginId='" & user1_ & "'"
'response.write user1_ & "<br>"
'response.write strsql 
  set rst = server.createobject("adodb.recordset") 
  set rst = BillingCon.execute(strsql)
  if not rst.eof then 
	if (trim(rst("RoleID")) = "Admin") or (trim(rst("RoleID")) = "Voucher") or (trim(rst("RoleID")) = "FMC") then 
	        strsql = " select * from PaymentDueDate"
       		set rst1 = server.createobject("adodb.recordset") 
		'response.write strsql 
        	set rst1 = BillingCon.execute(strsql)
	       	if not rst1.eof then 
		   PaymentDuedate_ = rst1("PaymentDueDate")
		   CeilingAmount_ = rst1("CeilingAmount")
		   DetailRecordAmount_ = rst1("DetailRecordAmount")
		   CashierMinimumAmount_ = rst1("CashierMinimumAmount")
        	end if
		'response.write PaymentDuedate_ & "<br>"
%>             
		<table align=center>  
		<tr>
		  <td align=right>Payment Due Date :</td>
		  <td>			
			<select name="cmbDueDate">		
<%			X = 1
			Do While X<=30 %>	
				<option value=<%=X%> <%If PaymentDuedate_ = X Then%> Selected<%End If%>><%=X%></option>
<%				X=X+1
			Loop
%>	
			</select>
		  </td>
		</tr>
		<tr>
		  <td align=right>Ceiling Amount for email notification [Kn] :</td>
		  <td>	
			<input type="input" id="txtCeilingAmount" name="txtCeilingAmount" size="8" value='<%=CeilingAmount_ %>'>
		  </td>
		</tr>
		<tr>
		  <td align=right>Hide detail record less than or equal to [Kn] :</td>
		  <td>	
			<input type="input" id="txtDetailRecordAmount" name="txtDetailRecordAmount" size="8" value='<%=DetailRecordAmount_ %>'>
		  </td>
		</tr>
		<tr>
		  <td align=right>Make the payment at cashier if accumulated debt is greater than [Kn] :</td>
		  <td>	
			<input type="input" id="txtCashierMinimumAmount" name="txtCashierMinimumAmount" size="8" value='<%=CashierMinimumAmount_ %>'>
		  </td>
		</tr>
		<tr>
			<td colspan="2"><br></td>
		</tr>
		<tr>
		  <td></td>
		  <td><input type="submit" id="btnSubmit" name="btnSubmit" value="Submit">
		      &nbsp;<input type="button" value="Cancel" name="btnCancel">
		 </td>
		</tr>  
		<tr>
		   <td colspan=2>&nbsp;</td>
		</tr>
		</table>
     <%  else %>
	<table>
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
		</tr>
	</table>
<%   end if 
else %>
	<table align="center">
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
		</tr>
	</table>

<% end if %>
</form>
</BODY>
</html>