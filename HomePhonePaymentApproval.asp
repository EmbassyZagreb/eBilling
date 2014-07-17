<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>
<style type="text/css">

</style>
<script language="JavaScript" src="calendar.js"></script>
<script type="text/javascript">
function validate_form()
{
	valid = true;
	msg = ""

	if (document.frmPaymentApproval.txtPaidAmount.value == "" )
	{
		msg = msg + "Please fill in your Paid Amount !!!\n"
		valid = false;
	}
	else
	{
		var myRegExp = new RegExp("^[/+|/-]?[0-9]*[/.]?[0-9]*$");
		if (myRegExp.test(document.frmPaymentApproval.txtPaidAmount.value) == false)
		{
			msg = msg + "Invalid data type for Paid Amount !!!\n"
			valid = false;
		}
	}


	if (valid == false)
	{
		alert(msg)
	}
	return valid;
}

function confirmSubmit()
{
var agree=confirm("Are you sure you wish to continue?");
if (agree)
	return true ;
else
	return false ;
}
</script>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">PAYMENT OF HOME PHONE</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<% 
 Homephone_ = "+" & trim(request("Homephone"))
' response.write Homephone_
 PageIndex_ = request("PageIndex")
' response.write PageIndex_

 dim user_ 
 dim user1_  
 dim rst 
 dim strsql
 
 user_ = request.servervariables("remote_user") 
  user1_ = right(user_,len(user_)-4)
user1_ = "pranataw"
'response.write user1_ & "<br>"

MonthP = request("MonthP")
YearP = request("YearP")
%>  

<form method="post" name="frmPaymentApproval" action="HomePhonePaymentApprovalSave.asp" onSubmit="return validate_form()">
<table align="center" cellspadding="1" cellspacing="0" width="100%">  
<tr>
	<td colspan="2" align="center">Billing Period : <Label style="color:blue"><%=MonthP %> - <%=YearP%></lable></td>
</tr>
</table>
<%
strsql = "Exec spGetPaymentList '2','" & MonthP & "','" & YearP & "','" & Homephone_ & "','','','',0,'X'" 
'response.write strsql & "<br>"
set rsData = server.createobject("adodb.recordset") 
set rsData = BillingCon.execute(strsql) 
if not rsData.eof then
	EmpName_ = rsData("EmpName")
	Period_ = rsData("MonthP") & " - " & rsData("YearP")
	Office_ = rsData("OfficeLocation")
'	HomePhone_ = rsData("Nomor")
	Status_ = rsData("Status")
	TotalCost_ = rsData("HomePhonePrsBillRp")
	ReceiptNo_ = rsData("ReceiptNo")
	PaidDate_ = rsData("PaidDate")
	if PaidDate_ ="" then
		PaidDate_ = date()
	end if
	CashierRemark_ = rsData("CashierRemark")
	TotalDollar_ = rsData("HomePhonePrsBillDlr")
	ExchangeRate_ = rsData("ExchangeRate")
	PaidCurrency_ = rsData("PaidCurrency")
end if
%>
<table cellspadding="1" cellspacing="0" width="65%">  
<tr>
          <td align="Left" colspan="6"><u><b>Personal Info<b></u></TD>
</tr>  
<tr>
	<td width="16%">Employee Name</td>
	<td width="1%">:</td>
	<td class="FontContent" colspan="4"><%=EmpName_%></td>
<!--
	<td width="20%">Billing Period (Month - Year)</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=Period_%></td>
-->
</tr>
<tr>
	<td>Agency / Office</td>
	<td width="1%">:</td>
	<td class="FontContent" colspan="4"><%=Office_%></td>
<!--
	<td>Total Cost</td>
	<td width="1%">:</td>
	<td class="FontContent">Kn. <%= formatnumber(TotalCost_ ,-1) %></td>
-->
</tr>
<tr>
	<td>HomePhone</td>
	<td width="1%">:</td>
	<td class="FontContent"><%= Homephone_ %></td>
	<td>Status </td>
	<td width="1%">:</td>
	<td class="FontContent"><%= Status_ %></td>
</tr>
<tr>
	<td colspan="6"><hr></td>
</tr>
<tr>
	<td align="Left" colspan="6"><u><b>Payment Info<b></u></TD>
</tr>
<tr>
	<td>Total Cost</td>
	<td width="1%">:</td>
	<td colspan="4">Kn. <a class="FontContent"><%= formatnumber(TotalCost_ ,-1) %></a>  &nbsp;|&nbsp; $. <a class="FontContent"><%= formatnumber(TotalDollar_ ,-1) %></a> &nbsp; <a class="FontHint" >(ExchangeRate : $1 / Kn <%= formatnumber(ExchangeRate_, -1) %> )</a> </td>
</tr>
<tr>
	<td>Paid amount</td>
	<td width="1%">:</td>
	<td class="FontContent" colspan="4">
		<select name="CurrencyList">
			<option value="Kn" <%if PaidCurrency_ = "Rp" then%>selected<%end if%> >Kuna
			<option value="Us" <%if PaidCurrency_ = "Us" then%>selected<%end if%> >US Dollar
		</select>
		<input type="input" name="txtPaidAmount" size="10" value='<%=TotalCost_%>' /> 
	</td>		
</tr>
<tr>
	<td>Receipt No.</td>
	<td width="1%">:</td>
	<td class="FontContent" colspan="4"><input type="input" name="txtReceiptNo" size="30" value='<%=ReceiptNo_%>' /> </td>		
</tr>
<tr>
	<td>Payment Date</td>
	<td width="1%">:</td>
	<td><input name="txtPaymentDate" type="Input" size="10" value='<%=PaidDate_ %>' maxlength="10" colspan="4">
	    <a href="javascript:cal0.popup();"><img src="images/calendar.gif" width="34" height="18" border="0" alt="Calendar"></a>
	</td>
</tr>
<tr>
	<td valign="top">Remark</td>
	<td valign="top" width="1%">:</td>
	<td colspan="4">
		<TextArea name="txtCashierRemark" Rows="5" Cols="60" Wrap><%=CashierRemark_%></textarea>
	</td>
</tr>
<tr>
	<td align="center" colspan="6"><br>
	</td>
</tr>
<tr>
	<td align="center" colspan="6">
        	<input type="submit" value="Submit" onClick="return confirmSubmit()">
       		&nbsp;<input type="button" value="Cancel" onClick="javascript:location.href='HomePhonePaymentList.asp?PageIndex=<%=PageIndex_%>&MonthP=<%=MonthP%>&YearP=<%=YearP%>'">
		<input type="hidden" name="txtHomePhone" value='<%=Homephone_%>' />
		<input type="hidden" name="txtMonthP" value='<%=MonthP%>' />
		<input type="hidden" name="txtYearP" value='<%=YearP%>' />
   	</td>
</tr>
		<script language="JavaScript">
	    	    var cal0 = new calendar1(document.forms['frmPaymentApproval'].elements['txtPaymentDate']);
			cal0.year_scroll = true;
			cal0.time_comp = false;
		</script>
</table>
</form>
</BODY>
</html>