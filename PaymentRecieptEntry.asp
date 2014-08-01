<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>

<script language="JavaScript" src="calendar.js"></script>
<script type="text/javascript">
function validate_form()
{
	valid = true;
	msg = ""

	if (document.frmPaymentApproval.cmbPaymentType.value == "" )
	{
		msg = msg + "Please select payment type !!!\n"
		valid = false;
	}

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

function CurrencyOnChange(obj)
{
	if (obj.selectedIndex == 0)
	{
		document.frmPaymentApproval.txtPaidAmount.value=document.frmPaymentApproval.elements['txtPaidAmountRp'].value;
	}
	else
	{
		document.frmPaymentApproval.txtPaidAmount.value=document.frmPaymentApproval.elements['txtPaidAmountDlr'].value;
	}
}
</script>



<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
   <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">ENTRY PAYMENT RECEIPT</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<% 
 PageIndex_ = request("PageIndex")
' response.write PageIndex_

 dim user_ 
 dim user1_  
 dim rst 
 dim strsql
 
 user_ = request.servervariables("remote_user") 
  user1_ = user_  'user1_ = right(user_,len(user_)-4)
'user1_ = "pranataw"
'response.write user1_ & "<br>"

EmpID = request("EmpID")
MonthP = request("MonthP")
YearP = request("YearP")
AlternateEmailFlag=Request("AlternateEmailFlag")

srStatus_ = Trim(request("srStatus"))
srsMonthP_ = Trim(request("srsMonthP"))
srsYearP_ = Trim(request("srsYearP"))
sreMonthP_ = Trim(request("sreMonthP"))
sreYearP_ = Trim(request("sreYearP"))
srEmpName_ = Trim(request("srEmpName"))
srOfficeSection_ = Trim(request("srOffice"))

strsql = "select RoleID from Users where loginId ='" & user1_ & "'"
set UserRS = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set UserRS = BillingCon.execute(strsql)
if not UserRS.eof then
	UserRole_ = UserRS("RoleID")
Else
	UserRole_ = ""
end if

If (UserRole_ = "Admin") or (UserRole_ = "Voucher") or (UserRole_ = "FMC") or (UserRole_ = "Cashier")  Then
%>  

<form method="post" name="frmPaymentApproval" action="PaymentRecieptEntrySave.asp" onSubmit="return validate_form();">
<%
HomePhoneBillRp_ = 0
HomePhoneBillDlr_ = 0
HomePhonePrsBillRp_ = 0
HomePhonePrsBillDlr_ = 0
OfficePhonePrsBillRp_ = 0
OfficePhonePrsBillDlr_ = 0
OfficePhoneBillRp_ = 0
OfficePhoneBillDlr_ = 0
CellPhoneBillRp_ = 0
CellPhoneBillDlr_ = 0
CellPhonePrsBillRp_ = 0
CellPhonePrsBillDlr_ = 0
TotalShuttleBillRp_ = 0
TotalShuttleBillDlr_ = 0
TotalBillingAmountPrsRp_ =0
TotalBillingAmountPrsDlr_ =0
TotalBillingRp_ = 0
TotalBillingDlr_ = 0

strsql = "Select * from vwMonthlyBilling Where EmpID='" & EmpID & "' and MonthP='" & MonthP & "' and YearP='" & YearP & "'"
'response.write strsql & "<br>"
set rsData = server.createobject("adodb.recordset") 
set rsData = BillingCon.execute(strsql) 
if not rsData.eof then
	EmpName_ = rsData("EmpName")
	Period_ = rsData("MonthP") & " - " & rsData("YearP")
	Office_ = rsData("Office")
	OfficePhone_ = rsData("WorkPhone")
	MobilePhone_ = rsData("MobilePhone")
	HomePhone_ = rsData("HomePhone")
	ExchangeRate_ = rsData("ExchangeRate")
	HomePhoneBillRp_ = rsData("HomePhoneBillRp")
	HomePhoneBillDlr_ = rsData("HomePhoneBillDlr")
	HomePhonePrsBillRp_ = rsData("HomePhonePrsBillRp")
	HomePhonePrsBillDlr_ = rsData("HomePhonePrsBillDlr")
	OfficePhonePrsBillRp_ = rsData("OfficePhonePrsBillRp")
	OfficePhonePrsBillDlr_ = rsData("OfficePhonePrsBillDlr")
	OfficePhoneBillRp_ = rsData("OfficePhoneBillRp")
	OfficePhoneBillDlr_ = rsData("OfficePhoneBillDlr")
	CellPhoneBillRp_ = rsData("CellPhoneBillRp")
	CellPhoneBillDlr_ = rsData("CellPhoneBillDlr")
	CellPhonePrsBillRp_ = rsData("CellPhonePrsBillRp")
	CellPhonePrsBillDlr_ = rsData("CellPhonePrsBillDlr")
	TotalShuttleBillRp_ = rsData("TotalShuttleBillRp")
	TotalShuttleBillDlr_ = rsData("TotalShuttleBillDlr")
	TotalBillingRp_ = rsData("TotalBillingRp")
	TotalBillingDlr_ = rsData("TotalBillingDlr")
'	response.write CellPhonePrsBillDlr_ 
	Status_ = rsData("ProgressDesc")
	PaidAmountDlr_ = rsData("PaidAmountDlr")
	PaidAmountRp_ = rsData("PaidAmountRp")
'	response.write PaidAmount_
	TotalBillingAmountPrsRp_ = rsData("TotalBillingAmountPrsRp")
	TotalBillingAmountPrsDlr_ = rsData("TotalBillingAmountPrsDlr")
	RemainingBillDlr_ = cdbl(TotalBillingAmountPrsDlr_) - cdbl(PaidAmountDlr_ )
	RemainingBillRp_ = cdbl(TotalBillingAmountPrsRp_ ) - cdbl(PaidAmountRp_ )
	FiscalStripNonVAT_ = rsData("FiscalStripNonVAT")
'	response.write RemainingBill_

end if
%>
<table cellspadding="1" cellspacing="0" width="65%"> 













<tr>
          <td colspan="6" align="Left"><u><b>Personal Info :<b></u></TD>
</tr>  
<!-- <tr>
	<td width="20%">Employee Name</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=EmpName_%></td>
	<td>Agency / Office</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=Office_%></td>
</tr> -->
<tr>
	<td colspan="6" align="Left">
	<table cellspadding="1" border="2" bordercolor="black" cellspacing="3" width="100%" bgColor="#999999" border="0">  

		<tr BGCOLOR="#999999">
			<td colspan="3" style="border: none;"><FONT color=#FFFFFF><b>Employee Name : <%=EmpName_%></b></font></td>
			<td colspan="3" style="border: none;" align="right"><FONT color=#FFFFFF><b>Phone Number : <%=MobilePhone_ %>&nbsp;</b></font></td>
		</tr>
		<tr BGCOLOR="#999999">
			<td colspan="6" style="border: none;"><FONT color=#FFFFFF><b>Position : <%=Position_%></b></font></td>
		</tr>
		<tr BGCOLOR="#999999">
			<td colspan="6" style="border: none;"><FONT color=#FFFFFF><b>Agency / Office : <%=Office_%></b></font></td>
		</tr>
		<tr BGCOLOR="#999999">
			<td colspan="6" style="border: none;"><FONT color=#FFFFFF><b>Fiscal Strip : <%=FiscalStripNonVAT_%></b></font></td>
		</tr>
	</table>
	</td>
</tr>
<!-- <tr>
	<td>Position</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=Position_ %></td>
	<td>Office Phone/Ext.</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=OfficePhone_ %></td>
</tr>
<tr>
	<td>Homephone</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=HomePhone_ %></td>
</tr>
<tr>
	<td>Mobile Phone</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=MobilePhone_ %></td>
	<td>Exchange Rate</td>
	<td width="1%">:</td>
	<td class="FontContent">Kn. <%= FormatNumber(ExchangeRate_,-1) %> / Dollar</td>

</tr>
<tr>
	<td>Mobile Phone</td>
	<td width="1%">:</td>
	<td class="FontContent" colspan="4"><%=MobilePhone_%></td>
</tr>
<tr>
	<td colspan="6"><hr></td>
</tr> -->

<tr>
	<td align="Left" colspan="5"><u><b>Billing detail :<b></u></TD>
</tr>
<tr>
	<td colspan="6">*Click on the bill for more detail</td>
</tr>
<tr>
	<td align="Left" colspan="6">
	<table cellspadding="1" border="1" bordercolor="black" cellspacing="0" width="100%" bgColor="white" border="0">  
	<tr align="center" height=26>
		<td width=20%><b>Action</b></td>
		<td width=20%><b>Billing Period</b></td>
		<td width=20%><b>Status</b></td>
		<td width=20%><b>Billing (Kn.)</b></td>
		<td width=20%><b>Should be paid (Kn.)</b></td>
	</tr>
<!-- <%if cdbl(OfficePhoneBillRp_) > 0 Then %>
	<tr>
		<td><a href="OfficePhoneDetail.asp?Extension=<%=OfficePhone_ %>&MonthP=<%=MonthP%>&YearP=<%=YearP%>" target="_blank">Office Phone</a></td>
		<td width="20%" class="FontContent" align="right"><%=formatnumber(OfficePhoneBillRp_,-1) %>&nbsp;</td>
		<td width="20%" class="FontContent" align="right"><%=formatnumber(OfficePhonePrsBillRp_ ,-1) %>&nbsp;</td>
		<td width="20%" class="FontContent" align="right"><%=formatnumber(OfficePhonePrsBillDlr_,-1) %>&nbsp;</td>		
	</tr>
<%else%>
	<tr>
		<td>Office Phone</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
	</tr>
<%end if%> -->
<!-- <%if cdbl(HomePhoneBillRp_) > 0 Then %>
	<tr>
		<td><a href="HomePhoneDetail.asp?HomePhone=<%=HomePhone_%>&MonthP=<%=MonthP%>&YearP=<%=YearP%>" target="_blank">Home Phone</a></td>
		<td class="FontContent" align="right"><%=formatnumber(HomePhoneBillRp_ ,-1) %>&nbsp;</td>
		<td class="FontContent" align="right"><%=formatnumber(HomePhonePrsBillRp_ ,-1) %>&nbsp;</td>
		<td class="FontContent" align="right"><%=formatnumber(HomePhonePrsBillDlr_ ,-1) %>&nbsp;</td>
	</tr>
<%else%>
	<tr>
		<td>Home Phone</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
	</tr>
<%end if%> -->
<%if cdbl(CellPhoneBillRp_ ) > 0 Then %>
	<tr height=26>
		<td>&nbsp;<a href="CellPhoneDetail.asp?CellPhone=<%=MobilePhone_ %>&MonthP=<%=MonthP%>&YearP=<%=YearP%>" target="_blank">View Submitted Bill</a></td>
	        <TD align="right">&nbsp;<%= MonthP%>-<%= YearP%></font>&nbsp;</TD>
	        <TD align="right"><%=Status_%>&nbsp;</font></TD>
		<td align="right"><%=formatnumber(CellPhoneBillRp_  ,-1) %>&nbsp;</td>
		<td align="right"><%=formatnumber(CellPhonePrsBillRp_ ,-1) %>&nbsp;</td>






	</tr>
<%else%>
	<tr>
		<td>Mobile Phone</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
	</tr>
<%end if%>
<!-- <%if cdbl(TotalShuttleBillRp_) > 0 Then %>
	<tr>
		<td><a href="ShuttleBusBillDetail.asp?Username=<%=user1_ %>&MonthP=<%=MonthP%>&YearP=<%=YearP%>" target="_blank">Shuttle Bus</a></td>
		<td class="FontContent" align="right"><%=formatnumber(TotalShuttleBillRp_ ,-1) %>&nbsp;</td>
		<td class="FontContent" align="right"><%=formatnumber(TotalShuttleBillRp_ ,-1) %>&nbsp;</td>
		<td class="FontContent" align="right"><%=formatnumber(TotalShuttleBillDlr_,-1) %>&nbsp;</td>
	</tr>
<%else%>
	<tr>
		<td>Shuttle Bus</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
	</tr>
<%end if%> -->
	</table>
	</TD>
</tr>

<tr>
	<td align="Left" colspan="6"><u><b>Payment Info:<b></u></TD>
</tr>
<!--
<tr>
	<td>Total Bill</td>
	<td width="1%">:</td>
	<td class="FontContent">Kn. <%= formatnumber(TotalBillingAmountPrsRp_ ,-1) %></td>
	<td>Already Paid</td>
	<td width="1%">:</td>
	<td class="FontContent">Kn. <%= formatnumber(PaidAmountRp_,-1) %></td>
</tr>
<tr>
	<td>Should be paid</td>
	<td width="1%">:</td>
	<td class="FontContent" colspan="4">Kn. <%= formatnumber(RemainingBill_ ,-1) %></td>
</tr>
<tr>
	<td>Paid amount</td>
	<td width="1%">:</td>
	<td class="FontContent" colspan="4">Kn. <input type="input" align="right" name="txtPaidAmount" size="10" value='<%=RemainingBill_ %>' /> </td>		
</tr>
-->
<tr>
	<td>Total Bill</td>
	<td width="1%">:</td>
	<td class="FontContent">Kn. <%= formatnumber(TotalBillingAmountPrsRp_ ,-1) %></td>
<!--	<td class="FontContent">Kn. <%= formatnumber(TotalBillingAmountPrsRp_ ,-1) %> ($. <%= formatnumber(TotalBillingAmountPrsDlr_,-1) %>)</td> -->
	<td>Already Paid</td>
	<td width="1%">:</td>
	<td class="FontContent">Kn. <%= formatnumber(PaidAmountRp_ ,-1) %></td>
<!--	<td class="FontContent">Kn. <%= formatnumber(PaidAmountRp_ ,-1) %> ($. <%= formatnumber(PaidAmountDlr_,-1) %>)</td> -->
</tr>
<tr>
	<td>Should be paid</td>
	<td width="1%">:</td>
	<td class="FontContent" colspan="4">Kn. <%= formatnumber(RemainingBillRp_ ,-1) %></td>
<!--	<td class="FontContent" colspan="4">Kn. <%= formatnumber(RemainingBillRp_ ,-1) %> ($. <%= formatnumber(RemainingBillDlr_ ,-1) %>)</td> -->
</tr>
<tr>
	<td>Paid amount</td>
	<td width="1%">:</td>
	<td colspan="4">
	<input type="hidden" name="txtPaidAmountRp" value='<%= RemainingBillRp_ %>' />
	<input type="hidden" name="txtPaidAmountDlr" value='<%= formatnumber(RemainingBillDlr_ ,-1) %>' />
<!--	<Select name="cmbCurrencyType" onChange="CurrencyOnChange(this);">
			<option value="Kn">Kuna</option>
			<option value="Dlr">Dollar</option>
		</select>&nbsp; --> <input type="input" align="right" name="txtPaidAmount" size="10" value='<%= formatnumber(RemainingBillRp_ ,-1) %>' /> * if the paid amount is less then suggested, the billing status will remain 'Paid Partial'</td>
	
</tr>
<!--	<tr>
	<td>Payment Type</td>
	<td width="1%">:</td>
	<td>
		<Select name="cmbPaymentType">
		<option value="">--Select--</option>
			<option value="F">Full Payment</option>
			<option value="P">Partial Payment</option>
		</select>
	</td>
</tr> -->
<tr>
	<td>Receipt No.</td>
	<td width="1%">:</td>
	<td class="FontContent" colspan="4"><input type="input" name="txtReceiptNo" size="30" value='<%=ReceiptNo_%>' /> </td>		
</tr>
<tr>
	<td>Payment Date</td>
	<td width="1%">:</td>
	<td><input name="txtPaymentDate" type="Input" size="10" value='<%=date()%>' maxlength="10" colspan="4">
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
	<td align="center" colspan="3">
        	<input type="submit" value="Save">
       		&nbsp;<input type="button" value="Cancel" onClick="javascript:location.href='PaymentReceiptList.asp?PageIndex=<%=PageIndex_%>&MonthP=<%=MonthP%>&YearP=<%=YearP%>&sMonthP=<%=srsMonthP_%>&sYearP=<%=srsYearP_%>&eMonthP=<%=sreMonthP_%>&eYearP=<%=sreYearP_%>&EmpName=<%=srEmpName_%>&Status=<%=srStatus_%>&OfficeSection=<%=srOfficeSection_%>'">
					
		<input type="hidden" name="txtEmpID" value='<%=EmpID%>' />
		<input type="hidden" name="txtMonthP" value='<%=MonthP%>' />
		<input type="hidden" name="txtYearP" value='<%=YearP%>' />
		<input type="hidden" name="txtExchangeRate" value='<%=ExchangeRate_%>' />
		<input type="hidden" name="txtRemainingBillRp" value='<%= formatnumber(RemainingBillRp_ ,-1) %>' />
		<input type="hidden" name="txtsrStatus" value='<%=srStatus_%>' />
		<input type="hidden" name="txtsrsMonthP" value='<%=srsMonthP_%>' />
		<input type="hidden" name="txtsrsYearP" value='<%=srsYearP_%>' />
		<input type="hidden" name="txtsreMonthP" value='<%=sreMonthP_%>' />
		<input type="hidden" name="txtsreYearP" value='<%=sreYearP_%>' />
		<input type="hidden" name="txtsrEmpName" value='<%=srEmpName_%>' />
		<input type="hidden" name="txtsrOffice" value='<%=srOfficeSection_%>' />
		<input type="hidden" name="txtsrPageIndex" value='<%=PageIndex_%>' />

   	</td>
</tr>
		<script language="JavaScript">
	    	    var cal0 = new calendar1(document.forms['frmPaymentApproval'].elements['txtPaymentDate']);
			cal0.year_scroll = true;
			cal0.time_comp = false;
		</script>
</table>
</form>
<%else %>
	<table>
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
		</tr>
	</table>

<%
end if 
%>
</BODY>
</html>