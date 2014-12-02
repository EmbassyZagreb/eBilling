<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>
<TITLE>U.S. Embassy Zagreb - zBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Home</TD>
   </TR>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<form method="post">
<%
 dim user_
 dim user1_
 dim rst
 dim strsql
 'response.end
 user_ = request.servervariables("remote_user")
'response.write user_ & "<br>"
user1_ = user_  'user1_ = right(user_,len(user_)-4)

'if (session("Month") = "") or (session("Year") = "") then
'	strsql = "Select MonthP, YearP From Period"
'	'response.write strsql & "<br>"
'	set rsData = server.createobject("adodb.recordset")
'	set rsData = BillingCon.execute(strsql)
'	if not rsData.eof then
'		session("Month") = rsData("MonthP")
'		session("Year") = rsData("YearP")
'	end if
'end if

strsql = "select RoleID from Users where loginId ='" & user1_ & "'"
set UserRS = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set UserRS = BillingCon.execute(strsql)
if not UserRS.eof then
	UserRole_ = UserRS("RoleID")
end if
%>

<table cellspadding="1" cellspacing="1" width="65%">
<tr>
	<td align="left" valign="top">
	   <UL><UL TYPE="square">
		<LI class="normal"><strong>Employee</strong>
			<UL TYPE="">
<!--
				<li><A HREF="HomePhonzBilling.asp"><strong>Home Phone</strong></A></li>
				<li><A HREF="OfficePhonzBilling.asp"><strong>Office Phone</strong></A></li>
-->
				<li><A HREF="MonthlyBilling.asp"><strong>Billing</strong></A></li>
				<li><A HREF="MonthlyBillingTracking.asp"><strong>Status</strong></A></li>
<!--				<li><A HREF="MonthlyBillList.asp"><strong>Print Monthly Billing</strong></A></li> -->
<!--				<li><A HREF="MonthlyBillList.asp"><strong>Monthly Billing Report</strong></A></li> -->
				<li><A HREF="UserSetting.asp"><strong>Supervisor Setting</strong></A></li>
<!--
				<li><A HREF="PersonalPhoneList.asp"><strong>Personal Phone List</strong></A></li>
				<li><A HREF="ePaymentForm.asp"><strong>Pay.gov form</strong></A></li>
-->
			</ul>
  		  <LI class="normal"><strong>Supervisor Functions</strong></li>
			<UL TYPE="">
<!--		<li><A HREF="BillingApprovalList.asp"><strong>Office Phone</strong></A></li> -->
				<li><A HREF="BillingApprovalList.asp"><strong>Subordinate Bills</strong></A></li>
			</ul>

	<%if (UserRole_= "IM") or (UserRole_= "Admin") then %>
		<br><LI class="normal"><strong>IRM Functions and Reports:
			<UL TYPE="">
<!--					<LI>&nbsp;<A HREF="LongDistanceCallsReport.asp"><strong>Long Distance Calls Report</strong></A>
			<LI>&nbsp;<A HREF="CellphoneCallsReport.asp"><strong>Cell phone Calls Report</strong></A> -->
<!--				<LI>&nbsp;<A HREF="LineUsagesReport.asp"><strong>Line Usages Report</strong></A> -->
<!--				<LI>&nbsp;<A HREF="NoLineUsagesReport.asp"><strong>No Line Usages Report</strong></A> -->
          <LI class="highlightlistitem"><A HREF="GenerateMonthlyBill.asp"><strong>Generate Monthly Billing (use sparingly)</strong></A></li>
          <LI class="normal"><strong>Admins:</strong></li>
          <LI><A HREF="UserList.asp"><strong>Manage Power Users</strong></A></li>
      </ul>
            <!--				<LI class="normal"><A HREF="AdminPage.asp"><strong>Admins Page(s)</strong></A>-->
	<%end if%>
<!--	<%if (UserRole_= "Trs") or (UserRole_= "Admin") or (UserRole_= "FMC") then %>
		<LI class="normal"><strong>Transportation :
			<UL TYPE="">
				<li><A HREF="ShuttleBusList.asp"><strong>Shuttle Bus Payment List</strong></A></li>
				<li><A HREF="ShuttleBusRateList.asp"><strong>Setup Shuttle Bus Rate</strong></A></li>
				<li><A HREF="InputShuttleBusUsage.asp"><strong>Input Shuttle Bus Usage</strong></A></li>
			</ul>
	<%end if%> -->
<!--		<LI class="normal"><A HREF="BillingReportList.asp"><strong>Reports</strong></A> -->
<!--
		<LI class="normal"><strong>Reports:
			<UL TYPE="square">
				<LI>&nbsp;<A HREF="HomePaymentBillingReport.asp"><strong>Home Payment Billing Report</strong></A>
				<LI>&nbsp;<A HREF="LongDistanceCallsReport.asp"><strong>Long Distance Calls Report</strong></A>
				<LI>&nbsp;<A HREF="CellphoneCallsReport.asp"><strong>Cell phone Calls Report</strong></A>
				<LI>&nbsp;<A HREF="LineUsagesReport.asp"><strong>Line Usages Report</strong></A>
				<LI>&nbsp;<A HREF="NoLineUsagesReport.asp"><strong>No Line Usages Report</strong></A>
				<LI>&nbsp;<A HREF="MonthlyBillingByExtReport.asp"><strong>Comparison Monthly Bill by Ext.</strong></A>
			</UL>
-->
<%if (UserRole_= "IM") or (UserRole_= "Admin") or (UserRole_= "FMC") then %>

<!--		<LI class="normal"><A HREF="PhoneList.asp"><strong>Phone/Extension Number List</strong></A>&nbsp;[&nbsp;<A HREF="OfficePhoneNumberList.asp"><strong>Office Phone</strong></A>&nbsp;|&nbsp;<A HREF="HomePhoneNumberList.asp"><strong>Home Phone</strong></A>&nbsp;|&nbsp;<A HREF="CellPhoneNumberList.asp"><strong>Cell Phone</strong></A>&nbsp;]&nbsp; -->
		<br><LI class="normal"><strong>User Management:</strong>
			<UL TYPE="">
				<li><A HREF="EmployeeList.asp"><strong>Employees</strong></A></li>
<!--				<li><A HREF="NonEmployeeList.asp"><strong>Non Employee List</strong></A></li> -->
				<li><A HREF="CellPhoneNumberList.asp"><strong>Cell Phones</strong></A></li>
			</ul>

<%end if%>

<%if (user1_= "PribanicM") then %>
		<LI class="normal">
			<strong>Others List :</strong>
				<UL TYPE="">
				       <li><A HREF="EmployeeList.asp"><strong>Employees</strong></A></li>
			</ul>
<%end if%>

<%if (UserRole_= "Admin") then %>

<%end if%>
			</ul>
		</li>
	   </UL></UL>
	</td>
	<td align="left" valign="top">
	   <UL>
		<UL TYPE="square">
<%if (UserRole_= "Admin") or (UserRole_= "FMC") or (UserRole_= "Voucher") then %>
		<LI class="normal"><strong>FMO Voucher Functions:
			<UL TYPE="">
<!--				<li><A HREF="HomePhoneSelfPaymentList.asp"><strong>Update Payment to Telkom</strong></A></li> -->
<!--				<li><A HREF="PaymentList.asp"><strong>Update Payment</strong></A></li>-->
				<li><A HREF="AgencyList.asp"><strong>Funding Agencies</strong></A></li>
<!--				<li><A HREF="ExchangeRateList.asp"><strong>Exchange Rate</strong></A></li> -->
				<li><A HREF="SetupPaymentDueDate.asp"><strong>Setup Payment Due Date & Ceiling Amount</strong></A></li>
	 		        <LI><A HREF="UpdateSupervisor.asp"><strong>Update supervisor </strong></A></li>
				<li><A HREF="SendNotification.asp"><strong>Send Notification</strong></A></li>
<!--				<li><A HREF="SendNotificationAll.asp"><strong>Send Notification - All </strong></A></li> -->
<!--				<li><A HREF="SendNotificationPayGov.asp"><strong>Send Notification - PayGov.com</strong></A></li> -->
			</ul>
		<br><LI class="normal"><strong>Reconciliation Reports :</strong>
			<UL TYPE="">
			 	<LI><A HREF="ReconciliationReport.asp"><strong>Reconciliation Report</strong></A></li>
			 	<LI><A HREF="SummaryofGenerateMonthlyBillsReport.asp"><strong>Summary of Generate Monthly Bills Result</strong></A></li>
			</ul>
		<br><LI class="normal"><strong>Reports:</strong>
			<UL TYPE="">
				<li><A HREF="BillingSettlement.asp"><strong>Billing Settlement</strong></A></li>
<!--				<li><A HREF="ARBillingReport.asp"><strong>A/R Billing Report(Personal)</strong></A></li>-->
				<li><A HREF="ARBillingReportAll.asp"><strong>A/R Billing Report(Complete)</strong></A></li>
				<li><A HREF="ARPaymentReport.asp"><strong>A/R Payment Report</strong></A></li>
				<li><A HREF="ARReminder.asp"><strong>A/R Reminder</strong></A></li>
<!--				<li><A HREF="ARAgingReport.asp"><strong>A/R Aging Report</strong></A></li> -->
				<li><A HREF="OutstandingReport.asp"><strong>Outstanding inquiry</strong></A></li>
				<li><A HREF="TopBillingReport.asp"><strong>Highest Bills</strong></A></li>
				<li><A HREF="UnknownCellphoneReport.asp"><strong>Unknown Cellphone Bill Report</strong></A></li>
				<li><A HREF="SupervisorReminder.asp"><strong>Supervisor Reminder Report</strong></A></li>
				<li><A HREF="FiscalDataReport.asp"><strong>Fiscal Data Report</strong></A></li>
			</ul>
		</li>
	<%end if%>
	<%if (UserRole_= "Cashier") or (UserRole_= "Admin") or (UserRole_= "FMC") then %>
		<br><LI class="normal"><strong>FMC Cashier :
			<UL TYPE="">
<!--				<li><A HREF="HomePhonePaymentList.asp"><strong>Home Phone</strong></A></li> -->
<!--				<li><A HREF="OfficePhonePaymentList.asp"><strong>Office Phone</strong></A></li> -->
				<li><A HREF="PaymentReceiptList.asp"><strong>Payment Receipt</strong></A></li>
<!--				<li><A HREF="MonthlyBillListAll.asp"><strong>Print Monthly Bill</strong></A></li> -->
<!--				<li><A HREF="MonthlyBillListAll.asp"><strong>Monthly Billing Report</strong></A></li>-->
<!--
				<li><strong>Reports :</strong>
					<UL TYPE="">    -->
						<li><A HREF="ARPaymentReport.asp"><strong>A/R Payment Report</strong></A></li>
<!--						<li><A HREF="ARReminder.asp"><strong>A/R Reminder</strong></A></li>
-->
<!--						<li><A HREF="ARAgingReport.asp"><strong>A/R Aging Report</strong></A></li>
						<li><A HREF="OutstandingReport.asp"><strong>Outstanding inquiry</strong></A></li>
					</ul>
				</li>

				<li><A HREF="ExchangeRateList.asp"><strong>Exchange Rate</strong></A></li>
-->
			</ul>
	<%end if%>
	   	</UL>

	    </UL>
	</td>
</tr>

</table>

</form>
</BODY>
</html>