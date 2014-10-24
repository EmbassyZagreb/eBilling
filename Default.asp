<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
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
		<LI class="normal"><B>Employee</b>
			<UL TYPE="">
<!--
				<li><A HREF="HomePhoneBilling.asp"><B>Home Phone</B></A></li>
				<li><A HREF="OfficePhoneBilling.asp"><B>Office Phone</B></A></li>
-->
				<li><A HREF="MonthlyBilling.asp"><B>Monthly Billing</B></A></li>
				<li><A HREF="MonthlyBillingTracking.asp"><B>Tracking Status</B></A></li>
<!--				<li><A HREF="MonthlyBillList.asp"><B>Print Monthly Billing</B></A></li> -->
<!--				<li><A HREF="MonthlyBillList.asp"><B>Monthly Billing Report</B></A></li> -->
				<li><A HREF="UserSetting.asp"><B>Supervisor Setting</B></A></li>
<!--
				<li><A HREF="PersonalPhoneList.asp"><B>Personal Phone List</B></A></li>
				<li><A HREF="ePaymentForm.asp"><B>Pay.gov form</B></A></li>
-->

			</ul>	
		<br><LI class="normal"><B>Supervisor</B>
			<UL TYPE="">
<!--				<li><A HREF="BillingApprovalList.asp"><B>Office Phone</B></A></li> -->
				<li><A HREF="BillingApprovalList.asp"><B>Approval List</B></A></li>
			</ul>	
	
	<%if (UserRole_= "IM") or (UserRole_= "Admin") then %>
		<br><LI class="normal"><B>IM/IPC reports :
			<UL TYPE="">
				<LI>&nbsp;<A HREF="LongDistanceCallsReport.asp"><b>Long Distance Calls Report</b></A>
<!--				<LI>&nbsp;<A HREF="CellphoneCallsReport.asp"><b>Cell phone Calls Report</b></A> -->
				<LI>&nbsp;<A HREF="LineUsagesReport.asp"><b>Line Usages Report</b></A>
<!--				<LI>&nbsp;<A HREF="NoLineUsagesReport.asp"><b>No Line Usages Report</b></A> -->
			</ul>
	<%end if%>
<!--	<%if (UserRole_= "Trs") or (UserRole_= "Admin") or (UserRole_= "FMC") then %>
		<LI class="normal"><B>Transportation :
			<UL TYPE="">
				<li><A HREF="ShuttleBusList.asp"><B>Shuttle Bus Payment List</B></A></li> 
				<li><A HREF="ShuttleBusRateList.asp"><B>Setup Shuttle Bus Rate</B></A></li> 
				<li><A HREF="InputShuttleBusUsage.asp"><B>Input Shuttle Bus Usage</B></A></li> 
			</ul>
	<%end if%> -->
<!--		<LI class="normal"><A HREF="BillingReportList.asp"><B>Reports</B></A> -->
<!--
		<LI class="normal"><B>Reports:
			<UL TYPE="square">
				<LI>&nbsp;<A HREF="HomePaymentBillingReport.asp"><b>Home Payment Billing Report</b></A>
				<LI>&nbsp;<A HREF="LongDistanceCallsReport.asp"><b>Long Distance Calls Report</b></A>
				<LI>&nbsp;<A HREF="CellphoneCallsReport.asp"><b>Cell phone Calls Report</b></A>
				<LI>&nbsp;<A HREF="LineUsagesReport.asp"><b>Line Usages Report</b></A>
				<LI>&nbsp;<A HREF="NoLineUsagesReport.asp"><b>No Line Usages Report</b></A>
				<LI>&nbsp;<A HREF="MonthlyBillingByExtReport.asp"><b>Comparison Monthly Bill by Ext.</b></A>
			</UL>
-->
<%if (UserRole_= "IM") or (UserRole_= "Admin") or (UserRole_= "FMC") then %>

<!--		<LI class="normal"><A HREF="PhoneList.asp"><B>Phone/Extension Number List</B></A>&nbsp;[&nbsp;<A HREF="OfficePhoneNumberList.asp"><B>Office Phone</B></A>&nbsp;|&nbsp;<A HREF="HomePhoneNumberList.asp"><B>Home Phone</B></A>&nbsp;|&nbsp;<A HREF="CellPhoneNumberList.asp"><B>Cell Phone</B></A>&nbsp;]&nbsp; -->
		<br><LI class="normal"><B>Others List :
			<UL TYPE="">
				<li><A HREF="EmployeeList.asp"><B> Employee List</B></A></li>
<!--				<li><A HREF="NonEmployeeList.asp"><B>Non Employee List</B></A></li> -->
				<li><A HREF="CellPhoneNumberList.asp"><B>Cell Phone</B></A></li>
			</ul>
		
<%end if%>

<%if (user1_= "PribanicM") then %>
		<LI class="normal">
			<b>Others List :</b>
				<UL TYPE="">
				       <li><A HREF="EmployeeList.asp"><B> Employee List</B></A></li>
			</ul>
<%end if%>

<%if (UserRole_= "Admin") then %>
		<br><LI class="normal">
			<b>Data Import:</b>
				<UL TYPE="">
						<li><A HREF="ImportBilling.asp"><B>Upload Billing Data</B></A></li>
				       <LI><A HREF="LanguageTranslation.asp"><B>Language Translation</B></A>
<!--				<LI class="normal"><A HREF="AdminPage.asp"><B>Admins Page(s)</B></A>-->
			</ul>
<%end if%>

<%if (UserRole_= "Admin") then %>
		<br><LI class="normal">
			<b>Admins:</b>
				<UL TYPE="">
				       <LI><A HREF="UserList.asp"><B>Manage User(s)</B></A>
<!--				<LI class="normal"><A HREF="AdminPage.asp"><B>Admins Page(s)</B></A>-->
<%end if%>
			</ul>
		</li>		
	   </UL></UL>
	</td>
	<td align="left" valign="top">
	   <UL>
		<UL TYPE="square">
<%if (UserRole_= "Admin") or (UserRole_= "FMC") or (UserRole_= "Voucher") then %>
		<LI class="normal"><b>FMC Voucher :
			<UL TYPE="">
<!--				<li><A HREF="HomePhoneSelfPaymentList.asp"><B>Update Payment to Telkom</B></A></li> -->
<!--				<li><A HREF="PaymentList.asp"><B>Update Payment</B></A></li>-->
				<li><A HREF="AgencyList.asp"><B>Funding Agency List</B></A></li>
<!--				<li><A HREF="ExchangeRateList.asp"><B>Exchange Rate</B></A></li> -->
				<li><A HREF="SetupPaymentDueDate.asp"><B>Setup Payment Due Date & Ceiling Amount</B></A></li>
	 		        <LI><A HREF="UpdateSupervisor.asp"><B>Auto update supervisor list</B></A></li>
				<li><A HREF="SendNotification.asp"><B>Send Notification</B></A></li>
<!--				<li><A HREF="SendNotificationAll.asp"><B>Send Notification - All </B></A></li> -->
<!--				<li><A HREF="SendNotificationPayGov.asp"><B>Send Notification - PayGov.com</B></A></li> -->
			</ul>
		<br><LI class="normal"><b>Reconciliation Reports :</B>
			<UL TYPE="">
				<LI><A HREF="GenerateMonthlyBill.asp"><B>Generate Monthly Billing</B></A>
			 	<LI><A HREF="ReconciliationReport.asp"><B>Reconciliation Report</B></A></li>
			 	<LI><A HREF="SummaryofGenerateMonthlyBillsReport.asp"><B>Summary of Generate Monthly Bills Result</B></A></li>
			</ul>	
		<br><LI class="normal"><b>Reports:</b>
			<UL TYPE="">
				<li><A HREF="BillingSettlement.asp"><B>Billing Settlement</B></A></li>
<!--				<li><A HREF="ARBillingReport.asp"><B>A/R Billing Report(Personal)</B></A></li>-->
				<li><A HREF="ARBillingReportAll.asp"><B>A/R Billing Report(Complete)</B></A></li>
				<li><A HREF="ARPaymentReport.asp"><B>A/R Payment Report</B></A></li>
				<li><A HREF="ARReminder.asp"><B>A/R Reminder</B></A></li>
<!--				<li><A HREF="ARAgingReport.asp"><B>A/R Aging Report</B></A></li> -->
				<li><A HREF="OutstandingReport.asp"><B>Outstanding inquiry</B></A></li>
				<li><A HREF="TopBillingReport.asp"><B>Top X Bill List</B></A></li>
				<li><A HREF="UnknownCellphoneReport.asp"><B>Unknown Cellphone Bill Report</B></A></li>
				<li><A HREF="SupervisorReminder.asp"><B>Supervisor Reminder Report</B></A></li>
				<li><A HREF="FiscalDataReport.asp"><B>Fiscal Data Report</B></A></li>
			</ul>
		</li>
	<%end if%>
	<%if (UserRole_= "Cashier") or (UserRole_= "Admin") or (UserRole_= "FMC") then %>
		<br><LI class="normal"><B>FMC Cashier :
			<UL TYPE="">
<!--				<li><A HREF="HomePhonePaymentList.asp"><B>Home Phone</B></A></li> -->
<!--				<li><A HREF="OfficePhonePaymentList.asp"><B>Office Phone</B></A></li> -->
				<li><A HREF="PaymentReceiptList.asp"><B>Payment Receipt</B></A></li>
<!--				<li><A HREF="MonthlyBillListAll.asp"><B>Print Monthly Bill</B></A></li> -->
<!--				<li><A HREF="MonthlyBillListAll.asp"><B>Monthly Billing Report</B></A></li>-->
<!--
				<li><B>Reports :</B>
					<UL TYPE="">    -->
						<li><A HREF="ARPaymentReport.asp"><B>A/R Payment Report</B></A></li>
<!--						<li><A HREF="ARReminder.asp"><B>A/R Reminder</B></A></li>
-->
<!--						<li><A HREF="ARAgingReport.asp"><B>A/R Aging Report</B></A></li> 
						<li><A HREF="OutstandingReport.asp"><B>Outstanding inquiry</B></A></li>
					</ul>
				</li>		

				<li><A HREF="ExchangeRateList.asp"><B>Exchange Rate</B></A></li>
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