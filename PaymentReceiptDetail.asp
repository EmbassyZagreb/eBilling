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

strsql = "Select * From vwPaymentReceiptHistory Where EmpID='" & EmpID & "' and MonthP='" & MonthP & "' and YearP='" & YearP & "'"
'response.write strsql & "<br>"
set rsPayment = server.createobject("adodb.recordset") 
set rsPayment = BillingCon.execute(strsql) 
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
	<td colspan="6" align="center">&nbsp;</td>	
</tr>
<tr>
	<td align="Left" colspan="6"><u><b>Payment History:<b></u></TD>
</tr>
<tr>
	<td align="Left" colspan="6">

<table cellspadding="1" cellspacing="0" width="100%" bgColor="white">  
<%if not rsPayment.eof then%>

<tr>
	<td colspan="6">
	<table align="center" cellpadding="1" cellspacing="0" width="100%" border="1" bordercolor="black"> 
	<TR BGCOLOR="#FFFFFF" align="center" cellpadding="0" cellspacing="0" height=26>
		<TD width="3%"><strong><label>No.</label></strong></TD>
		<TD width="17%"><strong><label>Receipt No</label></strong></TD>
		<TD width="20%"><strong><label>Paid Date</label></strong></TD>
		<TD width="20%"><strong><label>Cashier Remark</label></strong></TD>
		<TD width="20%"><strong><label>Payment Type</label></strong></TD>
	<!--	<TD width="18%"><strong><label>Currency</label></strong></TD>  -->
		<TD width="20%"><strong><label>Paid Amount (Kn.)</label></strong></TD>
	</TR>    
<% 
		dim no_ , TotalQty_ , TotalAmount_ 
		no_ = 1 
		TotalQty_ = 0
		TotalAmount_ = 0
		do while not rsPayment.eof 
	   	if bg="#DDDDDD" then bg="ffffff" else bg="#DDDDDD" 
%> 
 	<TR bgcolor="<%=bg%>">
		<td align="right"><%=No_%>&nbsp;</td>
        	<td>&nbsp;<%=rsPayment("ReceiptNo")%></font></td> 
		<td align="right"><%=rsPayment("PaidDate")%>&nbsp;</font></td> 
        	<td align="right"><%=rsPayment("CashierRemark")%>&nbsp;</font></td> 
		<td align="right"><%=rsPayment("PaymentType")%>&nbsp;</font></td>
	<!--		<td align="right"><FONT color=#330099 size=2><%=rsPayment("Currency")%>&nbsp;</font></td>  -->
        	<td align="right"><%=formatnumber(rsPayment("PaidAmount"),-1)%>&nbsp;</font></td> 
  	 </TR>
<%   
		TotalAmount_ = TotalAmount_ + formatnumber(rsPayment("PaidAmount"),2)
		'response.write TotalAmount_ 
 		rsPayment.movenext
   		no_ = no_ + 1
	loop
%>	
	<tr BGCOLOR="#999999" align="center" cellpadding="0" cellspacing="0" height=26>
		<td align="left" colspan="5"><FONT color=#FFFFFF><b>&nbsp;Total</b></font></td>
		<td width="20%" align="right"><FONT color=#FFFFFF><b><%=formatnumber(TotalAmount_  ,-1)%></b></font>&nbsp;</td>
	</tr>
	</table>
	</td>
</tr>

<%else%>
<tr>
	<td colspan="6" align="center">&nbsp;</td>	
</tr>
<tr>
	<td colspan="6" align="center">there is no payment data. Because the personal usage amount is less than the threshold amount.</td>
</tr>
<%end if%>
</table>

	</TD>
</tr>







<tr>
	<td align="center" colspan="6"><br>
	</td>
</tr>



<tr>
	<td align="center" colspan="3">
       		&nbsp;<input type="button" value="Back" onClick="javascript:location.href='PaymentReceiptList.asp?PageIndex=<%=PageIndex_%>&MonthP=<%=MonthP%>&YearP=<%=YearP%>&sMonthP=<%=srsMonthP_%>&sYearP=<%=srsYearP_%>&eMonthP=<%=sreMonthP_%>&eYearP=<%=sreYearP_%>&EmpName=<%=srEmpName_%>&Status=<%=srStatus_%>&OfficeSection=<%=srOfficeSection_%>'">
   	</td>
</tr>



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