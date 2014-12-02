<HTML>
<HEAD>
<!--#include file="connect.inc" -->
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 
<TITLE>U.S. Embassy Zagreb - zBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
<%
EmpID_ = Request("EmpID")
'response.write EmpID_ & "<br>"
sMonthP_ = request("sMonthP")
sYearP_ =  request("sYearP")
eMonthP_ =  request("eMonthP")
eYearP_ =  request("eYearP")
ProgressID_ =  request("ProgressID")
rbPeriod_ = request("rbPeriod")
'SendMailStatusID_ = request("SendMailStatusID")
'response.write "rbPeriod_  :" & rbPeriod_ 
if rbPeriod_ ="S" then
	sPeriod_ = sYearP_&sMonthP_
	ePeriod_ = eYearP_&eMonthP_
End if
BillType_ = Request("BillType")

curMonth_ = month(date())
curYear_ = year(date())
ccurYear_ = year(date())
if len(curMonth_)= 1 then
	curMonth_ = "0" & curMonth_
end if

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
TotalBillingRp_ = 0
TotalBillingDlr_ = 0

'response.write YearP_ & "<br>"
if rbPeriod_ = "X" then
'	strsql = "Select * From vwMonthlyBilling Where ProgressID in(4,8) And EmpID='" & EmpID_ & "' And (ProgressID=" & ProgressID_ & " or " & ProgressID_ & " =0) and SendMailStatusID=" &SendMailStatusID_ 
	strsql = "Select * From vwMonthlyBilling Where ProgressID in(4,8) And EmpID='" & EmpID_ & "' And (ProgressID=" & ProgressID_ & " or " & ProgressID_ & " =0)"
else
	strsql = "Select * From vwMonthlyBilling Where ProgressID in(4,8) And EmpID='" & EmpID_ & "' And (ProgressID=" & ProgressID_ & " or " & ProgressID_ & " =0) and YearP+MonthP>='" & sPeriod_ & "' and YearP+MonthP<='" & ePeriod_ & "'"
end if	

'response.write BillType_ & "<Br>"  
'response.write strsql & "<Br>"  
set rsData = server.createobject("adodb.recordset") 
set rsData = BillingCon.execute(strsql)
if not rsData.eof then
	Period_ = rsData("MonthP") & rsData("YearP")
	EmpName_ = rsData("EmpName")
	Office_ = rsData("Agency") & " - " & rsData("Office")
	Position_ = rsData("WorkingTitle")
	OfficePhone_ = rsData("WorkPhone")
	HomePhone_ = rsData("HomePhone")
	MobilePhone_ = rsData("MobilePhone")	
	EmpEmail_ = rsData("EmailAddress")
	LoginID_ = rsData("LoginID")
	FiscalData_ = rsData("FiscalStripVAT")
end if

if EmpEmail_ <>"" Then
	If Right(EmpID_,1)="N" then			
		LoginID_ = EmpID_
	End If
%>               
	<center>
	<IMG SRC='http://zagrebws03:8080/zBilling/images/embassytitle2.jpeg' HEIGHT='80' BORDER='0'><br>
	<table cellspadding='1' cellspacing='0' width='100%' bgColor='white'>
		<tr><td colspan='6'>&nbsp;</td></tr>
	<tr><td colspan='2' width='50%'>
		<table cellspacing='0' border='1' bordercolor='black'>
			<tr><td colspan='2' align='Center'><strong>Payee Info</strong></td></tr>
			<tr><td align='right'><i>Employee Name : </i></td><td><%=EmpName_%></td></tr>
			<tr><td align='right'><i>Department : </i></td><td><%=Office_%></td></tr>
		</table>
	     </td>

	    <td colspan='4' width='50%' align='right'>
		<table cellspacing='0'>
		<tr><td colspan='4'><strong><font color='red' size='5'>Bill of Collection</font></td></tr>
			<tr><td colspan='2'><i>Bill of Collection Number : </i></td><td colspan='2'><%=LoginID_%><%=curMonth_ %><%=ccurYear_ %></td></tr>
			<tr><td colspan='2'><i>Bill of Collection Date : </i></td><td colspan='2'><%=Date()%></td>
<!--
			<tr><td colspan='2' align='right'><i>Bill of Collection Number : </i></td><td colspan='2'><%=LoginID_%><%=curMonth_ %><%=ccurYear_ %></td></tr>
			<tr><td colspan='2' align='right'><i>Bill of Collection Date : </i></td><td colspan='2'><%=Date()%></td>
-->
		</tr>
		</table>
	     </td>
	</tr>
	<tr><td colspan='6'>&nbsp;</td></tr>
	<tr><td colspan='6'>&nbsp;</td></tr>
	<tr><td colspan='6'><table cellspacing='0' width='100%' border='1' bordercolor='black'>
	<tr align='center'><td><strong>Billing Period</strong></td><td><strong>Bill Type</strong></td><td><strong>Description</strong></td><td><strong>Amount Due (USD)</strong></td><td><strong>Exchange Rate</strong></td><td><strong>Amount Due (Kn)</strong></td></tr>
<%
		TotalBillRp_ = 0
		TotalBillDlr_ = 0
		 Do while not rsData.eof
			TotalBillRp_ = cdbl(TotalBillRp_) + cdbl(rsData("TotalBillingAmountPrsRp"))
			TotalBillDlr_ = cdbl(TotalBillDlr_) + cdbl(rsData("TotalBillingAmountPrsDlr"))
%>
	<tr><td>&nbsp;<%=rsData("YearP")%> - <%=rsData("MonthP")%></td><td>Mobile Phone</td><td>&nbsp;<%=rsData("MobilePhone")%></td><td align='right'>$ <%=formatnumber(cdbl(rsData("TotalBillingAmountPrsDlr")),-1)%>&nbsp;</td><td align='right'><%=formatnumber(rsData("ExchangeRate"),-1)%>&nbsp;</td><td align='right'>IDR <%=formatnumber(cdbl(rsData("TotalBillingAmountPrsRp")),-1)%>&nbsp;</td></tr>
<%
		 	rsData.movenext
		 Loop 
%>
	<tr><td colspan='3' align='center'><strong>Total</strong></td><td align='right'><strong>$ <%=formatnumber(TotalBillDlr_,-1)%></strong>&nbsp;</td><td>&nbsp;</td><td align='right'><strong>IDR <%=formatnumber(TotalBillRp_,-1)%></strong>&nbsp;</td></tr></table></td></tr>
	</table><div align='right'><i>*Please remit payment within 15 days of the Invoice Date</i></div>
	<br><br><br><div align='left'><strong>Payment Options</strong></div>
	<div align='left'><a href='https://pay.gov/paygov/forms/formInstance.html?nc=1382515161514&agencyFormId=42156675&userFormSearch=https%3A%2F%2Fpay.gov%2Fpaygov%2FagencySearchForms.html%3FshowingDetails%3Dtrue%26showingAll%3Dfalse%26sortProperty%3DagencyFormName%26totalResults%3D4%26nc%3D' target="new">Pay online with pay.gov (USD Only)</a></div>
	<div align='left'>Pay at the embassy cashier:</div>
	<div align='left'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & CashierInfo & "</div>
	<br><br><br><div align='left'><strong>Fund Cite Details (For Cashier)</strong></div>
	<div align='left'>&nbsp;&nbsp;&nbsp;$ <%=formatnumber(TotalBillDlr_,-1)%> 19-X45190001-00-5306-53063R3221-6195-2322 (ICASS Mobile Phone)</div>
	<br><br><br><br><br><div align='left'>Please contact <a href='mailto:zgbphonebill@state.gov'>zgbphonebill@state.gov</a> with any questions</div>
<%
end if

'  response.redirect("BillingApprovalList.asp")
%>
<table cellspadding="1" cellspacing="0" width="100%" align="center">
<tr>
	<td align="center"><input type="button" value="Close" id="btnclose" onclick="window.close()"></td>
</tr>
</table>
</body> 
</html>