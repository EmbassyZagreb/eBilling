<HTML>
<HEAD>
<!--#include file="connect.inc" -->
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 

<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
   <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">PAYMENT OF HOME PHONE UPDATE</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>

<%
   EmpID_= request.form("txtEmpID")
   MonthP_ = Request.Form("txtMonthP")
   'response.write MonthP_ & "<br>"
   YearP_ = Request.Form("txtYearP")
   'response.write YearP_ & "<br>"
   MobilePhone_ = Request.Form("txtCellPhone")   
   ReceiptNo_ = Request.Form("txtReceiptNo")
'   response.write ReceiptNo_ & "<br>"
'   CurrencyType_ = Request.Form("cmbCurrencyType")
   CurrencyType_ = "Kn"
   ExchangeRate_ = Request.Form("txtExchangeRate")
   PaidAmount_ = Request.Form("txtPaidAmount")
   If (CurrencyType_ ="Dlr") Then
	PaidAmountDlr_ = PaidAmount_
	PaidAmountRp_ = cdbl(PaidAmount_) *  cdbl(ExchangeRate_)
   Else
	PaidAmountDlr_ = Round(cdbl(PaidAmount_)/cdbl(ExchangeRate_),2)
	PaidAmountRp_ = PaidAmount_
   End If
   PaidDate_  = Request.Form("txtPaymentDate")
'   response.write PaidDate_ & "<br>"
   CashierRemark_ = replace(Request.Form("txtCashierRemark"),"'","''")
'   response.write CashierRemark_ & "<br>"
   'PaymentType_  = Request.Form("cmbPaymentType")
   srStatus_ = Request.Form("txtsrStatus")
   srEmpName_ = Request.Form("txtsrEmpName")

srsMonthP_ = Request.Form("txtsrsMonthP")
srsYearP_ = Request.Form("txtsrsYearP")
sreMonthP_ = Request.Form("txtsreMonthP")
sreYearP_ = Request.Form("txtsreYearP")
srOffice_ = Request.Form("txtsrOffice")
srPageIndex_ = Request.Form("txtsrPageIndex")


   RemainingBillRp_ = Request.Form("txtRemainingBillRp")
   If (cdbl(RemainingBillRp_) > cdbl(PaidAmountRp_)) Then PaymentType_ = "P" Else PaymentType_ = "F"

%>

<%
'4. SHOW
%>


<p>
<table border="0" align=center width="100%" cellspacing="0" cellpadding="1">
</tr>
<tr>
	<!--<td colspan="2" align="center">Billing Period : <Label style="color:blue"><%=MonthP_%> - <%=YearP_%></lable></td> -->
</tr>
<tr>
	<td colspan="2" align="center"><br></td>
</tr>
<%
	'3. SAVING TO Billing Header
	'strsql = "spPaymentReceipt_IUD 'I',0,'" & EmpID_ & "','" & MonthP_ & "','" & YearP_ & "','" & ReceiptNo_ & "','" & CurrencyType_ & "'," & PaidAmountDlr_ & "," & PaidAmountRp_ & ",'" & PaidDate_ & "','" & CashierRemark_ & "','" & PaymentType_ & "'"
	strsql = "spPaymentReceipt_IUD 'I',0,'" & EmpID_ & "','" & MonthP_ & "','" & YearP_ & "','" & MobilePhone_ & "','" & ReceiptNo_ & "','" & CurrencyType_ & "'," & PaidAmountDlr_ & "," & PaidAmountRp_ & ",'" & PaidDate_ & "','" & CashierRemark_ & "','" & PaymentType_ & "'"
	'response.write strsql
	BillingCon.execute(strsql)

%>
<tr>
	<td align="center" colspan="2">This data has already been saved.</td>

</tr>
<tr>
	<td align="center" colspan="2">You will be automatically redirected...</td>
</tr>
<!-- <tr>
	<td colspan="2" align=center>
		<input type="button" value="Back" id="btnMain" onclick="javascript:document.location.href('PaymentReceiptList.asp')">
	        &nbsp;&nbsp;<input type="button" value="Close this window" onclick="javascript:window.close();" name=btnclose>
        </td>
</tr>  -->
</table>
<%
Response.AddHeader "REFRESH","1;URL=PaymentReceiptList.asp?PageIndex=" & srPageIndex_ & "&sMonthP=" & srsMonthP_ & "&sYearP=" & srsYearP_ & "&eMonthP=" & sreMonthP_ & "&eYearP=" & sreYearP_ & "&Status=" & srStatus_ & "&EmpName=" & srEmpName_ & "&OfficeSection=" & srOffice_ & ""
					
%>


</p>
</BODY>
</HTML>