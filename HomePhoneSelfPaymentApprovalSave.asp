<HTML>
<HEAD>
<!--#include file="connect.inc" -->
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 
<style type="text/css">

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
   EmpID_ = request.form("txtEmpID")
   HomePhone_ = request.form("txtHomePhone")
   ReceiptNo_ = Request.Form("txtReceiptNo")
'   response.write ReceiptNo_ & "<br>"
   PaidAmount_ = Request.Form("txtPaidAmount")
   PaidDate_  = Request.Form("txtPaymentDate")
'   response.write PaidDate_ & "<br>"
   CashierRemark_ = replace(Request.Form("txtCashierRemark"),"'","''")
'   response.write CashierRemark_ & "<br>"
   MonthP_ = Request.Form("txtMonthP")
   'response.write MonthP_ & "<br>"
   YearP_ = Request.Form("txtYearP")
   'response.write YearP_ & "<br>"
   TotCost_ = Request.Form("txtTotCost")
   'response.write TotCost_ & "<br>"
   if cdbl(PaidAmount_)>cdbl(TotCost_) then
	TotPaid_ = TotCost_
   else
	TotPaid_ = PaidAmount_
   end if
%>

<%
'4. SHOW
%>


<p>
<table border="0" align=center width="100%" cellspacing="0" cellpadding="1">
</tr>
<tr>
	<td colspan="2" align="center">Billing Period : <Label style="color:blue"><%=MonthP_%> - <%=YearP_%></lable></td>
</tr>
<tr>
	<td colspan="2" align="center"><br></td>
</tr>
<%
	'3. SAVING TO Billing Header
	strsql = "Update HomePhone Set Status='P', PaidAmount=" & PaidAmount_ & ", ReceiptNo='" & ReceiptNo_ & "', PaidDate ='" & PaidDate_ & "', CashierRemark ='" & CashierRemark_ & "' Where Nomor='" & HomePhone_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ & "'"
	'response.write strsql
	BillingCon.execute(strsql)

	'Update to MonthBill
	strsql = "Update MonthlyBilling Set PaidAmount=" & TotPaid_ & " Where EmpID='" & EmpID_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ & "'"
	'response.write strsql
	BillingCon.execute(strsql)


%>
<tr>
	<td align="center" colspan="2">This data has already been saved.</td>
</tr>
<tr>
	<td colspan="2"><br></td>
</tr>
<tr>
	<td colspan="2" align=center>
		<input type="button" value="Back" id="btnMain" onclick="javascript:document.location.href('HomePhoneSelfPaymentList.asp')">
	        &nbsp;&nbsp;<input type="button" value="Close this window" onclick="javascript:window.close();" name=btnclose>
        </td>
</tr>
</table>
</p>
</BODY>
</HTML>