<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>
<style type="text/css">
<!--
.FontContent {
	font-size: 12px;
        color: blue;
}
-->
</style>
<script type="text/javascript">
function checkall(obj)
{
	var c = document.frmCellPhoneBilling.elements.length
	for (var x=0; x<frmCellPhoneBilling.elements.length; x++)
	{
		cbElement = frmCellPhoneBilling.elements[x]
		if (cbElement.type == "checkbox")
		{
			cbElement.checked= obj.checked?true:false
		}
	}
}
</script>
<% 
 
CellPhone_ = trim(Request("CellPhone"))
'response.write "HomePhone_  :" & HomePhone_ & "<br>"
MonthP_ = Request("MonthP")
'response.write MonthP_ & "<br>"
YearP_ = Request("YearP")
'response.write YearP_ & "<br>"
%> 

<TITLE>U.S. Mission Jakarta e-Billing</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<STYLE TYPE="text/css"><!--
  A:ACTIVE { color:#003399; font-size:8pt; font-family:Verdana; }
  A:HOVER { color:#003399; font-size:8pt; font-family:Verdana; }
  A:LINK { color:#003399; font-size:8pt; font-family:Verdana; }
  A:VISITED { color:#003399; font-size:8pt; font-family:Verdana; }
  body {scrollbar-3dlight-color:#FFFFFF; scrollbar-arrow-color:#E3DCD5; scrollbar-base-color:#FFFFFF; scrollbar-darkshadow-color:#FFFFFF;	scrollbar-face-color:#FFFFFF; scrollbar-highlight-color:#E3DCD5; scrollbar-shadow-color:#E3DCD5; }
  p { font-family: verdana; font-size: 12px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; color: #003399; text-decoration: none}
  h3 { font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 16px; font-style: normal; line-height: normal; font-weight: bold; color: #003399; letter-spacing: normal; word-spacing: normal; font-variant: small-caps}
  td { font-family: verdana; font-size: 10px; font-style: normal; font-weight: normal; color: #000000}
  .title { font-size:16px; font-weight:bold; color:#000080; }
  .SubTitle { font-size:16px; font-weight:bold; color:#000080;  }
  A.menu { text-decoration:none; font-weight:bold; }
  A.mmenu { text-decoration:none; color:#FFFFFF; font-weight:bold; }
  .normal { font-family:Verdana,Arial; color:black}
  .style5 {color: #FFFFFF;}
  .ActivePage {color: red; font-weight:bold; }
--></STYLE>
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
  <Center><FONT COLOR=#009900><B>SENSITIVE BUT UNCLASSIFIED</Center></FONT></B>
  <BR>
<CENTER>
  <IMG SRC="images/embassytitle2.jpeg" WIDTH="661" HEIGHT="80" BORDER="0"> 
  <TABLE WIDTH="65%" BORDER="0" CELLPADDING="0" CELLSPACING="0">
  <CAPTION><H3 STYLE="font-size:17px;color:#000040">Mission Jakarta - Billing Application</H3></CAPTION>
  <TR>
  	<TD COLSPAN="4" class="title" align="center">Cellphone Bill Detail</TD>
   </TR>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<%
TotalCellPhoneBillRp_ = 0
TotalCellPhonePrsBillRp_ = 0

strsql = "Select * From vwMonthlyBilling Where MobilePhone ='" & CellPhone_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "'"
'response.write strsql & "<br>"
set rsCellPhone = server.createobject("adodb.recordset") 
set rsCellPhone = BillingCon.execute(strsql)
if not rsCellPhone.eof then
	EmpID_ = rsCellPhone("EmpID")
	EmpName_ = rsCellPhone("EmpName")
	Office_ = rsCellPhone("Office")
	SupervisorEmail_ = rsCellPhone("SupervisorEmail")
	Notes_ = rsCellPhone("Notes")
	SpvRemark_ = rsCellPhone("SupervisorRemark")
	TotalCellPhoneBillRp_ = rsCellPhone("CellPhoneBillRp")
	TotalCellPhonePrsBillRp_ = rsCellPhone("CellPhonePrsBillRp")
	ProgressID_ = rsCellPhone("ProgressID")
	'response.write "TotalCellPhoneBillRp_ :" & TotalCellPhoneBillRp_ 
	'response.write "TotalCellPhonePrsBillRp_ :" & TotalCellPhonePrsBillRp_ 
end if
%>           
<form method="post" action="CellPhoneDetailSave.asp" name="frmCellPhoneBilling"> 
<table cellspadding="1" cellspacing="0" width="80%" bgColor="white" align="center">  
<%if not rsCellPhone.eof then%>
  <tr>
	<td colspan="6" align="center"><u><b>Billing period (Month - Year) :</b> <a class="FontContent"><%=MonthP_ %> - <%=YearP_ %> </a></u></td>
  </tr>
  <tr>
	<td colspan="6" align="Center">
		<%
			strsql = "Select * From vwCellphone Where PhoneNumber='" & CellPhone_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "'"
			'response.write strsql & "<br>"
			set rsCellPhone = BillingCon.execute(strsql)
			if not rsCellPhone.eof then
				PreviousBalance_= rsCellPhone("PreviousBalance")
				Payment_= rsCellPhone("Payment")
				Adjustment_= rsCellPhone("Adjustment")
				BalanceDue_= rsCellPhone("BalanceDue")
				SubscriptionFee_= rsCellPhone("SubscriptionFee")
				LocalCall_= rsCellPhone("LocalCall")
				Interlocal_= rsCellPhone("Interlocal")
				IDD_= rsCellPhone("IDD")
				SMS_= rsCellPhone("SMS")
				IRS_= rsCellPhone("IRS")
				IRL_= rsCellPhone("IRL")
				Prepaid_= rsCellPhone("Prepaid")
				FARIDA_= rsCellPhone("FARIDA")
				MobileBanking_= rsCellPhone("MobileBanking")
				DetailedCallRecord_= rsCellPhone("DetailedCallRecord")
				Internet_= rsCellPhone("Internet")
				FARIDA_= rsCellPhone("FARIDA")
				DataRoam_= rsCellPhone("DataRoam")
				MinUsage_= rsCellPhone("MinUsage")
				SubTotal_= rsCellPhone("SubTotal")
				PPN_= rsCellPhone("PPN")
				StampFee_= rsCellPhone("StampFee")
				CurrentBalance_= rsCellPhone("CurrentBalance")
				Total_= rsCellPhone("Total")
			end if
		%>
		<table cellspadding="0" cellspacing="0" bordercolor="black" border="1" width="90%" bgColor="white">  
		<tr>
			<td colspan="4" align="center" class="SubTitle">USAGE SUMMARY
			</td>
		</tr>
		<tr>
			<td>&nbsp;<b>Previous Balance</b> / <i>Tagihan Sebelumnya</i><div align="center">Rp.&nbsp;&nbsp;<%=formatnumber(PreviousBalance_,0) %></div></td>
			<td>&nbsp;<b>Payment</b> / <i>Pembayaran</i><div align="center">Rp.&nbsp;&nbsp;<%=formatnumber(Payment_,0) %></div></td>
			<td>&nbsp;<b>Adjustment</b> / <i>Koreksi</i><div align="center">Rp.&nbsp;&nbsp;<%=formatnumber(Adjustment_,0) %></div></td>
			<td>&nbsp;<b>Balance Due</b> / <i>Sisa Tagihan</i><div align="center">Rp.&nbsp;&nbsp;<%=formatnumber(BalanceDue_,0) %></div></td>
		</tr>
		<tr>
			<td colspan="4">&nbsp;<u><b>Usage Charges</b> / <i>Biaya Percakapan:<i/></u></td>
		</tr>
		<tr>
			<td colspan="4">
			<table cellspadding="0" cellspacing="0" bordercolor="black" width="100%" bgColor="white">  
			<tr>
				<td width="70%">&nbsp;<b>Subscription Fee</b> / <i>Abonemen<i/></td>
				<td width="3%">&nbsp;Rp.</td>
				<td align="right"><%=formatnumber(SubscriptionFee_,0) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>Local</b> / <i>Lokal<i/></td>
				<td>&nbsp;Rp.</td>
				<td align="right"><%=formatnumber(LocalCall_,0) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>Interlocal</b> / <i>SLJJ<i/></td>
				<td>&nbsp;Rp.</td>
				<td align="right"><%=formatnumber(Interlocal_,0) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>IDD</b> / <i>SLI<i/></td>
				<td>&nbsp;Rp.</td>
				<td align="right"><%=formatnumber(IDD_,0) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>SMS</b> / <i>SMS<i/></td>
				<td>&nbsp;Rp.</td>
				<td align="right"><%=formatnumber(SMS_,0) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>International Roaming Surcharge</b> / <i>Surcharge Jelajah Internasional<i/></td>
				<td>&nbsp;Rp.</td>
				<td align="right"><%=formatnumber(IRS_,0) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>International Roaming Leg</b> / <i>Roaming Leg Jelajah Internasional<i/></td>
				<td>&nbsp;Rp.</td>
				<td align="right"><%=formatnumber(IRL_,0) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td colspan="4">&nbsp;<u><b>Value Added Services</b> / <i>Layanan Tambahan:<i/></u></td>
		</tr>
		<tr>
			<td colspan="4">
			<table cellspadding="0" cellspacing="0" bordercolor="black" width="100%" bgColor="white">  
			<tr>
				<td width="70%">&nbsp;<b>Prepaid Recharge</b> / <i>Isi Ulang Prabayar<i/></td>
				<td width="3%">&nbsp;Rp.</td>
				<td align="right"><%=formatnumber(Prepaid_,0) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>Fax Response and Interactive Data</b> / <i>FARIDA<i/></td>
				<td>&nbsp;Rp.</td>
				<td align="right"><%=formatnumber(FARIDA_,0) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>Mobile Banking</b> / <i>Mobile Banking<i/></td>
				<td>&nbsp;Rp.</td>
				<td align="right"><%=formatnumber(MobileBanking_,0) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>Detailed Call Record Print</b> / <i>Print Rincian Percakapan<i/></td>
				<td>&nbsp;Rp.</td>
				<td align="right"><%=formatnumber(DetailedCallRecord_,0) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>3G, HSDPA, GPRS, MMS, Wifi, Premium Content</b> / <i>3G, HSDPA, GPRS, MMS, Wifi, Konten Premium<i/></td>
				<td>&nbsp;Rp.</td>
				<td align="right"><%=formatnumber(Internet_,0) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>Ventus / Blackberry, iPhone, Bridge Dataroam, Data Roam</b> / <i>Ventus / Blackberry, iPhone, Bridge Dataroam, Data Roam<i/></td>
				<td>&nbsp;Rp.</td>
				<td align="right"><%=formatnumber(DataRoam_,0) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td colspan="4">
			<table cellspadding="0" cellspacing="0" bordercolor="black" width="100%" bgColor="white">  
			<tr>
				<td width="70%">&nbsp;<b>Variance To Minimum Usage Guarantee</b> / <i>Selisih Penggunaan Minimum<i/></td>
				<td width="3%">&nbsp;Rp.</td>
				<td align="right"><%=formatnumber(MinUsage_,0) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>Sub Total</b> / <i>Sub Total<i/></td>
				<td>&nbsp;Rp.</td>
				<td align="right"><%=formatnumber(SubTotal_,0) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>VAT 10%</b> / <i>PPN 10%<i/></td>
				<td>&nbsp;Rp.</td>
				<td align="right"><%=formatnumber(PPN_,0) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>Stamp Duty Fee</b> / <i>Biaya Materai pembayaran bulan lalu)<i/></td>
				<td>&nbsp;Rp.</td>
				<td align="right"><%=formatnumber(StampFee_,0) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>Current Balance</b> / <i>Total Tagihan Bulan Ini<i/></td>
				<td>&nbsp;Rp.</td>
				<td align="right"><u><b><%=formatnumber(CurrentBalance_,0) %></b></u>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			</table>
			</td>
		</tr>
<!--
		<tr>
			<td colspan="4">
			<table cellspadding="0" cellspacing="0" bordercolor="black" width="100%" bgColor="white">  
			<tr>
				<td width="70%">&nbsp;<b>Amount Due To be Paid</b> / <i>Jumlah yang harus dibayarkan<i/></td>
				<td width="3%">&nbsp;Rp.</td>
				<td align="right"><u><b><%=formatnumber(Total_,0) %></b></u>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			</table>
			</td>
		</tr>
-->
		<tr>
			<td colspan="4" align="center" class="SubTitle">Usage Detail:</td>
		</tr>
		</table>
	</td>	
  </tr>
  <tr>
	<td colspan="6" align="Center">
		<table cellspadding="0" cellspacing="0" bordercolor="black" border="1" width="90%" bgColor="white">  
		<tr bgcolor="#330099" align="center" cellpadding="0" cellspacing="0" >
			<TD width="5%"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
		       	<TD><strong><label STYLE=color:#FFFFFF>Dialed Date/time</label></strong></TD>
			<TD width="20%"><strong><label STYLE=color:#FFFFFF>Dialed Number</label></strong></TD>
			<TD><strong><label STYLE=color:#FFFFFF>Call Type</label></strong></TD>
			<TD><strong><label STYLE=color:#FFFFFF>Duration</label></strong></TD>
			<TD width="10%"><strong><label STYLE=color:#FFFFFF>Amount (Rp.)</label></strong></TD>
			<TD width="10%"><strong><label STYLE=color:#FFFFFF>Personal used</label></strong><br>
<%			if (ProgressID_ = 1) or (ProgressID_ = 3) then %>
				<input type="checkbox" name="cbAll" value="true" onclick="checkall(this)" />
<%			end if %>
			</TD>
		</tr>
		<%
		strsql = "Select * from CellPhoneDt Where PhoneNumber='" & CellPhone_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "'"
		'response.write strsql & "<br>"
		set rsCellPhone = BillingCon.execute(strsql) 
		no_ = 1 
		do while not rsCellPhone.eof
   			if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
		%>
			<tr bgcolor="<%=bg%>">
				<td align="right"><%=No_%>&nbsp;</td>
			        <td><FONT color=#330099 size=2>&nbsp;<%=rsCellPhone("DialedDatetime")%></font></td> 
		        	<td><FONT color=#330099 size=2>&nbsp;<%=rsCellPhone("DialedNumber")%></font></td>
		        	<td><FONT color=#330099 size=2>&nbsp;<%=rsCellPhone("CallType")%></font></td>
		        	<td><FONT color=#330099 size=2>&nbsp;<%=rsCellPhone("CallDuration")%></font></td>
			        <td align="right"><FONT color=#330099 size=2><%=formatnumber(rsCellPhone("Cost"),0)%>&nbsp;</font></td> 
<%'			if (ProgressID_ = 1) or (ProgressID_ = 3) then %>
<%			if cdbl(ProgressID_)<4 then %>
			        <td align="center">
				<%if rsCellPhone("isPersonal") = "Y" then%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsCellPhone("CallRecordID")%>' Checked>				
				<%else%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsCellPhone("CallRecordID")%>' >
				<%end if%>
				</td>
<%			else%>
			        <td align="center">
				<%if rsCellPhone("isPersonal") = "Y" then%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsCellPhone("CallRecordID")%>' Checked disabled>				
				<%else%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsCellPhone("CallRecordID")%>'  disabled>
				<%end if%>
				</td>
<%			end if %>
		<%   
			rsCellPhone.movenext
			no_ = no_ + 1
		loop
		%>
		</tr>
		</table>
	</td>
  </tr>
  <tr>
	<td colspan="6" align="center">
		<table cellspadding="0" cellspacing="0" bordercolor="black" width="90%" bgColor="white">  
		<tr>
			<td align="right" colspan="3"><b>Sub Total (Rp.) </b>&nbsp;</td>
			<td width="15%" class="FontContent" align="right"><b><%=formatnumber(TotalCellPhoneBillRp_ ,0)%></b>&nbsp;</td>
			<td width="10%" class="FontContent" align="right"><b><u><%=formatnumber(TotalCellPhonePrsBillRp_ ,0)%></u></b>&nbsp;</td>
		</tr>
		</table>
	</td>
  </tr>
  <tr>
	<td colspan="6" align="center">&nbsp;</td>
  </tr>
<%'		if (ProgressID_ = 1) or (ProgressID_ = 3) then%>
<%	
		'response.write "No:" & no_
		if (ProgressID_)< 4 and no_ >1 then%>
  <tr>
	<td colspan="6" align="center">	
		<input type="submit" name="btnSubmit" Value="Update Change(s)" />&nbsp;&nbsp;
		<input type="hidden" name="txtCellPhone" value='<%=CellPhone_ %>' />
		<input type="hidden" name="txtMonthP" value='<%=MonthP_%>' />
		<input type="hidden" name="txtYearP" value='<%=YearP_%>' />
		<input type="hidden" name="txtEmpID" value='<%=EmpID_ %>' />
	<td>
  </tr>
	<%end if%>
<%else%>
<tr>
	<td colspan="6" align="center">&nbsp;</td>	
</tr>
<tr>
	<td colspan="6" align="center">there is not data.</td>	
</tr>
<%end if%>
</table>
</form>
</BODY>
</html>