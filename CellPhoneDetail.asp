<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>

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
AlternateEmailFlag_ = trim(Request("AlternateEmailFlag"))

SortBy_ = Request.Form("SortList")
'response.write "SortBy" & SortBy_
if (SortBy_ ="") then
	if Request("SortBy")<>"" then
		SortBy_ = Request("SortBy")
	Else
		SortBy_ = "DialedDatetime"
	end if
end if

Order_ = Request("OrderList")	
if (Order_ ="") then
	if Request.Form("OrderList")<>"" then
		Order_ = Request.Form("OrderList")
	Else
		Order_ = "Asc"
	end if
end if
%> 

<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
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
	MobilePhone_ = rsCellPhone("MobilePhone")
	SupervisorEmail_ = rsCellPhone("SupervisorEmail")
	Notes_ = rsCellPhone("Notes")
	SpvRemark_ = rsCellPhone("SupervisorRemark")
	TotalCellPhoneBillRp_ = rsCellPhone("CellPhoneBillRp")
	TotalCellPhonePrsBillRp_ = rsCellPhone("CellPhonePrsBillRp")
	ProgressID_ = rsCellPhone("ProgressID")
	FiscalStripNonVAT_ = rsCellPhone("FiscalStripNonVAT")
	Status_ = rsCellPhone("ProgressDesc")
	'response.write "TotalCellPhoneBillRp_ :" & TotalCellPhoneBillRp_ 
	'response.write "TotalCellPhonePrsBillRp_ :" & TotalCellPhonePrsBillRp_ 
end if
'response.write ProgressID_  & "<br>"
strsql = "Select DetailRecordAmount From PaymentDueDate"
'response.write strsql & "<br>"
set rsDetailRecord = server.createobject("adodb.recordset") 
set rsDetailRecord = BillingCon.execute(strsql)
if not rsDetailRecord.eof then
	DetailRecordAmount_ = rsDetailRecord("DetailRecordAmount")
	'response.write "DetailRecordAmount :" & DetailRecordAmount_ 
end if

%>           
<table cellspadding="1" cellspacing="0" width="80%" bgColor="white" align="center">  
<%if not rsCellPhone.eof then%>
  <tr>
	<td colspan="6" align="center"><u><b>Billing period (Month - Year) :</b> <a class="FontContent"><%=MonthP_ %> - <%=YearP_ %> </a></u></td>
  </tr>
  <tr>
          <td colspan="6" align="Left"><u><b>Personal Info<b></u></TD>
  </tr>  
  <tr>
	<td width="20%">Employee Name</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=EmpName_%></td>
	<td class="FontContent" colspan="3"><%=FiscalStripNonVAT_ %></td>
  </tr>
  <tr>
	<td>Mobile Phone</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=MobilePhone_%></td>
	<td width="15%">Status</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=Status_ %></td>
  </tr>
  <tr>
	<td colspan="6"><hr></td>
  </tr>
  <tr>
	<td colspan="6" align="Center">
		<%
			strsql = "Select * From vwCellphoneHd Where PhoneNumber='" & CellPhone_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "'"
			'response.write strsql & "<br>"
			set rsCellPhone = BillingCon.execute(strsql)
			if not rsCellPhone.eof then
				PreviousBalance_= rsCellPhone("PreviousBalance")
				Payment_= rsCellPhone("Payment")
				Adjustment_= rsCellPhone("Adjustment")
				BalanceDue_= rsCellPhone("BalanceDue")
				SubscriptionFee_= rsCellPhone("SubscriptionFee")
				LocalCall_= rsCellPhone("LocalCall")
				Interlocal_= rsCellPhone("SLJJ")
				IDD_= rsCellPhone("SLI")
				SMS_= rsCellPhone("SMS")
				IRL_= rsCellPhone("IRL")
				Prepaid_= rsCellPhone("Prepaid")
				FARIDA_= rsCellPhone("FARIDA")
				MobileBanking_= rsCellPhone("MobileBanking")
				DetailedCallRecord_= rsCellPhone("DetailedCallRecord")				
				GPRS_= rsCellPhone("GPRS")
				IPHONE_= rsCellPhone("IPHONE")
				'FARIDA_= rsCellPhone("FARIDA")
				'DataRoam_= rsCellPhone("DataRoam")
				MinUsage_= rsCellPhone("MinUsage")
				DiskonBicara_= rsCellPhone("DiskonBicara")
				GPRS_= rsCellPhone("GPRS")
				DiskonSMS_= rsCellPhone("DiskonSMS")
				DiskonGPRS_= rsCellPhone("DiskonGPRS")
				DiskonMMS_= rsCellPhone("DiskonMMS")
				DiskonPenggunaan_= rsCellPhone("DiskonPenggunaan")
				SubTotalTKP_= rsCellPhone("SubTotalTKP")
				SubTotalKP_= rsCellPhone("SubTotalKP")
				PPN_= rsCellPhone("PPN")
				StampFee_= rsCellPhone("StampFee")
				CurrentBalance_= rsCellPhone("CurrentBalance")
				Total_= rsCellPhone("Total")
			end if
		%>
		<table cellspadding="0" cellspacing="0" bordercolor="black" border="1" width="100%" bgColor="white">  
		<tr>
			<td colspan="4" align="center" class="SubTitle">USAGE SUMMARY
			</td>
		</tr>
		<tr>
			<td colspan="4">&nbsp;<u><b>Monthly Fees</b> / <i>Mjesecne pretplate:<i/></u></td>
		</tr>
		<tr>
			<td colspan="4">
			<table cellspadding="0" cellspacing="0" bordercolor="black" width="100%" bgColor="white">  
			<tr>
				<td width="70%">&nbsp;<b>Subscription Monthly Fee</b> / <i>Mjesecna naknada za pretplatnicki broj<i/></td>
				<td width="3%">&nbsp;Kn.</td>
				<td align="right"><%=formatnumber(SubscriptionFee_,-1) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>Data Monthly Fee</b> / <i>Mjesecna naknada za mobilni prijenos podataka<i/></td>
				<td>&nbsp;Kn.</td>
				<td align="right"><%=formatnumber(FARIDA_,-1) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>Other Charges</b> / <i>Ostale usluge<i/></td>
				<td>&nbsp;Kn.</td>
				<td align="right"><%=formatnumber(DetailedCallRecord_,-1) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td colspan="4">&nbsp;<u><b>Usage Charges</b> / <i>Pozivi i prijenos podataka:<i/></u></td>
		</tr>
		<tr>
			<td colspan="4">
			<table cellspadding="0" cellspacing="0" bordercolor="black" width="100%" bgColor="white">  
			<tr>
				<td>&nbsp;<b>VPN Network Calls</b> / <i>Pozivi unutar VPN mre�e<i/></td>
				<td>&nbsp;Kn.</td>
				<td align="right"><%=formatnumber(LocalCall_,-1) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>Calls to VIP Network</b> / <i>Pozivi prema VIP mobilnoj mre�i<i/></td>
				<td>&nbsp;Kn.</td>
				<td align="right"><%=formatnumber(BalanceDue_,-1) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>Calls to Landlines in Croatia</b> / <i>Pozivi prema fiksnim mre�ama u Hrvatskoj<i/></td>
				<td>&nbsp;Kn.</td>
				<td align="right"><%=formatnumber(Interlocal_,-1) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>Calls to Other Mobile Networks</b> / <i>Pozivi prema ostalim mobilnim mre�ama<i/></td>
				<td>&nbsp;Kn.</td>
				<td align="right"><%=formatnumber(IDD_,-1) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>SMS</b> / <i>SMS poruke<i/></td>
				<td>&nbsp;Kn.</td>
				<td align="right"><%=formatnumber(SMS_,-1) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>MMS</b> / <i>MMS Poruke<i/></td>
				<td>&nbsp;Kn.</td>
				<td align="right"><%=formatnumber(GPRS_,-1) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>International Calls from Croatia</b> / <i>Medunarodni pozivi iz Hrvatske<i/></td>
				<td>&nbsp;Kn.</td>
				<td align="right"><%=formatnumber(IRL_,-1) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td width="70%">&nbsp;<b>Incoming Calls in Roaming</b> / <i>Dolazni pozivi u roamingu<i/></td>
				<td width="3%">&nbsp;Kn.</td>
				<td align="right"><%=formatnumber(PreviousBalance_,-1) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>Outgoing Calls in Roaming</b> / <i>Odlazni pozivi u roamingu<i/></td>
				<td>&nbsp;Kn.</td>
				<td align="right"><%=formatnumber(Adjustment_,-1) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>GPRS/EDGE/UMTS Data Transfer</b> / <i>GPRS/EDGE/UMTS prijenos podataka<i/></td>
				<td>&nbsp;Kn.</td>
				<td align="right"><%=formatnumber(IPHONE_,-1) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td colspan="4">
			<table cellspadding="0" cellspacing="0" bordercolor="black" width="100%" bgColor="white">  
			<tr>
				<td>&nbsp;<b>Neto Total</b> / <i>Neto Total<i/></td>
				<td>&nbsp;Kn.</td>
				<td align="right"><%=formatnumber(Payment_,-1) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>VAT</b> / <i>PDV<i/></td>
				<td>&nbsp;Kn.</td>
				<td align="right"><%=formatnumber(PPN_,-1) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>Services Exempted from VAT</b> / <i>Usluge na koje se ne obracunava PDV<i/></td>
				<td>&nbsp;Kn.</td>
				<td align="right"><%=formatnumber(StampFee_,-1) %>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;<b>Grand Total</b> / <i>Bruto Total<i/></td>
				<td>&nbsp;Kn.</td>
				<td align="right"><u><b><%=formatnumber(Total_,-1) %></b></u>&nbsp;&nbsp;&nbsp;</td>
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
				<td width="3%">&nbsp;Kn.</td>
				<td align="right"><u><b><%=formatnumber(Total_,-1) %></b></u>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			</table>
			</td>
		</tr>
-->
		</table>
	</td>	
  </tr>
  </table>

		<form method="post" name="frmSearch" Action="CellPhoneDetail.asp?CellPhone=<%=CellPhone_%>&MonthP=<%=MonthP_%>&YearP=<%=YearP_%>">
		<table width="90%">
		<tr>
			<td colspan="3" align="center" class="SubTitle">Usage Detail:</td>
		</tr>
		<tr>				
			<td align="right"><b>Sort By</b>&nbsp;</td>
			<td>:</td>
			<td>
				<Select name="SortList">
					<Option value="DialedDatetime" <%if SortBy_ ="DialedDatetime" then %>Selected<%End If%> >Dialed Datetime</Option>
					<Option value="DialedNumber" <%if SortBy_ ="DialedNumber" then %>Selected<%End If%> >Dialed Number</Option>
					<Option value="CallType" <%if SortBy_ ="CallType" then %>Selected<%End If%> >Call Type</Option>
				</Select>&nbsp;
				<Select name="OrderList">
					<Option value="Asc" <%if Order_ ="Asc" then %>Selected<%End If%> >Asc</Option>
					<Option value="Desc" <%if Order_ ="Desc" then %>Selected<%End If%> >Desc</Option>
				</Select>
				<input type="submit" name="Submit" value="Refresh" />
			</td>
		</tr>			
		</table>
		</form>
		<form method="post" action="CellPhoneDetailSave.asp" name="frmCellPhoneBilling"> 
		<table cellspadding="0" cellspacing="0" bordercolor="#EEEEEE" border="1" width="90%" bgColor="white">  
		<tr bgcolor="#330099" align="center" cellpadding="0" cellspacing="0" >
			<TD width="5px"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
		       	<TD><strong><label STYLE=color:#FFFFFF>Dialed Date/time</label></strong></TD>
			<TD width="10%"><strong><label STYLE=color:#FFFFFF>Dialed Number</label></strong></TD>
			<TD><strong><label STYLE=color:#FFFFFF>Call Type</label></strong></TD>
			<TD><strong><label STYLE=color:#FFFFFF>Duration</label></strong></TD>
			<TD width="10%"><strong><label STYLE=color:#FFFFFF>Amount (Kn.)</label></strong></TD>
			<TD><strong><label STYLE=color:#FFFFFF>Check if<br>personal</label></strong><br>
<%			if (ProgressID_ = 1) or (ProgressID_ = 3) then %>
				<input type="checkbox" name="cbAll" value="true" onclick="checkall(this)" />
<%			end if %>
			</TD>
		</tr>
		<%
		'if (cdbl(ProgressID_)<4 or cdbl(ProgressID_ = 8)) then
		'	strsql = "Update CellPhoneDt Set isPersonal='N' Where CallType Like '%" & ExemptedCallType_  & "%' and PhoneNumber='" & CellPhone_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "' "
       		'	'response.write strsql 
       		'	BillingCon.execute strsql
		'end if


		strsql = "Select * from CellPhoneDt Where PhoneNumber='" & CellPhone_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "' Order by " & SortBy_ & " " & Order_ 
		'response.write strsql & "<br>"
		set rsCellPhone = BillingCon.execute(strsql) 

		no_ = 1 
		do while not rsCellPhone.eof
   			if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
'			if (ProgressID_ = 4) then
			if (cdbl(rsCellPhone("Cost")) <> cdbl(DetailRecordAmount_ )) then
		%>			
			<tr bgcolor="<%=bg%>">
				<td align="right"><%=No_%>&nbsp;</td>
			        <td>&nbsp;<%=rsCellPhone("DialedDatetime")%></font></td> 
		        	<td>&nbsp;<%=rsCellPhone("DialedNumber")%></font></td>
		        	<td>&nbsp;<%=rsCellPhone("CallType")%></font></td>
		        	<td>&nbsp;<%=rsCellPhone("CallDuration")%></font></td>
			        <td align="right"><%=formatnumber(rsCellPhone("Cost"),-1)%></font></td> 
<%'			if (ProgressID_ = 1) or (ProgressID_ = 3) then %>
<%'			if (cdbl(ProgressID_)<4 or (ProgressID_ = 4 and AlternateEmailFlag_="Y")) then %>

<%			if (((cdbl(ProgressID_) < 4 or cdbl(ProgressID_) = 8) and (InStr(1,rsCellPhone("CallType"),ExemptedIfOfficialCallType_,1) = 0 and InStr(1,rsCellPhone("CallType"),AlwaysExemptedCallType_,1) = 0))) then %>
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

			</tr>
		<%      end if
			rsCellPhone.movenext
			no_ = no_ + 1
		loop
		%>
		</table>
		<table cellspadding="0" cellspacing="0" bordercolor="black" width="90%" bgColor="white">  
		<tr>
			<td align="right" colspan="3"><b>Sub Total (Kn.) </b>&nbsp;</td>
			<td width="15%" class="FontContent" align="right"><b><%=formatnumber(TotalCellPhoneBillRp_ ,-1)%></b>&nbsp;</td>
			<td width="10%" class="FontContent" align="right"><b><u><%=formatnumber(TotalCellPhonePrsBillRp_ ,-1)%></u></b>&nbsp;</td>
		</tr>
  		<tr>
			<td colspan="5" align="center">&nbsp;</td>
		</tr>
<%'		if (ProgressID_ = 1) or (ProgressID_ = 3) then%>
<%	
			'response.write "No:" & no_
			if ((ProgressID_< 4 and no_ >1) or (ProgressID_ = 4 and AlternateEmailFlag_="Y")) then%>
		<tr>
			<td colspan="5" align="center">	
				<input type="submit" name="btnSubmit" Value="Update Change(s)" />&nbsp;&nbsp;
				<input type="button" value="Cancel" onClick="javascript:location.href='MonthlyBilling.asp?CellPhone=<%=CellPhone_%>&MonthP=<%=MonthP_%>&YearP=<%=YearP_%>'">

				<input type="hidden" name="txtCellPhone" value='<%=CellPhone_ %>' />
				<input type="hidden" name="txtMonthP" value='<%=MonthP_%>' />
				<input type="hidden" name="txtYearP" value='<%=YearP_%>' />
				<input type="hidden" name="txtEmpID" value='<%=EmpID_ %>' />
			<td>
		</tr>
		<%else%>
		<tr>
			<td colspan="5" align="center">	
			<input type="button" value="Back" onClick="javascript:location.href='MonthlyBilling.asp?CellPhone=<%=CellPhone_%>&MonthP=<%=MonthP_%>&YearP=<%=YearP_%>'">
			<td>
		</tr>
		<%end if%>

		</table>

		</form>
<%else%>
<table width="100%">
<tr>
	<td align="center">&nbsp;</td>	
</tr>
<tr>
	<td align="center">there is not data.</td>	
</tr>
</table>
<%end if%>

</BODY>
</html>
