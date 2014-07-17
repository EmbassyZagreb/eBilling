<%@ Language=VBScript %>
<!--#include file="connect.inc" -->

<% 
CellPhone_ = request("CellPhone")
MonthP_ = request("MonthP")
YearP_ = request("YearP")

Set fso = CreateObject("Scripting.FileSystemObject")

fileName = "BillingDetail.xls"
If fso.FileExists (fileName) THEN
	set objFile = fso.GetFile (fileName)
	objFile.Delete
end If 

Set objFile = fso.CreateTextFile(Server.MapPath(fileName))

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

objFile.Writeline "<HTML>"
objFile.Writeline "<HEAD><TITLE>Billing</TITLE>"
objFile.Writeline "<style type='text/css'>"
objFile.Writeline "<!--"
objFile.Writeline ".style4 {color: #FFFFFF; font-weight: bold;}"
objFile.Writeline ".smallfont{font-size: x-small;}"
objFile.Writeline "-->"
objFile.Writeline "</style>"
objFile.Writeline "</HEAD>"
objFile.Writeline "<BODY>"
objFile.Writeline "<form>"
objFile.Writeline "   <table cellpadding='0' cellspacing='0' border='0' width='100%'>"
objFile.Writeline "      <tr>"
objFile.Writeline "		<td><b>Personal Usage Detail for Period <b>" & MonthP_ & " - " & YearP_ & "</b> :</b></td>"
objFile.Writeline "      </tr>"
objFile.Writeline "     <tr>"
objFile.Writeline "		<td>&nbsp;</td>"
objFile.Writeline "     </tr>"

objFile.Writeline "     <tr>"
objFile.Writeline "     	<td align='Center'>"
objFile.Writeline "     		<table cellspadding='0' cellspacing='0' bordercolor='black' border='1' width='90%' bgColor='white'>"
objFile.Writeline "     		<tr align='center' cellpadding='0' cellspacing='0'>"
objFile.Writeline "     			<TD width='5%'><strong>No</strong></TD>"
objFile.Writeline "     		     	<TD><strong>Dialed Date/time</strong></TD>"
objFile.Writeline "     			<TD width='20%'><strong>Dialed Number</strong></TD>"
objFile.Writeline "     			<TD><strong>Call Type</strong></TD>"
objFile.Writeline "     			<TD><strong>Duration</strong></TD>"
objFile.Writeline "     			<TD width='10%'><strong>Amount (Kn)</strong></TD>"
objFile.Writeline "     		</tr>"

				strsql = "Select * from CellPhoneDt Where PhoneNumber='" & CellPhone_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "'"
				'response.write strsql & "<br>"
				set rsCellPhone = BillingCon.execute(strsql) 
				No_ = 1 
				do while not rsCellPhone.eof
	   			if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4"
objFile.Writeline "     		<tr bgcolor='" & bg & "'>"
objFile.Writeline "     			<td align='right'>" & No_ & "&nbsp;</td>"
objFile.Writeline "     			<td><FONT color=#330099 size=2>&nbsp;" & rsCellPhone("DialedDatetime") & "</font></td>"
objFile.Writeline "     			<td><FONT color=#330099 size=2>&nbsp;" & rsCellPhone("DialedNumber") & "</font></td>"
objFile.Writeline "     			<td><FONT color=#330099 size=2>&nbsp;" & rsCellPhone("CallType") & "</font></td>"
objFile.Writeline "     			<td><FONT color=#330099 size=2>&nbsp;" & rsCellPhone("CallDuration") & "</font></td>"
objFile.Writeline "     			<td align='right'><FONT color=#330099 size=2>" & formatnumber(rsCellPhone("Cost"),-1) & "&nbsp;</font></td>"
objFile.Writeline "			</tr>"

				rsCellPhone.movenext
				No_ = No_ + 1
				loop
objFile.Writeline "     		<tr>"
objFile.Writeline "     			<td align='center' colspan='5'><b>Total (Kn.) </b>&nbsp;</td>"
objFile.Writeline "				<td align='right'><b><u>" & formatnumber(TotalCellPhonePrsBillRp_ ,-1) & "</u></b>&nbsp;</td>"
objFile.Writeline "			</tr>"
objFile.Writeline "			</table>"
objFile.Writeline "		</td>"
objFile.Writeline "	</tr>"
objFile.Writeline "	</table>"
objFile.Writeline "</form>"
objFile.Writeline "</BODY>"
objFile.Writeline "</HTML>"

%>