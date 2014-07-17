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
  	<TD COLSPAN="4" ALIGN="center" Class="title">Billing Notification</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
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
TotalBillingRp_ = 0
TotalBillingDlr_ = 0

'response.write "Test : " & Request("cbApproval")
   If Request("cbApproval") <> "" then

	Set fso = CreateObject("Scripting.FileSystemObject")

	Dim send_from, send_to, send_cc, noMail, fileName
	send_from = BillingDL

	fileName = "Files\BillingDetail.xls"

	Dim ObjMail
	noMail=0
	For Each loopIndex in Request("cbApproval")
		'response.write loopIndex & "<br>"
		X = len(loopIndex)
		'response.write X & "<br>"
		EmpID_ = Left(loopIndex, X-7)
		'response.write EmpID_ & "<br>"
		Period = mid(loopIndex,X-6,6)
		MonthP_ = left(Period,2)
		'response.write MonthP_ & "<br>"
		YearP_ = Right(Period,4)
		BillType_ = Right(loopIndex,1)
		'response.write YearP_ & "<br>"
		strsql = "Select * From vwMonthlyBilling Where EmpID='" & EmpID_ & "' And MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "'"
		'response.write BillType_ & "<Br>"  
		'response.write strsql & "<Br>"  
		set rsData = server.createobject("adodb.recordset") 
		set rsData = BillingCon.execute(strsql)
		Period_ = MonthP_ & " - " & YearP_
		if not rsData.eof then
			EmpName_ = rsData("EmpName")
			Office_ = rsData("Agency") & " - " & rsData("Office")
			Position_ = rsData("WorkingTitle")
			OfficePhone_ = rsData("WorkPhone")
			HomePhone_ = rsData("HomePhone")
			MobilePhone_ = rsData("MobilePhone")
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
			EmpEmail_ = rsData("EmailAddress")
			TotalBillingAmountPrsRp_ = rsData("TotalBillingAmountPrsRp")
			TotalBillingAmountPrsDlr_ = rsData("TotalBillingAmountPrsDlr")
			AlternateEmailFlag_ = rsData("AlternateEmailFlag")
			DummyFlag_ = rsData("DummyFlag")
			ProgressID_ = rsData("ProgressID")
		end if

		if EmpEmail_ <>"" Then

			Set ObjMail = Server.CreateObject("CDO.Message")
			Set objConfig = CreateObject("CDO.Configuration") 
			objConfig.Fields(cdoSendUsingMethod) = 2  
'			objConfig.Fields(cdoSMTPServer) = "10.4.16.170"
			objConfig.Fields(cdoSMTPServer) = SMTPServer
'			objConfig.Fields(cdoSMTPServer) = "JAKARTAEX01.eap.state.sbu"
			objConfig.Fields.Update 
			Set objMail.Configuration = objConfig 
			'Send mail
			send_to = EmpEmail_ 
			'send_to = "kurniawane@state.gov"
			'response.write send_to
			objMail.From = send_from
			objMail.To = send_to 	
'			if ProgressID_ ="7" and EmpID_ <> "2490L" then
			if ProgressID_ ="7" then
				objMail.Subject = "Info: eBilling System - No Invoice This Period"
				objMail.HTMLBody = "<html><head>"
				ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
			
					& " <title>e-Billing Application</title> "_              
					& " <meta name='Microsoft Border' content='none, default'><style type='text/css'><!--.FontContent {font-size: 12px;color: blue;}--></style> "_     
					& " </head><body bgcolor='#ffffff'> "_              
					& " <p><table cellspadding='1' cellspacing='0' width='80%' bgColor='white'>"_    
					& "    <tr> "_           
					& "        <td colspan='6'>Your invoice has been processed and your usage did not meet the threshold to require review.  You have no further action.</td></tr> "_
					& "    <tr> "_           
					& "        <td colspan='6'><br></td></tr> "_
					& "    <tr> "_           
					& "        <td colspan='6'>Please reply if the phone number assignment below is not accurate.</td></tr> "_
					& "    <tr> "_           
					& "        <td colspan='6'>&nbsp;</td></tr> "_
					& "    <tr> "_           
					& "        <td colspan='6' align='Left'><b><u>Billing Data:</u></b></td></tr> "_
					& "    <tr> "_           
					& "        <td>Employee Name </td><td width='1%'>:</td><td class='FontContent' colspan='4'>" & EmpName_ & "</td></tr> "_
					& "    <tr> "_           
					& "        <td>Office Location </td><td width='1%'>:</td><td class='FontContent' colspan='4'>" & Office_ & "</td></tr> "_
					& "    <tr> "_ 
					& "        <td>Mobile Phone No </td><td width='1%'>:</td><td class='FontContent' colspan='4'>" & MobilePhone_ & "</td></tr> "_
					& "    <tr> "_ 
					& "        <td>Period </td><td width='1%'>:</td><td class='FontContent' colspan='4'>" & Period_ & "</td></tr> "_
					& "    <tr> "_ 
					& "        <td>Total Usage </td><td width='1%'>:</td><td class='FontContent' colspan='4'>Kn. " & formatnumber(TotalBillingAmountPrsRp_,2) & " <b>(no payment necessary)</b></td></tr> "_
					& "    <tr> "_ 
					& "        <td>Total Billing Amount </td><td width='1%'>:</td><td class='FontContent' colspan='4'>Kn. " & formatnumber(TotalBillingRp_,2) & "</td></tr> "_
					& "    <tr> "_       
					& "        <td colspan='6'>&nbsp; </td></tr> "_      
	
					& " </table></p>"_ 
		
					& "</body></html>"
			else
				objMail.Subject = "Action Required: eBilling System – Monthly Billing Reminder"
'				objMail.Subject = "e-Billing System - Monthly Billing Reminder for period " & Period_
				objMail.HTMLBody = "<html><head>"

				if AlternateEmailFlag_ ="N" and DummyFlag_="N" Then	
					ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
			
					& " <title>e-Billing Application</title> "_              
					& " <meta name='Microsoft Border' content='none, default'><style type='text/css'><!--.FontContent {font-size: 12px;color: blue;}--></style> "_     
					& " </head><body bgcolor='#ffffff'> "_              
					& " <p><table cellspadding='1' cellspacing='0' width='80%' bgColor='white'>"_    
					& "    <tr> "_           
					& "        <td colspan='6'>Your invoice has been processed for this billing period and action is required.</td></tr> "_
					& "    <tr> "_           
					& "        <td colspan='6'><b>Do NOT</b> make a payment yet - Please follow the instructions below:</td></tr> "_
					& "    <tr> "_           
					& "        <td colspan='6'>1) Click <a href='"& WebSiteAddress & "/MonthlyBilling.asp?Month="& MonthP_ & "&Year=" & YearP_ &"' target='_blank'>here </a> to access the ebilling application.</td></tr> "_
					& "    <tr> "_           
					& "        <td colspan='6'>2) In the application, click on the CellPhone hyper link.</td></tr> "_
					& "    <tr> "_           
					& "        <td colspan='6'>3) Uncheck any official calls.</td></tr> "_
					& "    <tr> "_           
					& "        <td colspan='6'>4) Click ""update"" to subtotal remaining personal calls.</td></tr> "_
					& "    <tr> "_           
					& "        <td colspan='6'>5) Submit your invoice to your supervisor for approval.</td></tr> "_
					& "    <tr> "_           
					& "        <td colspan='6'>6) Make payment if necessary - only <b>AFTER</b> your supervisor has approved.  You will receive a confirmation email informing you if you need to make a payment.</td></tr> "_
					& "    <tr> "_           
					& "        <td colspan='6' align='center'>&nbsp;</td></tr> "_
					& "    <tr> "_           
					& "        <td colspan='6' align='center'><u>Billing Period (Month - Year) : <a class='FontContent'>" & Period_ & "</a></u></td></tr> "_
					& "    <tr> "_           
					& "        <td colspan='6' align='Left'><u><b>Personal Info<b></u></td></tr> "_
					& "    <tr> "_           
					& "        <td width='20%'>Employee Name</td><td width='1%'>:</td><td class='FontContent'>" & EmpName_ & "</td><td>Agency / Office</td><td width='1%'>:</td><td class='FontContent'>" & Office_ & "</td></tr> "_
					& "    <tr> "_           
					& "        <td>Position</td><td width='1%'>:</td><td class='FontContent'>" & Position_ & "</td><td>Office Phone/Ext.</td><td width='1%'>:</td><td class='FontContent'>" & OfficePhone_ & "</td></tr> "_
					& "    <tr> "_           
					& "        <td>Homephone</td><td width='1%'>:</td><td class='FontContent' colspan='4'>" & HomePhone_ & "</td></tr> "_
					& "    <tr> "_ 
					& "        <td>Mobile Phone</td><td width='1%'>:</td><td class='FontContent'>" & MobilePhone_ & "</td><td>Exchange Rate</td><td width='1%'>:</td><td class='FontContent'>Kn." & FormatNumber(ExchangeRate_,2) & "/ Dollar</td></tr> "_
					& "    <tr> "_ 
					& "        <td colspan='6'><hr></td></tr> "_
					& "    <tr> "_ 
					& "        <td align='Left' colspan='6'><u><b>Billing Detail:<b></u></td></tr> "_
					& "    <tr> "_ 
					& "        <td align='Left' colspan='6'><table cellspadding='1' border='1' bordercolor='black' cellspacing='0' width='100%' bgColor='white'><tr align='center'><td rowspan='2'><b>Type</b></td><td rowspan='2'><b>Billing (Kn.)</b></td><td colspan='2'><b>Personal Usage</b></td></tr>"_
					& "    <tr> "_ 
					& "        <td align='center'><b>In Kuna (Kn.)</b></td><td align='center'><b>In US Dollar ($)</b></td></tr> "
					if (cdbl(OfficePhoneBillRp_) > 0) And ((BillType_ ="X") or (BillType_ ="O")) Then 
						ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
						& "<tr>"_ 
						& " 	<td>Office Phone</td>"_
						& "	<td width='20%' class='FontContent' align='right'>" & formatnumber(OfficePhoneBillRp_,2) & "&nbsp;</td>"_
						& "	<td width='20%' class='FontContent' align='right'>" & formatnumber(OfficePhonePrsBillRp_ ,2) & "&nbsp;</td> "_
						& "	<td width='20%' class='FontContent' align='right'>" & formatnumber(OfficePhonePrsBillDlr_,2) & "&nbsp;</td> "_		
						& " </tr>"
					elseIf (BillType_ ="X") then
						ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
						& " <tr>"_
						& " 	<td>Office Phone</td>"_
						& " 	<td class='FontContent' align='right'>- &nbsp;</td>"_
						& " 	<td class='FontContent' align='right'>- &nbsp;</td>"_
						& " 	<td class='FontContent' align='right'>- &nbsp;</td>"_
						& " </tr>"
					end if
					if (cdbl(HomePhoneBillRp_) > 0) And ((BillType_ ="X") or (BillType_ ="H")) Then 
						ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
						& " <tr>"_
						& " <td>Home Phone</td>"_
						& " <td class='FontContent' align='right'>" & formatnumber(HomePhoneBillRp_ ,2) & "&nbsp;</td> "_
						& " <td class='FontContent' align='right'>" & formatnumber(HomePhonePrsBillRp_ ,2) & "&nbsp;</td> "_
						& " <td class='FontContent' align='right'>" & formatnumber(HomePhonePrsBillDlr_ ,2) & "&nbsp;</td> "_
						& " </tr>"
					elseIf (BillType_ ="X") then
						ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
						& " <tr>"_
						& " <td>Home Phone</td>"_
						& " <td class='FontContent' align='right'>- &nbsp;</td>"_
						& " <td class='FontContent' align='right'>- &nbsp;</td>"_
						& " <td class='FontContent' align='right'>- &nbsp;</td>"_
						& " </tr>"
					end if	
					if (cdbl(CellPhoneBillRp_ ) > 0) And ((BillType_ ="X") or (BillType_ ="C")) Then 
						ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
						& " <tr>"_
							& " <td>Mobile Phone</td>"_
							& " <td class='FontContent' align='right'>"& formatnumber(CellPhoneBillRp_  ,2) & "&nbsp;</td>"_
							& " <td class='FontContent' align='right'>"& formatnumber(CellPhonePrsBillRp_ ,2) & "&nbsp;</td>"_
							& " <td class='FontContent' align='right'>"& formatnumber(CellPhonePrsBillDlr_ ,2) & "&nbsp;</td>"_
						& " </tr>"
					elseIf (BillType_ ="X") then
						ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
						& " <tr>"_
							& " <td>Mobile Phone</td>"_
							& " <td class='FontContent' align='right'>- &nbsp;</td>"_
							& " <td class='FontContent' align='right'>- &nbsp;</td>"_
							& " <td class='FontContent' align='right'>- &nbsp;</td>"_
						& " </tr>"
					end if		
					if (cdbl(TotalShuttleBillRp_) > 0) And ((BillType_ ="X") or (BillType_ ="S"))  Then
						ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
						& " <tr>"_
							& " <td>Shuttle Bus</td>"_
							& " <td class='FontContent' align='right'>"& formatnumber(TotalShuttleBillRp_ ,2) &"&nbsp;</td>"_
							& " <td class='FontContent' align='right'>"& formatnumber(TotalShuttleBillRp_ ,2) &"&nbsp;</td>"_
							& " <td class='FontContent' align='right'>"& formatnumber(TotalShuttleBillDlr_,2) &"&nbsp;</td>"_
						& " </tr>"
					elseIf (BillType_ ="X") then
						ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
						& " <tr>"_
							& " <td>Shuttle Bus</td>"_
							& " <td class='FontContent' align='right'>- &nbsp;</td>"_
							& " <td class='FontContent' align='right'>- &nbsp;</td>"_
							& " <td class='FontContent' align='right'>- &nbsp;</td>"_
						& " </tr>"
					end if
		
					if (BillType_ ="X") Then
					ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
						& " <tr>"_
							& " <td align='center'>Total</td>"_
							& " <td class='FontContent' align='right'><b><u>" & formatnumber(TotalBillingRp_ , 0) & "</u></b>&nbsp;</td>"_
							& " <td class='FontContent' align='right'><b><u>" & formatnumber(TotalBillingAmountPrsRp_ , 2) & "</u></b>&nbsp;</td>"_
							& " <td class='FontContent' align='right'><b><u>" & formatnumber(TotalBillingAmountPrsDlr_ ,2) & "</u></b>&nbsp;</td>"_
						& " </tr>"
					end if
	
					ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
					& "        </table></td></tr> "
		
			'		if (BillType_ ="X") Then
		'			ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
		'			& "    <tr> "_ 
		'			& "        <td colspan='6'><table cellspadding='1' cellspacing='0' width='100%' bgColor='white' border='0'><tr><td align='center' width='40%'><b>Total</b></td><td width='20%' class='FontContent' align='right'>&nbsp;</td><td width='20%' class='FontContent' align='right'><b><u>" & formatnumber(TotalBillingRp_ , 0) & "</u></b>&nbsp;</td><td width='20%' class='FontContent' align='right'><b><u>" & formatnumber(TotalBillingDlr_ ,2) & "</u></b>&nbsp;</td></tr></table></td></tr> "
		'			end if

					ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
					& "    <tr> "_       
					& "        <td colspan='6'>&nbsp; </td></tr> "_      
	
					& "    <tr> "_       
					& "        <td colspan='6'> "_
					& "        <p>You must take action. Click <a href='" & WebSiteAddress & "/MonthlyBilling.asp?Month="& MonthP_ & "&Year=" & YearP_ &"' target='_blank'>Here </a> to review your invoice.</p></td> "_  

					& "    <tr> "_       
					& "        <td colspan='6'>&nbsp; </td></tr> "_      
			
					& " </table></p>"_ 
		
					& "</body></html>"
	
				else
					If fso.FileExists (fileName) THEN
						set objFile = fso.GetFile (fileName)
						objFile.Delete
					end If 	
	
					Set objFile = fso.CreateTextFile(Server.MapPath(fileName))
	
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
	
								'strsql = "Select DetailRecordAmount From PaymentDueDate"
								strsql = "Select * From PaymentDueDate"
								'response.write strsql & "<br>"
								set rsDetailRecord = server.createobject("adodb.recordset") 
								set rsDetailRecord = BillingCon.execute(strsql)
								CeilingAmount_ = rsDetailRecord("CeilingAmount")
								if not rsDetailRecord.eof then
									DetailRecordAmount_ = rsDetailRecord("DetailRecordAmount")
									'response.write "DetailRecordAmount :" & DetailRecordAmount_ 
								end if
								'strsql = "Select * from CellPhoneDt Where PhoneNumber='" & MobilePhone_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "' and cost>" & DetailRecordAmount_ 
								strsql = "Select * from CellPhoneDt Where PhoneNumber='" & MobilePhone_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "' and cost>'" & DetailRecordAmount_ & "' Order by DialedDatetime Asc" 
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
					objFile.Writeline "     			<td align='right'><FONT color=#330099 size=2>" & formatnumber(rsCellPhone("Cost"),-1) & "</font></td>"
					objFile.Writeline "			</tr>"

								rsCellPhone.movenext
								No_ = No_ + 1
								loop
					objFile.Writeline "     		<tr>"
					objFile.Writeline "     			<td align='center' colspan='5'><b>Total (Kn.) </b>&nbsp;</td>"
					objFile.Writeline "				<td align='right'><b><u>" & formatnumber(CellPhonePrsBillRp_ ,-1) & "</u></b>&nbsp;</td>"
					objFile.Writeline "			</tr>"
					objFile.Writeline "			</table>"
					objFile.Writeline "		</td>"
					objFile.Writeline "	</tr>"
					objFile.Writeline "	</table>"
					objFile.Writeline "</form>"
					objFile.Writeline "</BODY>"
					objFile.Writeline "</HTML>"
					objFile.close
	
					'response.write MobilePhone_ & MonthP_ & YearP_ 
					strsql = "Select * From vwCellphoneHd Where PhoneNumber='" & MobilePhone_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "'"
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
					'response.write Total_
					ObjMail.HTMLBody = ObjMail.HTMLBody & " "_			
					& " <title>e-Billing Application</title> "_              
					& " <meta name='Microsoft Border' content='none, default'><style type='text/css'><!--.FontContent {font-size: 12px;color: blue;}--></style> "_     
					& " </head><body bgcolor='#ffffff'> "_              
					& " <p><table cellspadding='1' cellspacing='0' width='80%' bgColor='white'>"_    
					& "    <tr> "_       
					& "        <td colspan='6'>You have received this email because you have been identified as the supervisor for the phone number below, which is assigned to a group, an employee without open-net access, or an employee who is responsible for multiple phones.  Please follow the instructions :</td> "_
					& "    </tr> "_ 
					& "    <tr> "_
					& "        <td colspan='6'>1) Review the summary below and the Usage Detail in the attached MS Excel file.</td> "_
					& "    </tr> "_ 
					& "    <tr> "_ 
					& "        <td colspan='6'>2) Work with the users of the phone to determine if any of the calls are personal.</td> "_
					& "    </tr> "_ 				
					& "    <tr> "_
					& "        <td colspan='6'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;a.If personal calls amount to <b>less than or equal to " & formatnumber(CeilingAmount_,-1) & " kuna</b>, reply to this email and write “No Payment”.</td> "_
					& "    </tr> "_ 				
					& "    <tr> "_
					& "        <td colspan='6'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;in the email content.</td> "_
					& "    </tr> "_ 				
					& "    <tr> "_
					& "        <td colspan='6'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;b.If the personal calls amount to <b>more than " & formatnumber(CeilingAmount_,-1) & " kuna</b>, print this email, write in the personal call amount</td> "_ 
					& "    </tr> "_ 
					& "    <tr> "_
					& "        <td colspan='6'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;at the bottom, sign at the bottom, and instruct the employee to make payment with the cashier.</td> "_ 
					& "    </tr> "_ 
					& "    <tr> "_       
					& "        <td colspan='6'>&nbsp; </td></tr> "_    
					& "    <tr> "_           
					& "        <td colspan='6' align='center'><u>Billing Period (Month - Year) : <a class='FontContent'>" & Period_ & "</a></u></td></tr> "_
					& "    <tr> "_           
					& "        <td colspan='6' align='Left'><u><b>Personal Info<b></u></td></tr> "_
					& "    <tr> "_           
					& "        <td width='20%'>Employee Name</td><td width='1%'>:</td><td class='FontContent'>" & EmpName_ & "</td><td>Agency / Office</td><td width='1%'>:</td><td class='FontContent'>" & Office_ & "</td></tr> "_
					& "    <tr> "_           
					& "        <td>Position</td><td width='1%'>:</td><td class='FontContent'>" & Position_ & "</td><td>Office Phone/Ext.</td><td width='1%'>:</td><td class='FontContent'>" & OfficePhone_ & "</td></tr> "_
					& "    <tr> "_           
					& "        <td>Homephone</td><td width='1%'>:</td><td class='FontContent' colspan='4'>" & HomePhone_ & "</td></tr> "_
					& "    <tr> "_ 
					& "        <td>Mobile Phone</td><td width='1%'>:</td><td class='FontContent'>" & MobilePhone_ & "</td><td>Exchange Rate</td><td width='1%'>:</td><td class='FontContent'>Kn." & FormatNumber(ExchangeRate_,2) & " / Dollar</td></tr> "_
					& "    <tr> "_ 
					& "        <td colspan='6'><hr></td></tr> "_
					& "    <tr> "_ 
					& "        <td align='Left' colspan='6'><u><b>Billing Detail :<b></u></td></tr> "_
					& "    <tr> "_ 
					& "        <td align='Left' colspan='6'> "_
					& "		<table cellspadding='0' border='1' bordercolor='black' cellspacing='0' width='100%' bgColor='white'> "_
					& "		<tr><td colspan='4' align='center' class='SubTitle'>USAGE SUMMARY</td></tr> "_
					& "		<tr><td>&nbsp;<b>Previous Balance</b> / <i>Tagihan Sebelumnya</i><div align='center'>Rp.&nbsp;&nbsp;"& formatnumber(PreviousBalance_,0) &" </div></td>" _
					& "		    <td>&nbsp;<b>Payment</b> / <i>Pembayaran</i><div align='center'>Rp.&nbsp;&nbsp;" & formatnumber(Payment_,0) &"</div></td>"_
					& "		    <td>&nbsp;<b>Adjustment</b> / <i>Koreksi</i><div align='center'>Rp.&nbsp;&nbsp;" & formatnumber(Adjustment_,0) &"</div></td>"_
					& "		    <td>&nbsp;<b>Balance Due</b> / <i>Sisa Tagihan</i><div align='center'>Rp.&nbsp;&nbsp;" & formatnumber(BalanceDue_,0) &"</div></td>"_
					& "		</tr>"_
					& "		<tr><td colspan='4'>&nbsp;<u><b>Usage Charges</b> / <i>Biaya Percakapan:<i/></u></td></tr>"_
					& "		<tr><td colspan='4'><table cellspadding='0' cellspacing='0' bordercolor='black' width='100%' bgColor='white'>"_
					& "		    <tr><td width='70%'>&nbsp;<b>Subscription Fee</b> / <i>Abonemen<i/></td><td width='3%'>&nbsp;Rp.</td><td align='right'>" & formatnumber(SubscriptionFee_,0) & "&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<b>Local</b> / <i>Lokal<i/></td><td>&nbsp;Rp.</td><td align='right'>"& formatnumber(LocalCall_,0) & "&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<b>Interlocal</b> / <i>SLJJ<i/></td><td>&nbsp;Rp.</td><td align='right'>"& formatnumber(Interlocal_,0) & "&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<b>IDD</b> / <i>SLI<i/></td><td>&nbsp;Rp.</td><td align='right'>"& formatnumber(IDD_,0) &"&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<b>SMS</b> / <i>SMS<i/></td><td>&nbsp;Rp.</td><td align='right'>"& formatnumber(SMS_,0) &"&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<b>International Roaming Surcharge</b> / <i>Surcharge Jelajah Internasional<i/></td><td>&nbsp;Rp.</td><td align='right'>"& formatnumber(IRS_,0) & "&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<b>International Roaming Leg</b> / <i>Roaming Leg Jelajah Internasional<i/></td><td>&nbsp;Rp.</td><td align='right'>"& formatnumber(IRL_,0) & "&nbsp;&nbsp;&nbsp;</td></tr></table></td></tr>"_
					& "		    <tr><td colspan='4'>&nbsp;<u><b>Value Added Services</b> / <i>Layanan Tambahan:<i/></u></td></tr>"_
					& "		    <tr><td colspan='4'><table cellspadding='0' cellspacing='0' bordercolor='black' width='100%' bgColor='white'>"_
					& "		    <tr><td width='70%'>&nbsp;<b>Prepaid Recharge</b> / <i>Isi Ulang Prabayar<i/></td><td width='3%'>&nbsp;Rp.</td><td align='right'>"& formatnumber(Prepaid_,0) & "&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<b>Fax Response and Interactive Data</b> / <i>FARIDA<i/></td><td>&nbsp;Rp.</td><td align='right'>"& formatnumber(FARIDA_,0) & "&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<b>Mobile Banking</b> / <i>Mobile Banking<i/></td><td>&nbsp;Rp.</td><td align='right'>"& formatnumber(MobileBanking_,0) & "&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<b>Detailed Call Record Print</b> / <i>Print Rincian Percakapan<i/></td><td>&nbsp;Rp.</td><td align='right'>"& formatnumber(DetailedCallRecord_,0) & "&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<b>3G, HSDPA, GPRS, MMS, Wifi, Premium Content</b> / <i>3G, HSDPA, GPRS, MMS, Wifi, Konten Premium<i/></td><td>&nbsp;Rp.</td><td align='right'>"& formatnumber(GPRS_,0) &"&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<b>Ventus / Blackberry, iPhone, Bridge Dataroam, Data Roam</b> / <i>Ventus / Blackberry, iPhone, Bridge Dataroam, Data Roam<i/></td><td>&nbsp;Rp.</td><td align='right'>"& formatnumber(IPHONE_,0) &"&nbsp;&nbsp;&nbsp;</td></tr></table></td></tr>"_
					& "		    <tr><td colspan='4'><table cellspadding='0' cellspacing='0' bordercolor='black' width='100%' bgColor='white'>"_
					& "		    <tr><td width='70%'>&nbsp;<b>Variance To Minimum Usage Guarantee</b> / <i>Selisih Penggunaan Minimum<i/></td><td width='3%'>&nbsp;Rp.</td><td align='right'>"& formatnumber(MinUsage_,0) &"&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<b>Sub Total</b> / <i>Sub Total<i/></td><td>&nbsp;Rp.</td><td align='right'>"& formatnumber(SubTotal_,0) &"&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<b>VAT 10%</b> / <i>PPN 10%<i/></td><td>&nbsp;Rp.</td><td align='right'>"& formatnumber(PPN_,0) &"&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<b>Stamp Duty Fee</b> / <i>Biaya Materai pembayaran bulan lalu)<i/></td><td>&nbsp;Rp.</td><td align='right'>"& formatnumber(StampFee_,0)&"&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<b>Current Balance</b> / <i>Total Tagihan Bulan Ini<i/></td><td>&nbsp;Rp.</td><td align='right'><u><b>"& formatnumber(CurrentBalance_,0)&"</b></u>&nbsp;&nbsp;&nbsp;</td></tr></table></td></tr>"


					ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
					& "        </table></td></tr> "_	
					& "    <tr> "_       
					& "        <td colspan='6'>&nbsp; </td></tr> "_      
					& "    <tr> "_       
					& "        <td colspan='6'>&nbsp; </td></tr> "_      
	
					& "    <tr> "_       
					& "        <td colspan='6'>&nbsp;&nbsp;Amount to be paid for personal calls: _______________________  Supervisor Signature:________________________</td> "_ 
					& "    </tr> "_ 

					& " </table></p>"_ 
		
					& "</body></html>"

					ObjMail.AddAttachment Server.MapPath(fileName)
				end if
			end if

				objMail.Send

'			strsql = "Execute spSendNotificationUpdate '" & EmpID_ & "','" & MonthP_ & "','" & YearP_ & "','" & AlternateEmailFlag_ & "'"
			strsql = "Execute spSendNotificationUpdate '" & EmpID_ & "','" & MonthP_ & "','" & YearP_ & "'"
			'response.write strsql & "<Br>"  
			set rsData = server.createobject("adodb.recordset") 
			set rsData = BillingCon.execute(strsql)					
			
			noMail=noMail+1
		end if
	next
	Set objMail = Nothing 
	Set objConfig = Nothing 
  End If

'  response.redirect("BillingApprovalList.asp")
%>
<table cellspadding="1" cellspacing="0" width="100%" align="center">
<tr>
	<td><br></td>
</tr>
<tr>
	<td align="center"><%=noMail%> message(s) was/were sent.</td>
</tr>
<tr>
	<td>&nbsp;</td>
</tr>
<tr>
	<td align="center"><input type="button" value="Close" id="btnclose"></td>
</tr>
<tr>
	<td align="center"><br><a href="javascript:history.go(-1)"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
</tr>
</table>
</body> 
</html>