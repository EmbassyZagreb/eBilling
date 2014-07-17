<HTML>
<HEAD>
<!--#include file="connect.inc" -->
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
'response.write "Test : " & Request("cbApproval")
   If Request("cbApproval") <> "" then
	   strsql = "Select CeilingAmount From PaymentDueDate"
	   'response.write strsql & "<br>"
	   set rsData = server.createobject("adodb.recordset") 
	   set rsData = BillingCon.execute(strsql) 
	   if not rsData.eof then
		CeilingAmount_ = rsData("CeilingAmount")
	   else
		CeilingAmount_ = 0
	   end if
	For Each loopIndex in Request("cbApproval")
		'response.write loopIndex & "<br>"
		X = len(loopIndex)
		'response.write X & "<br>"
		EmpID_ = Left(loopIndex, X-6)
		'response.write EmpID_ & "<br>"
		Period = right(loopIndex,6)
		MonthP_ = left(Period,2)
		'response.write MonthP_ & "<br>"
		YearP_ = Right(Period,4)
		'response.write YearP_ & "<br>"
		'strsql = "Exec spBillingApprovalAll '" & EmpID_ & "','" & MonthP_ & "','" & YearP_ &"'"
'		response.write strsql & "<Br>"  
		'BillingCon.execute(strsql) 

		strsql = "Select * from vwMonthlyBilling Where EmpID='" & EmpID_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "'"
		'response.write strsql & "<br>"
		set rsData = server.createobject("adodb.recordset") 
		set rsData = BillingCon.execute(strsql) 
		'response.write Period_  & "<br>"
		if not rsData.eof then
			EmpName_ = rsData("EmpName")
			MonthP_ = rsData("MonthP")
			YearP_ = rsData("YearP")
			EmpEmail_ = rsData("EmailAddress")
			Office_ = rsData("Agency") & " - " & rsData("Office")

			Position_ = rsData("WorkingTitle")
			OfficePhone_ = rsData("WorkPhone")
			HomePhone_ = rsData("HomePhone")
			MobilePhone_ = rsData("MobilePhone")	
			EmpEmail_ = rsData("EmailAddress")
			LoginID_ = rsData("LoginID")
			FiscalData_ = rsData("FiscalStripVAT")

			Extension_ = rsData("WorkPhone")
			TotalBillingPrsRp_ = rsData("TotalBillingAmountPrsRp")
			'TotalBillingRp_ = rsData("TotalBillingRp")
			Period_ = MonthP_ & " - " & YearP_
		end if

		if cdbl(TotalBillingPrsRp_) < cdbl(CeilingAmount_ ) Then
			ProgressId_ = 7
		Else
			ProgressId_ = 4
		End if

		if EmpEmail_ <>"" then		
			'response.write "send mail"
			Dim send_from, send_to, send_cc, send_bcc

			send_from = BillingDL
			send_to = EmpEmail_

			Dim ObjMail
			Set ObjMail = Server.CreateObject("CDO.Message")
			Set objConfig = CreateObject("CDO.Configuration") 
			objConfig.Fields(cdoSendUsingMethod) = 2  
			objConfig.Fields(cdoSMTPServer) = SMTPServer
			objConfig.Fields.Update 
			Set objMail.Configuration = objConfig 
			objMail.From = send_from
			objMail.To = send_to 
			'objMail.CC = send_cc

			objMail.Subject = "Info: eBilling System � Approval Notification"

			objMail.HTMLBody = "<html><head>"

			if (ProgressId_ = 7) then
				ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
			
				& " <title>e-Billing Application</title> "_              
				& " <meta name='Microsoft Border' content='none, default'> "_     
				& " </head><body bgcolor='#ffffff'> "_              
				& " <p><table border='0' width='100%'> "_    

				& "    <tr> "_       
				& "        <td> Your billing <b>has been APPROVED</b> by your supervisor.</td></tr> "_      
				& "    <tr> "_       
				& "        <td>&nbsp; </td></tr> "_        

				& "    <tr> "_
				& "        <td>Please click <A HREF='" & WebSiteAddress & "/MonthlyBillList.asp?sMonthP=" & MonthP_ & "&sYearP=" & YearP_ & "&eMonthP=" & MonthP_ & "&eYearP=" & YearP_ & "'>here</A> to review your billing.</td> "_ 

				& "    <tr> "_       
				& "        <td>&nbsp; </td></tr> "_ 
	
				& "    <tr> "_       
				& "        <td><h3><u>Billing Data</u> </h3></td></tr> "_      
		
				& "    <tr> "_       
				& "        <td>- Employee Name : <b>" & EmpName_  & "</b></td></tr> "_      
	
				& "    <tr> "_       
				& "        <td>- Office Location : <b>" & Office_ & "</b></td></tr> "_      
				& "    <tr> "_       
				& "        <td>- Phone Ext. : <b>" & Extension_ & "</b></td></tr> "_      

				& "    <tr> "_       
				& "        <td>- Period : <b>" & Period_ & "</b></td></tr> "_      

				& "    <tr> "_       
				& "        <td>- Total Personal Usage : <b>Kn. " & formatnumber(TotalBillingPrsRp_,-1) & "</b>.&nbsp;&nbsp;<font color='blue'><b>Your personal usage amount is less than the threshold amount, no payment is required.</b></font></td></tr> "_ 

				& "    <tr> "_       
				& "        <td>- Note  : <b>" & Notes_ & "</b></td></tr> "_  
	     
				& "    <tr> "_       
				& "        <td>&nbsp; </td></tr> "_      

				& "    <tr> "_       
				& "        <td>&nbsp; </td></tr> "_      

				& "    <tr> "_       
				& "        <td align='middle'>NOTE: This e-mail was automatically generated. Please do not respond to this e-mail address.</td> "_       
				& "    </tr> "_ 

				& " </table></p>"_ 
				& "</body></html>"
			else
				ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
				& " <title>e-Billing Application</title> "_              
			& " <meta name='Microsoft Border' content='none, default'><style type='text/css'><!--.FontContent {font-size: 12px;color: blue;}--></style> "_     
			& " </head><body bgcolor='#ffffff'><center> "_              
			& " <IMG SRC='"& WebSiteAddress & "/images/embassytitle2.jpeg' WIDTH='100%' HEIGHT='80' BORDER='0'><br>"_
			& " <p><table border='0' width='100%'> "_    
			& "    <tr> "_       
			& "        <td> Your billing <b>has been APPROVED</b> by your supervisor.</td></tr> "_      
			& "    <tr> "_       
			& "        <td>&nbsp; </td></tr></table></p>"_
			& " <p><table cellspadding='1' cellspacing='0' width='100%' bgColor='white'>"_ 
			& " <tr><td colspan='2' width='50%'><table cellspacing='0' border='1' bordercolor='black'>"_
						& " <tr><td colspan='2' align='Center'><b>Payee Info</b></td></tr>"_
						& " <tr><td align='right'><i>Employee Name : </i></td><td>" & EmpName_ & "</td></tr>"_
						& " <tr><td align='right'><i>Department : </i></td><td>" & Office_ & "</td></tr></table></td> "_
				& "<td colspan='4' width='50%' align='right'><table cellspacing='0'>"_
					& " <tr><td colspan='4'><b><font color='red' size='5'>Bill of Collection</font></td></tr>"_
					& " <tr><td colspan='2'><i>Bill of Collection Number : </i></td><td colspan='2'>" & LoginID_ & curMonth_ & curYear_ &"</td></tr>"_
					& " <tr><td colspan='2'><i>Bill of Collection Date : </i></td><td colspan='2'>" & Date() & "</td></tr>"_
			& " </table></td></tr>"_
			& "<tr><td colspan='6'>&nbsp;</td></tr>"_
			& "<tr><td colspan='6'>&nbsp;</td></tr>"_
			& "<tr><td colspan='6'><table cellspacing='0' width='100%' border='1' bordercolor='black'>"_
				& "<tr align='center'><td><b>Billing Period</b></td><td><b>Bill Type</b></td><td><b>Description</b></td><td><b>Amount Due (USD)</b></td><td><b>Exchange Rate</b></td><td><b>Amount Due (Kn)</b></td></tr>"
			TotalBillRp_ = 0
			TotalBillDlr_ = 0
			 Do while not rsData.eof
				TotalBillRp_ = cdbl(TotalBillRp_) + cdbl(rsData("TotalBillingAmountPrsRp"))
				TotalBillDlr_ = cdbl(TotalBillDlr_) + cdbl(rsData("TotalBillingAmountPrsDlr"))
				ObjMail.HTMLBody = ObjMail.HTMLBody & "<tr><td>&nbsp;" & rsData("YearP") & "-" & rsData("MonthP") & "</td><td>Mobile Phone</td><td>&nbsp;" & rsData("MobilePhone") & "</td><td align='right'>$ " & formatnumber(cdbl(rsData("TotalBillingAmountPrsDlr")),-1) & "&nbsp;</td><td align='right'>" & formatnumber(rsData("ExchangeRate"),-1) & "&nbsp;</td><td align='right'>" & formatnumber(cdbl(rsData("TotalBillingAmountPrsRp")),-1) & "&nbsp;</td></tr>"
			 	rsData.movenext
			 Loop 
			ObjMail.HTMLBody = ObjMail.HTMLBody & "<tr><td colspan='3' align='right'><b>Total</b>&nbsp;</td><td align='right'><b>$ " & formatnumber(TotalBillDlr_,-1) & "</b>&nbsp;</td><td>&nbsp;</td><td align='right'><b>" & formatnumber(TotalBillRp_,-1) & "</b>&nbsp;</td></tr></table></td></tr>"_			
			& " </table><div align='right'><i>*Please remit payment within 15 days of the Invoice Date</i></div>"_
			& "<br><br><br><div align='left'><b>Payment Options</b></div>"_
			& "<div align='left'>Pay at the embassy cashier:</div>"_
			& "<div align='left'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & CashierInfo & "</div>"_
			& "<br><br><br><div align='left'><b>Fund Cite Details (For Cashier)</b></div>"_
			& "<div align='left'>&nbsp;&nbsp;&nbsp;$ " & formatnumber(TotalBillDlr_,-1) & " 19-X45190001-00-5306-53063R3221-6195-2322 (ICASS Mobile Phone)</div>"_
			& "<br><br><br><br><br><div align='left'>Please contact <a href='mailto:zgbphonebill@state.gov'>zgbphonebill@state.gov</a> with any questions</div>"_
			& "</p></body></html>"

			end if

		'	response.write "send mail"
			objMail.Send 
			Set objMail = Nothing 
			Set objConfig = Nothing 
		End If
		
		'Update MonthlyBill
		strsql = "Update MonthlyBilling Set ProgressId=" & ProgressId_ & ", ProgressIdDate=GetDate(), SupervisorRemark='" & Remark_ & "', SupervisorApproveDate='" & Date() & "' Where EmpID='" & EmpID_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ &"'"
		BillingCon.execute(strsql)
		response.write strsql

	next
  End If

  response.redirect("BillingApprovalListAddOn.asp?ProgressID=2")
%>
</body> 
</html>