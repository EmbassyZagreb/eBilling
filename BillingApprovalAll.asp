<HTML>
<HEAD>
<!--#include file="connect.inc" -->
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 
<TITLE>U.S. Embassy Zagreb - zBilling Application</TITLE>
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
		'EmpID_ = Left(loopIndex, X-6)
		MobilePhone_ = Left(loopIndex, X-6)
		'response.write EmpID_ & "<br>"
		Period = right(loopIndex,6)
		MonthP_ = left(Period,2)
		'response.write MonthP_ & "<br>"
		YearP_ = Right(Period,4)
		'response.write YearP_ & "<br>"
		'strsql = "Exec spBillingApprovalAll '" & EmpID_ & "','" & MonthP_ & "','" & YearP_ &"'"
'		response.write strsql & "<Br>"  
		'BillingCon.execute(strsql) 

		strsql = "Select * from vwMonthlyBilling Where MobilePhone='" & MobilePhone_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "'"
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
			CellPhoneBillRp_ = rsData("CellPhoneBillRp")
			Position_ = rsData("WorkingTitle")
			'OfficePhone_ = rsData("WorkPhone")
			'HomePhone_ = rsData("HomePhone")
			'MobilePhone_ = rsData("MobilePhone")	
			EmpEmail_ = rsData("EmailAddress")
			LoginID_ = rsData("LoginID")
			FiscalDataVAT_ = rsData("FiscalStripVAT")
			FiscalDataNonVAT_ = rsData("FiscalStripNonVAT")
			'Extension_ = rsData("WorkPhone")
			TotalBillingPrsRp_ = rsData("TotalBillingAmountPrsRp")
			'TotalBillingRp_ = rsData("TotalBillingRp")
			Period_ = MonthP_ & " - " & YearP_
			'ProgressDesc_ = rsData("ProgressDesc")
			CellPhonePrsBillRp_ = rsData("CellPhonePrsBillRp")
		

		end if

		if (cdbl(TotalBillingPrsRp_) < cdbl(CeilingAmount_ )) or (cdbl(TotalBillingPrsRp_ )=0 ) Then
		'if (cdbl(TotalBillingAmount_) < cdbl(CeilingAmount_ )) or (cdbl(TotalBillingAmount_)=0 ) Then
			ProgressId_ = 7
		Else
			ProgressId_ = 4
		End if

		'Update MonthlyBill
		strsql = "Update MonthlyBilling Set ProgressId=" & ProgressId_ & ", ProgressIdDate=GetDate(), SupervisorRemark='" & Remark_ & "', SupervisorApproveDate='" & Date() & "' Where PhoneNumber='" & MobilePhone_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ &"'"
		BillingCon.execute(strsql)
		'response.write strsql

		strsql = "Select ProgressDesc from ProgressStatus Where ProgressId='" & ProgressId_ & "'"
		set rsProgress = server.createobject("adodb.recordset") 
		set rsProgress = BillingCon.execute(strsql) 
		if not rsProgress.eof then
			ProgressDesc_ = rsProgress("ProgressDesc")
		end if

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

			objMail.Subject = "Info: zBilling System � Approval Notification"


















		objMail.HTMLBody = "<html><head>"
		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
	
		& " <title>e-Billing Application</title> "_              
		& " <meta name='Microsoft Border' content='none, default'><style type='text/css'><!--.FontContent {font-family: verdana;font-size: 11px;color: black;}--></style> "_     
		& " </head><body bgcolor='#ffffff'> "_              
		& " <p><table cellspadding='1' cellspacing='0' width='80%' bgColor='white'>"_
		& "    <tr> "_           
		& "        <td colspan='6' align='center'><font face='Verdana, Arial, Helvetica' color='#999999' size='5'>zBilling Approval Notification</font></td></tr> "_
		& "    <tr> "_ 
		& "        <td colspan='6'>&nbsp; </td></tr> " _
		& "    <tr> "_           
		& "        <td colspan='6' align='Left' class='FontContent'>Your billing <strong>has been APPROVED</strong> by your supervisor.</td></tr> "_
		& "    <tr> "_       
		& "        <td colspan='6' align='Left' class='FontContent'><br></td></tr> "_    
		& "    <tr> "_           
		& "        <td colspan='6' align='Left' class='FontContent'>&nbsp;<u><strong>Personal Info:<strong></u></td></tr> "_
		& "    <tr> "_ 
		& "    <td colspan='6' align='Left'> "_
		& "    	<table cellspadding='1' border='2' bordercolor='black' cellspacing='3' width='100%' bgColor='#999999' border='0'>   "_
		& "    		<tr BGCOLOR='#999999'> "_
		& "    			<td colspan='3' style='border: none;' class='FontContent'><FONT color=#FFFFFF><strong>Employee Name : " & EmpName_ & "</strong></font></td> "_
		& "    			<td colspan='3' style='border: none;' align='right' class='FontContent'><FONT color=#FFFFFF><strong>Phone Number : " & MobilePhone_ & "&nbsp;</strong></font></td> "_
		& "    		</tr> "_
		& "    		<tr BGCOLOR='#999999'> "_
		& "    			<td colspan='6' style='border: none;' class='FontContent'><FONT color=#FFFFFF><strong>Position : " & Position_ & "</strong></font></td> "_
		& "    		</tr> "_
		& "    		<tr BGCOLOR='#999999'> "_
		& "    			<td colspan='6' style='border: none;' class='FontContent'><FONT color=#FFFFFF><strong>Agency / Office : " & Office_ & "</strong></font></td> "_
		& "    		</tr> "_
		& "    	</table></td></tr> "_
		& "    <tr> "_ 
		& "        <td align='Right' colspan='6'><font face='Verdana, Arial, Helvetica' color='#999999' size='4'><strong>Billing Summary&nbsp;<strong></font></td></tr> "_
		& "    <tr> "_ 
		& "        <td align='Left' colspan='6' class='FontContent'>&nbsp;<u><strong>Billing Detail:<strong></u></td></tr> "_
		& "    <tr> "_
		& "    <td align='Left' colspan='6'> "_
		& "    <table cellspadding='1' border='1' bordercolor='black' cellspacing='0' width='100%' bgColor='white'> "_  
		& "    	<tr align='center' height=26> "_
		& "    		<td width='25%' class='FontContent'><strong>Action</strong></td> "_
		& "    		<td width='25%' class='FontContent'><strong>Billing Period</strong></td> "_
		& "    	<!--	<td width='20%' class='FontContent'><strong>Status</strong></td> --> "_
		& "    		<td width='25%' class='FontContent'><strong>Total Bill</strong></td> "_
		& "    		<td width='25%' class='FontContent'><strong>Personal Amount Due</strong></td> "_
		& "    	</tr> "

		if cdbl(CellPhoneBillRp_ ) > 0 Then

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    	<tr height=26> "

		if ProgressID_ <= 3 Then

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    	<td class='FontContent'>&nbsp;<a href='" & WebSiteAddress & "/CellPhoneDetail.asp?CellPhone=" & MobilePhone_ & "&MonthP=" & MonthP_ & "&YearP=" & YearP_ & "' target='_blank'>Tick your calls</a></td> "

		else

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    	<td class='FontContent'>&nbsp;<a href='" & WebSiteAddress & "/CellPhoneDetail.asp?CellPhone=" & MobilePhone_ & "&MonthP=" & MonthP_ & "&YearP=" & YearP_ & "' target='_blank'>View Submitted Bill</a></td> "

		end if

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
	        & "    	<TD align='right' class='FontContent'>&nbsp;" & MonthP_ & "-" & YearP_ & "</font>&nbsp;</TD> "_
	        & "   <!-- 	<TD align='right' class='FontContent'>" & ProgressDesc_ & "&nbsp;</font></TD> --> "_
		& "    	<td align='right' class='FontContent'>" & formatnumber(CellPhoneBillRp_  ,-1) & "&nbsp;</td> "_
		& "    	<td align='right' class='FontContent'>" & formatnumber(CellPhonePrsBillRp_ ,-1) & "&nbsp;</td> "_
		& "    	</tr> "

		else

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr> "_
		& "    <td class='FontContent'>Mobile Phone</td> "_
		& "    <td class='FontContent' align='right'>- &nbsp;</td> "_
		& "    <td class='FontContent' align='right'>- &nbsp;</td> "_
		& "    <td class='FontContent' align='right'>- &nbsp;</td> "_
		& "    <td class='FontContent' align='right'>- &nbsp;</td> "_
		& "    </tr> "

		end if

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_

		& "    </table></td><tr> "_       
		& "        <td colspan='6'>&nbsp; </td></tr> " 


	if (ProgressId_ ="7") Then
     
		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr> "_  
		& "        <td class='FontContent' colspan='6'>Your personal usage amount is below the collection threshold amount. No payment is necessary for this bill.</td></tr> "  

	else
    
		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr> "_       
		& "        <td class='FontContent' colspan='6'> "_
		& "        You must take action. Click <A HREF='" & WebSiteAddress & "/MonthlyBilling.asp'>Here</A> to review your invoice.</td> "_  
		& "    <tr> "_       
		& "        <td colspan='6'>&nbsp; </td></tr> " 

 	end if

	if (ProgressId_ <> "7") Then

    		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr> "_       
		& "        <td class='FontContent' colspan='6'>&nbsp;<u><strong>Fund Cite Details (For Cashier):</strong></u></td> "_
		& "    <tr> "_       
		& "        <td class='FontContent' colspan='6'>&nbsp; Kn " & formatnumber(CellPhonePrsBillRp_,-1) & " " & FiscalDataNonVAT_ & "</td></tr> "_
		& "    <tr> "_       
		& "        <td colspan='6'>&nbsp; </td></tr> "_
		& "    <tr> "_       
		& "        <td class='FontContent' colspan='6'>&nbsp; Please contact <a href='mailto:" & BillingDL & "'>" & BillingDL & "</a> if you have any questions</td></tr> " 

	end if

    		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr> "_       
		& "        <td colspan='6'>&nbsp; </td></tr> "    

		strsql = "Select * From vwMonthlyBilling Where LoginID ='" & LoginID_ & "' and (ProgressId='4' or ProgressId='5')"
		set AwaitingRS = server.createobject("adodb.recordset")
		set AwaitingRS = BillingCon.execute(strsql)

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "  <table cellspadding='1' cellspacing='0' width='100%' bgColor='white'>  "_  
		& "    <tr> "_ 
		& "        <td align='Left' class='FontContent'><u><strong>Accumulated Debt:</strong></u></TD> "_ 
		& "    </tr> " 

if not AwaitingRS.eof Then
		
		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr> "_ 
		& "    	<td class='FontContent'>*Click on each bill for more detail</td> "_ 
		& "    </tr> "_ 
		& "    <tr> "_ 
		& "    	<td align='Left'> "_ 
		& "    	<table cellspadding='1' border='1' bordercolor='black' cellspacing='0' width='100%' bgColor='white'>  "_  
		& "    	<tr align='center' height=26> "_ 
		& "    		<td width='25%' class='FontContent'><strong>Action</strong></td> "_ 
		& "    		<td width='25%' class='FontContent'><strong>Billing Period</strong></td> "_ 
		& "    	<!--	<td width='20%' class='FontContent'><strong>Status</strong></td> --> "_ 
		& "    		<td width='25%' class='FontContent'><strong>Total Bill</strong></td> "_ 
		& "    		<td width='25%' class='FontContent'><strong>Personal Amount Due</strong></td> "_ 
		& "    	</tr> " 

   PrevEmpName_ =""
   TotalBill_ = 0
   TotalPrs_ = 0
   do while not AwaitingRS.eof

		AwaitingEmpName_ = AwaitingRS("EmpName")
		AwaitingMobilePhone_ = AwaitingRS("MobilePhone")
		AwaitingMonthP_ = AwaitingRS("MonthP") 
		AwaitingYearP_ = AwaitingRS("YearP")
		AwaitingProgressDesc_ = AwaitingRS("ProgressDesc")
		AwaitingTotalBillingRp_ = AwaitingRS("TotalBillingRp")
		AwaitingTotalBillingAmountPrsRp_ = AwaitingRS("TotalBillingAmountPrsRp")

	    if bg="#dddddd" then bg="#ffffff" else bg="#dddddd" 
	    if PrevEmpName_ <> AwaitingEmpName_ Then
		SubTotalBill_ = 0
		SubTotalPrs_ = 0

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr BGCOLOR='#999999' height=26> "_ 
		& "    <td colspan='2' style='border: none;' class='FontContent'><FONT color=#FFFFFF><strong>&nbsp;Employee Name : " & AwaitingEmpName_ & "</strong></font></td> "_ 
		& "    	<td colspan='2' style='border: none;' align='right' class='FontContent'><FONT color=#FFFFFF><strong>Phone Number : " & AwaitingMobilePhone_ & "&nbsp;</strong></font></td> "_ 
		& "    	</tr>	 " 	

	    end if
		
		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_      
	   	& "    <TR bgcolor=" & bg & "> "_ 
	        & "    <TD class='FontContent'>&nbsp;<a href='" & WebSiteAddress & "/CellPhoneDetail.asp?CellPhone=" & AwaitingMobilePhone_ & "&MonthP=" & AwaitingMonthP_ & "&YearP=" & AwaitingYearP_ & "' target='_blank'>View Submitted Bill</a></TD> "_ 
	        & "    <TD align='right' class='FontContent'>&nbsp;" & AwaitingMonthP_ & "-" & AwaitingYearP_ & "</font>&nbsp;</TD> "_ 
	        & "   <!-- <TD align='right' class='FontContent'>" & AwaitingProgressDesc_ & "&nbsp;</font></TD> --> "_ 
	        & "    <TD align='right' class='FontContent'>" & formatnumber(AwaitingTotalBillingRp_,-1) & "&nbsp;</font></TD> "_ 
		& "    <TD align='right' class='FontContent'>&nbsp;" & formatnumber(AwaitingTotalBillingAmountPrsRp_,-1) & "&nbsp;</font></TD> "_ 
	   	& "    </TR> " 

   		SubTotalBill_ = cdbl(SubTotalBill_) + cdbl(AwaitingTotalBillingRp_)
		SubTotalPrs_ = cdbl(SubTotalPrs_) + cdbl(AwaitingTotalBillingAmountPrsRp_)
		TotalBill_ = cdbl(TotalBill_) + cdbl(AwaitingTotalBillingRp_)
		TotalPrs_ = cdbl(TotalPrs_) + cdbl(AwaitingTotalBillingAmountPrsRp_)
		PrevEmpName_ = AwaitingEmpName_
	   AwaitingRS.movenext
		if AwaitingRS.eof Then 

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr> "_ 
		& "    	<td colspan='2' class='FontContent'><strong>&nbsp;SubTotal</strong></td> "_ 			
		& "    	<td align='right' class='FontContent'><strong>&nbsp;" & formatnumber(SubTotalBill_,-1) & "&nbsp;</font></strong></td> "_ 			
		& "    	<td align='right' class='FontContent'><strong>&nbsp;" & formatnumber(SubTotalPrs_,-1) & "&nbsp;</font></strong></td> "_ 			
		& "    </tr> " 

		elseif (PrevEmpName_ <> AwaitingRS("EmpName")) Then 

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr> "_ 
		& "    	<td colspan='2' class='FontContent'><strong>&nbsp;SubTotal</strong></td> "_ 			
		& "    	<td align='right' class='FontContent'><strong>&nbsp;" & formatnumber(SubTotalBill_,-1) & "&nbsp;</font></strong></td> "_ 			
		& "    	<td align='right' class='FontContent'><strong>&nbsp;" & formatnumber(SubTotalPrs_,-1) & "&nbsp;</font></strong></td> "_ 			
		& "    </tr> " 
	
		end if
   loop 

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr  BGCOLOR='#999999' height=26> "_
		& "    	<td colspan='2' class='FontContent'><FONT color=#FFFFFF><strong>&nbsp;Total</strong></font></td> "_			
		& "    	<td align='right' class='FontContent'><FONT color=#FFFFFF><strong>&nbsp;" & formatnumber(TotalBill_,-1) & "&nbsp;</font></strong></td>	 "_		
		& "    	<td align='right' class='FontContent'><FONT color=#FFFFFF><strong>&nbsp;" & formatnumber(TotalPrs_,-1) & "&nbsp;</font></strong></td> "_			
		& "    </tr> "


	        strsql = " select CashierMinimumAmount from PaymentDueDate"
       		set rst1 = server.createobject("adodb.recordset") 
        	set rst1 = BillingCon.execute(strsql)
	       	if not rst1.eof then 
		   CashierMinimumAmount_ = rst1("CashierMinimumAmount")
        	end if
		if cdbl(TotalPrs_) > cdbl(CashierMinimumAmount_) Then

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    	<tr BGCOLOR='#990000' height=36> "_
		& "    		<td  colspan='4' align='center' class='FontContent'><FONT color=#FFFFFF><strong>Your total accumulated debt is greater than " & formatnumber(CashierMinimumAmount_,-1) & " Kuna. Please make the payment at cashier office.<br>" & CashierInfo & "</font></strong></td> "_
		& "    	</tr> "
		else
		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    	<tr BGCOLOR='#999999' height=26> "_
		& "    		<td colspan='4' align='center' class='FontContent'><FONT color=#FFFFFF><strong>Your total accumulated debt is less than " & formatnumber(CashierMinimumAmount_,-1) & " Kuna. No payment is necessary at this point.</font></strong></td> "_
		& "    	</tr> "
		end if

	ObjMail.HTMLBody = ObjMail.HTMLBody & " "_ 			
	& "    	</table> "_
	& "    	</td> "_
& "    	</tr> "

else

ObjMail.HTMLBody = ObjMail.HTMLBody & " "_ 
& "    	<tr> "_
& "    		<td align='Left'> "_
& "    		<table cellspadding='1' border='1' bordercolor='black' cellspacing='0' width='100%' bgColor='white' border='0'>   "_
& "    		<tr align='center' BGCOLOR='#999999' height=26> "_
& "    			<td class='FontContent'><FONT color=#FFFFFF><strong>There is no accumulated debt for your cell phone(s).</strong></font></td> "_
& "    		</tr> "

end if

ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
& "    </table> "_


		& "    <tr> "_       
		& "        <td height=26 align='center' colspan='6' class='FontContent'>NOTE: This e-mail was automatically generated. Please do not respond to this e-mail address.</td> "_ 
		& "    </tr> "_ 
		& " </table></p>"_ 
		& "</body></html>"






		'	response.write "send mail"
			objMail.Send 
			Set objMail = Nothing 
			Set objConfig = Nothing 
		End If
		
		'Update MonthlyBill
		'strsql = "Update MonthlyBilling Set ProgressId=" & ProgressId_ & ", ProgressIdDate=GetDate(), SupervisorRemark='" & Remark_ & "', SupervisorApproveDate='" & Date() & "' Where EmpID='" & EmpID_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ &"'"
		'BillingCon.execute(strsql)
		'response.write strsql

	next
  End If

  response.redirect("BillingApprovalList.asp?ProgressID=2")
%>
</body> 
</html>