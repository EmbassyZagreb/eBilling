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
  	<TD COLSPAN="4" ALIGN="center" Class="title">AR REMINDER</TD>
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

	Dim send_from, send_to, send_cc, noMail
	send_from = BillingDL
	Dim ObjMail
	Set ObjMail = Server.CreateObject("CDO.Message")
	Set objConfig = CreateObject("CDO.Configuration") 
	objConfig.Fields(cdoSendUsingMethod) = 2  
	objConfig.Fields(cdoSMTPServer) = SMTPServer
	objConfig.Fields.Update 
	Set objMail.Configuration = objConfig 
	noMail=0
	For Each loopIndex in Request("cbApproval")
		'response.write loopIndex & "<br>"
		X = len(loopIndex)
		'response.write X & "<br>"
		MobilePhone_ = Left(loopIndex, X-6)
		'response.write EmpID_ & "<br>"
		Period = right(loopIndex,6)
		MonthP_ = left(Period,2)
		'response.write MonthP_ & "<br>"
		YearP_ = Right(Period,4)
		'response.write YearP_ & "<br>"
		strsql = "Select * From vwMonthlyBilling Where MobilePhone='" & MobilePhone_ & "' And MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "'"
'		response.write strsql & "<Br>"  
		set rsData = server.createobject("adodb.recordset") 
		set rsData = BillingCon.execute(strsql)
		Period_ = MonthP_ & " - " & YearP_
		if not rsData.eof then
			EmpName_ = rsData("EmpName")
			Office_ = rsData("Agency") & " - " & rsData("Office")
			Position_ = rsData("WorkingTitle")
			'OfficePhone_ = rsData("WorkPhone")
			'HomePhone_ = rsData("HomePhone")
			'MobilePhone_ = rsData("MobilePhone")
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
			ProgressDesc_ = rsData("ProgressDesc")
			ProgressID_ = rsData("ProgressID")
			LoginID_ = rsData("LoginID") 
		end if
		
		'Send mail
		send_to = EmpEmail_ 
		'send_to = "kurniawane@state.gov"
		'response.write send_to
		objMail.From = send_from
		objMail.To = send_to 	
		objMail.Subject = "Action Required: eBilling System – Monthly Billing Repeat Notice"
		objMail.HTMLBody = "<html><head>"
		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
	
		& " <title>e-Billing Application</title> "_              
		& " <meta name='Microsoft Border' content='none, default'><style type='text/css'><!--.FontContent {font-family: verdana;font-size: 11px;color: black;}--></style> "_     
		& " </head><body bgcolor='#ffffff'> "_              
		& " <p><table cellspadding='1' cellspacing='0' width='80%' bgColor='white'>"_    
		& "    <tr> "_           
		& "        <td colspan='6' align='center'><font face='Verdana, Arial, Helvetica' color='#999999' size='5'>eBilling Reminder Notice</font></td></tr> "_
		& "    <tr> "_           
		& "        <td colspan='6' align='Left' class='FontContent'><u><b>&nbsp;Personal Info:<b></u></td></tr> "_
		& "    <tr> "_ 
		& "    <td colspan='6' align='Left'> "_
		& "    	<table cellspadding='1' border='2' bordercolor='black' cellspacing='3' width='100%' bgColor='#999999' border='0'>   "_
		& "    		<tr BGCOLOR='#999999'> "_
		& "    			<td colspan='3' style='border: none;' class='FontContent'><FONT color=#FFFFFF><b>Employee Name : " & EmpName_ & "</b></font></td> "_
		& "    			<td colspan='3' style='border: none;' align='right' class='FontContent'><FONT color=#FFFFFF><b>Phone Number : " & MobilePhone_ & "&nbsp;</b></font></td> "_
		& "    		</tr> "_
		& "    		<tr BGCOLOR='#999999'> "_
		& "    			<td colspan='6' style='border: none;' class='FontContent'><FONT color=#FFFFFF><b>Position : " & Position_ & "</b></font></td> "_
		& "    		</tr> "_
		& "    		<tr BGCOLOR='#999999'> "_
		& "    			<td colspan='6' style='border: none;' class='FontContent'><FONT color=#FFFFFF><b>Agency / Office : " & Office_ & "</b></font></td> "_
		& "    		</tr> "_
		& "    	</table></td></tr> " _
		& "    <tr> "_ 
		& "        <td align='Left' colspan='6' class='FontContent'><u><b>&nbsp;Billing Detail:<b></u></td></tr> "_
		& "    <tr> "_
		& "    <td align='Left' colspan='6'> "_
		& "    <table cellspadding='1' border='1' bordercolor='black' cellspacing='0' width='100%' bgColor='white'> "_  
		& "    	<tr align='center' height=26> "_
		& "    		<td width='20%' class='FontContent'><b>Action</b></td> "_
		& "    		<td width='20%' class='FontContent'><b>Billing Period</b></td> "_
		& "    		<td width='20%' class='FontContent'><b>Status</b></td> "_
		& "    		<td width='20%' class='FontContent'><b>Billing (Kn.)</b></td> "_
		& "    		<td width='20%' class='FontContent'><b>Should be paid (Kn.)</b></td> "_
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
	        & "    	<TD align='right' class='FontContent'>" & ProgressDesc_ & "&nbsp;</font></TD> "_
		& "    	<td align='right' class='FontContent'>" & formatnumber(CellPhoneBillRp_  ,-1) & "&nbsp;</td> "_
		& "    	<td align='right' class='FontContent'>" & formatnumber(CellPhonePrsBillRp_ ,-1) & "&nbsp;</td> "_
		& "    	</tr> "

		else

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr> "_
		& "    <td>Mobile Phone</td> "_
		& "    <td class='FontContent' align='right'>- &nbsp;</td> "_
		& "    <td class='FontContent' align='right'>- &nbsp;</td> "_
		& "    <td class='FontContent' align='right'>- &nbsp;</td> "_
		& "    <td class='FontContent' align='right'>- &nbsp;</td> "_
		& "    </tr> "

		end if

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_

		& "    </table></td><tr> "_       
		& "        <td colspan='6'>&nbsp; </td></tr> "_  
		& "    <tr> "_       
		& "        <td class='FontContent' colspan='6'><p>Your bill is still in <b>" & ProgressDesc_ & "</b> status.</td></tr> "_      
		& "    <tr> "_       
		& "        <td class='FontContent' colspan='6'> "_
		& "        <p>You must take action. Click <A HREF='" & WebSiteAddress & "/MonthlyBilling.asp'>Here</A> to review your invoice. </p><p>&nbsp;</p></td> "_  
		& "    <tr> "_       
		& "        <td colspan='6'>&nbsp; </td></tr> "_      
		& "    <tr> "_       
		& "        <td colspan='6'>&nbsp; </td></tr> "      

		strsql = "Select * From vwMonthlyBilling Where LoginID ='" & LoginID_ & "' and (ProgressId='4' or ProgressId='5')"
		set AwaitingRS = server.createobject("adodb.recordset")
		set AwaitingRS = BillingCon.execute(strsql)

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "  <table cellspadding='1' cellspacing='0' width='100%' bgColor='white'>  "_  
		& "    <tr> "_ 
		& "        <td align='Left' class='FontContent'><u><b>Accumulated Debt :</b></u></TD> "_ 
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
		& "    		<td width='20%' class='FontContent'><b>Action</b></td> "_ 
		& "    		<td width='20%' class='FontContent'><b>Billing Period</b></td> "_ 
		& "    		<td width='20%' class='FontContent'><b>Status</b></td> "_ 
		& "    		<td width='20%' class='FontContent'><b>Billing (Kn.)</b></td> "_ 
		& "    		<td width='20%' class='FontContent'><b>Should be paid (Kn.)</b></td> "_ 
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
		& "    <td colspan='2' style='border: none;' class='FontContent'><FONT color=#FFFFFF><b>&nbsp;Employee Name : " & AwaitingEmpName_ & "</b></font></td> "_ 
		& "    	<td colspan='3' style='border: none;' align='right' class='FontContent'><FONT color=#FFFFFF><b>Phone Number : " & AwaitingMobilePhone_ & "&nbsp;</b></font></td> "_ 
		& "    	</tr>	 " 	

	    end if
		
		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_      
	   	& "    <TR bgcolor=" & bg & "> "_ 
	        & "    <TD class='FontContent'>&nbsp;<a href='" & WebSiteAddress & "/CellPhoneDetail.asp?CellPhone=" & AwaitingMobilePhone_ & "&MonthP=" & AwaitingMonthP_ & "&YearP=" & AwaitingYearP_ & "' target='_blank'>View Submitted Bill</a></TD> "_ 
	        & "    <TD align='right' class='FontContent'>&nbsp;" & AwaitingMonthP_ & "-" & AwaitingYearP_ & "</font>&nbsp;</TD> "_ 
	        & "    <TD align='right' class='FontContent'>" & AwaitingProgressDesc_ & "&nbsp;</font></TD> "_ 
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
		& "    	<td colspan='3' class='FontContent'><b>&nbsp;SubTotal</b></td> "_ 			
		& "    	<td align='right' class='FontContent'><b>&nbsp;" & formatnumber(SubTotalBill_,-1) & "&nbsp;</font></b></td> "_ 			
		& "    	<td align='right' class='FontContent'><b>&nbsp;" & formatnumber(SubTotalPrs_,-1) & "&nbsp;</font></b></td> "_ 			
		& "    </tr> " 

		elseif (PrevEmpName_ <> AwaitingRS("EmpName")) Then 

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr> "_ 
		& "    	<td colspan='3' class='FontContent'><b>&nbsp;SubTotal</b></td> "_ 			
		& "    	<td align='right' class='FontContent'><b>&nbsp;" & formatnumber(SubTotalBill_,-1) & "&nbsp;</font></b></td> "_ 			
		& "    	<td align='right' class='FontContent'><b>&nbsp;" & formatnumber(SubTotalPrs_,-1) & "&nbsp;</font></b></td> "_ 			
		& "    </tr> " 
	
		end if
   loop 

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr  BGCOLOR='#999999' height=26> "_
		& "    	<td colspan='3' class='FontContent'><FONT color=#FFFFFF><b>&nbsp;Total</b></font></td> "_			
		& "    	<td align='right' class='FontContent'><FONT color=#FFFFFF><b>&nbsp;" & formatnumber(TotalBill_,-1) & "&nbsp;</font></b></td>	 "_		
		& "    	<td align='right' class='FontContent'><FONT color=#FFFFFF><b>&nbsp;" & formatnumber(TotalPrs_,-1) & "&nbsp;</font></b></td> "_			
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
		& "    		<td  colspan='5' align='center' class='FontContent'><FONT color=#FFFFFF><b>Your total accumulated debt is greater than " & formatnumber(CashierMinimumAmount_,-1) & " Kuna. Please make the payment at cashier office.<br>" & CashierInfo & "</font></b></td> "_
		& "    	</tr> "
		else
		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    	<tr BGCOLOR='#999999' height=26> "_
		& "    		<td colspan='5' align='center' class='FontContent'><FONT color=#FFFFFF><b>Your total accumulated debt is less than " & formatnumber(CashierMinimumAmount_,-1) & " Kuna. No payment is necessary at this point.</font></b></td> "_
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
& "    			<td class='FontContent'><FONT color=#FFFFFF><b>There is no accumulated debt for your cell phone(s).</b></font></td> "_
& "    		</tr> "

end if

ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
& "    </table> "_


		& "    <tr> "_       
		& "        <td height=26 align='center' colspan='6' class='FontContent'>NOTE: This e-mail was automatically generated. Please do not respond to this e-mail address.</td> "_ 
		& "    </tr> "_ 
		& " </table></p>"_ 
		& "</body></html>"

		objMail.Send 
		noMail=noMail+1
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
	<td align="center"><br><a href="ARReminder.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
</tr>
</table>
</body> 
</html>