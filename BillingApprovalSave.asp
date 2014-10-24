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
  	<TD COLSPAN="4" ALIGN="center" Class="title">OFFICE PHONE BILLING</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>

<%
   Status_ = Request.Form("SupervisorSign")
   'response.write Status_ & "<br>"
   Remark_ = replace(Request.Form("txtRemark"),"'","''")
   Extension_ = Request.Form("txtExtension")
   'response.write Extension_ & "<br>"
   MonthP_ = Request.Form("txtMonthP")
   'response.write MonthP_ & "<br>"
   YearP_ = Request.Form("txtYearP")
   'response.write YearP_ & "<br>"
   EmpEmail_ = Request.Form("txtEmpEmail")
   'response.write EmpEmail_ & "<br>"

   EmpID_ = replace(Request.Form("txtEmpID"),"'","''")
  ' response.write EmpID_ & "<br>"
   EmpName_ = replace(Request.Form("txtEmpName"),"'","''")
   Period_ = replace(Request.Form("txtPeriod"),"'","''")
   Office_ = replace(Request.Form("txtOffice"),"'","''")
   TotalCost_ = replace(Request.Form("txtTotalCost"),"'","''")
   TotalBillingAmount_ = replace(Request.Form("txtTotalBillingAmount"),"'","''")
   'response.write TotalBillingAmount_ 
   Notes_  = replace(Request.Form("txtNote"),"'","''")


   strsql = "Select CeilingAmount From PaymentDueDate"
   'response.write strsql & "<br>"
   set rsDataX = server.createobject("adodb.recordset") 
   set rsDataX = BillingCon.execute(strsql) 
   if not rsDataX.eof then
	CeilingAmount_ = rsDataX("CeilingAmount")
   else
	CeilingAmount_ = 0
   end if
 '  TotalBillingAmount_ = 300000
   if Status_ = "A" then
	if (cdbl(TotalBillingAmount_) < cdbl(CeilingAmount_ )) or (cdbl(TotalBillingAmount_)=0 ) Then
		ProgressId_ = 7
	Else
		ProgressId_ = 4
	End if
   Else
	ProgressId_ = 3
   end if

	'3. SAVING TO Billing Header
	strsql = "Update MonthlyBilling Set ProgressId=" & ProgressId_ & ", ProgressIdDate=GetDate(), SupervisorRemark='" & Remark_ & "', SupervisorApproveDate='" & Date() & "' Where EmpID='" & EmpID_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ &"'"
'	strsql = "Update BillingHd Set Status='" & Status_ & "', SpvRemark='" & Remark_ & "', SpvApprovalDate='" & Date() & "' Where Extension='" & Extension_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ & "'"
	BillingCon.execute(strsql)
	'response.write strsql
%>

<%
	strsql = "Select * From vwMonthlyBilling Where EmpID='" & EmpID_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "'"
	'response.write BillType_ & "<Br>"  
	'response.write strsql & "<Br>"  
	set rsData = server.createobject("adodb.recordset") 
	set rsData = BillingCon.execute(strsql)
		if not rsData.eof then
			Period_ = rsData("MonthP") & rsData("YearP")
			FiscalDataVAT_ = rsData("FiscalStripVAT")
			FiscalDataNonVAT_ = rsData("FiscalStripNonVAT")
			EmpName_ = rsData("EmpName")
			Office_ = rsData("Agency") & " - " & rsData("Office")
			Position_ = rsData("WorkingTitle")
			'OfficePhone_ = rsData("WorkPhone")
			'HomePhone_ = rsData("HomePhone")
			MobilePhone_ = rsData("MobilePhone")
			CellPhoneBillRp_ = rsData("CellPhoneBillRp")
			CellPhoneBillDlr_ = rsData("CellPhoneBillDlr")
			CellPhonePrsBillRp_ = rsData("CellPhonePrsBillRp")
			CellPhonePrsBillDlr_ = rsData("CellPhonePrsBillDlr")
			TotalBillingRp_ = rsData("TotalBillingRp")
			TotalBillingDlr_ = rsData("TotalBillingDlr")
			EmpEmail_ = rsData("EmailAddress")
			ProgressDesc_ = rsData("ProgressDesc")
			LoginID_ = rsData("LoginID") 
		end if


	'response.write "send mail"
	Dim send_from, send_to, send_cc, send_bcc

	send_from = BillingDL
	send_to = EmpEmail_ 

	'send_to = "kurniawane@state.gov"
	Dim ObjMail
	Set ObjMail = Server.CreateObject("CDO.Message")
	Set objConfig = CreateObject("CDO.Configuration") 
	objConfig.Fields(cdoSendUsingMethod) = 2  
	objConfig.Fields(cdoSMTPServer) = SMTPServer
	objConfig.Fields.Update 
	Set objMail.Configuration = objConfig 
	objMail.From = send_from
	objMail.To = send_to 
	'objMail.To = "kurniawane@state.gov"
	'objMail.CC = send_cc


		objMail.HTMLBody = "<html><head>"
		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
	
		& " <title>e-Billing Application</title> "_              
		& " <meta name='Microsoft Border' content='none, default'><style type='text/css'><!--.FontContent {font-family: verdana;font-size: 11px;color: black;}--></style> "_     
		& " </head><body bgcolor='#ffffff'> "_              
		& " <p><table cellspadding='1' cellspacing='0' width='80%' bgColor='white'>"  

	if (Status_ = "A") Then
		
		objMail.Subject = "Info: eBilling System – Approval Notification"

  		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr> "_           
		& "        <td colspan='6' align='center'><font face='Verdana, Arial, Helvetica' color='#999999' size='5'>eBilling Approval Notification</font></td></tr> "_
		& "    <tr> "_ 
		& "        <td colspan='6'>&nbsp; </td></tr> " _
		& "    <tr> "_           
		& "        <td colspan='6' align='Left' class='FontContent'>Your billing <b>has been APPROVED</b> by your supervisor.</td></tr> "

	Elseif Status_ = "C" Then

		objMail.Subject = "Info: eBilling System – Correction Notification"

  		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr> "_           
		& "        <td colspan='6' align='center'><font face='Verdana, Arial, Helvetica' color='#990000' size='5'>eBilling Correction Notification</font></td></tr> "_
		& "    <tr> "_ 
		& "        <td colspan='6'>&nbsp; </td></tr> " _
		& "    <tr> "_           
		& "        <td colspan='6' align='Left' class='FontContent'>Your billing <b>has been returned to you for CORRECTIONS</b> by your supervisor. Please make the necessary changes and re-submit it.</td></tr> "_
		& "    <tr> "_       
		& "        <td colspan='6' align='Left' class='FontContent'>Approver's Remarks/Corrections: <i>" & Remark_  & ".</i></td></tr> "      
		
	End If

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr> "_       
		& "        <td colspan='6' align='Left' class='FontContent'><br></td></tr> "_    
		& "    <tr> "_           
		& "        <td colspan='6' align='Left' class='FontContent'>&nbsp;<u><b>Personal Info:<b></u></td></tr> "_
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
		& "    	</table></td></tr> "

	if (Status_ = "A") Then

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr> "_ 
		& "        <td align='Right' colspan='6'><font face='Verdana, Arial, Helvetica' color='#999999' size='4'><b>Billing Summary&nbsp;<b></font></td></tr> "

	End If

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr> "_ 
		& "        <td align='Left' colspan='6' class='FontContent'>&nbsp;<u><b>Billing Detail:<b></u></td></tr> "_
		& "    <tr> "_
		& "    <td align='Left' colspan='6'> "_
		& "    <table cellspadding='1' border='1' bordercolor='black' cellspacing='0' width='100%' bgColor='white'> "_  
		& "    	<tr align='center' height=26> "_
		& "    		<td width='25%' class='FontContent'><b>Action</b></td> "_
		& "    		<td width='25%' class='FontContent'><b>Billing Period</b></td> "_
		& "  <!--		<td width='20%' class='FontContent'><b>Status</b></td>  --> "_
		& "    		<td width='25%' class='FontContent'><b>Total Bill</b></td> "_
		& "    		<td width='25%' class='FontContent'><b>Personal Amount Due</b></td> "_
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
	        & "  <!-- 	<TD align='right' class='FontContent'>" & ProgressDesc_ & "&nbsp;</font></TD>  --> "_
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


	if (Status_ = "A") and (ProgressId_ ="7") Then
     
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

	if (Status_ = "A") and (ProgressId_ <> "7") Then

    		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr> "_       
		& "        <td class='FontContent' colspan='6'>&nbsp;<u><b>Fund Cite Details (For Cashier):</b></u></td> "_
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

if (Status_ = "A") Then    

		strsql = "Select * From vwMonthlyBilling Where LoginID ='" & LoginID_ & "' and (ProgressId='4' or ProgressId='5')"
		set AwaitingRS = server.createobject("adodb.recordset")
		set AwaitingRS = BillingCon.execute(strsql)

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "  <table cellspadding='1' cellspacing='0' width='100%' bgColor='white'>  "_  
		& "    <tr> "_ 
		& "        <td align='Left' class='FontContent'><u><b>Accumulated Debt:</b></u></TD> "_ 
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
		& "    		<td width='25%' class='FontContent'><b>Action</b></td> "_ 
		& "    		<td width='25%' class='FontContent'><b>Billing Period</b></td> "_ 
		& "  <!--		<td width='20%' class='FontContent'><b>Status</b></td>  --> "_
		& "    		<td width='25%' class='FontContent'><b>Total Bill</b></td> "_ 
		& "    		<td width='25%' class='FontContent'><b>Personal Amount Due</b></td> "_ 
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
		& "    	<td colspan='2' style='border: none;' align='right' class='FontContent'><FONT color=#FFFFFF><b>Phone Number : " & AwaitingMobilePhone_ & "&nbsp;</b></font></td> "_ 
		& "    	</tr>	 " 	

	    end if
		
		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_      
	   	& "    <TR bgcolor=" & bg & "> "_ 
	        & "    <TD class='FontContent'>&nbsp;<a href='" & WebSiteAddress & "/CellPhoneDetail.asp?CellPhone=" & AwaitingMobilePhone_ & "&MonthP=" & AwaitingMonthP_ & "&YearP=" & AwaitingYearP_ & "' target='_blank'>View Submitted Bill</a></TD> "_ 
	        & "    <TD align='right' class='FontContent'>&nbsp;" & AwaitingMonthP_ & "-" & AwaitingYearP_ & "</font>&nbsp;</TD> "_ 
	        & " <!-- <TD align='right' class='FontContent'>" & AwaitingProgressDesc_ & "&nbsp;</font></TD>   --> "_ 
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
		& "    	<td colspan='2' class='FontContent'><b>&nbsp;SubTotal</b></td> "_ 			
		& "    	<td align='right' class='FontContent'><b>&nbsp;" & formatnumber(SubTotalBill_,-1) & "&nbsp;</font></b></td> "_ 			
		& "    	<td align='right' class='FontContent'><b>&nbsp;" & formatnumber(SubTotalPrs_,-1) & "&nbsp;</font></b></td> "_ 			
		& "    </tr> " 

		elseif (PrevEmpName_ <> AwaitingRS("EmpName")) Then 

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr> "_ 
		& "    	<td colspan='2' class='FontContent'><b>&nbsp;SubTotal</b></td> "_ 			
		& "    	<td align='right' class='FontContent'><b>&nbsp;" & formatnumber(SubTotalBill_,-1) & "&nbsp;</font></b></td> "_ 			
		& "    	<td align='right' class='FontContent'><b>&nbsp;" & formatnumber(SubTotalPrs_,-1) & "&nbsp;</font></b></td> "_ 			
		& "    </tr> " 
	
		end if
   loop 

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr  BGCOLOR='#999999' height=26> "_
		& "    	<td colspan='2' class='FontContent'><FONT color=#FFFFFF><b>&nbsp;Total</b></font></td> "_			
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
		& "    		<td  colspan='4' align='center' class='FontContent'><FONT color=#FFFFFF><b>Your total accumulated debt is greater than " & formatnumber(CashierMinimumAmount_,-1) & " Kuna. Please make the payment at cashier office.<br>" & CashierInfo & "</font></b></td> "_
		& "    	</tr> "
		else
		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    	<tr BGCOLOR='#999999' height=26> "_
		& "    		<td colspan='4' align='center' class='FontContent'><FONT color=#FFFFFF><b>Your total accumulated debt is less than " & formatnumber(CashierMinimumAmount_,-1) & " Kuna. No payment is necessary at this point.</font></b></td> "_
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
& "    </table> "

end if

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    <tr> "_       
		& "        <td height=26 align='center' colspan='6' class='FontContent'>NOTE: This e-mail was automatically generated. Please do not respond to this e-mail address.</td> "_ 
		& "    </tr> "_ 
		& " </table></p>"_ 
		& "</body></html>"


	objMail.Send 
	Set objMail = Nothing 
	Set objConfig = Nothing 

%>

<%
'4. SHOW
%>


<table border="0" align=center width="100%" cellspacing="0" cellpadding="1">    
<tr>
	<td colspan="2" align="center">Billing Period : <Label style="color:blue"><%=MonthP_%> - <%=YearP_%></lable></td>
</tr>
<tr>
	<td colspan="2" align="center"><br></td>
</tr>
<%
	'3. SAVING TO Billing Header
'	strsql = "Update MonthlyBilling Set ProgressId=" & ProgressId_ & ", ProgressIdDate=GetDate(), SupervisorRemark='" & Remark_ & "', SupervisorApproveDate='" & Date() & "' Where EmpID='" & EmpID_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ &"'"
'	strsql = "Update BillingHd Set Status='" & Status_ & "', SpvRemark='" & Remark_ & "', SpvApprovalDate='" & Date() & "' Where Extension='" & Extension_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ & "'"
'	BillingCon.execute(strsql)
	'response.write strsql

%>
<tr>
	<td align="center" colspan="2">Your data has already been saved.</td>
</tr>
<tr>
	<td align="center" colspan="2">Thank you for your approval</td>
</tr>
<tr>
	<td colspan="2"><br></td>
</tr>
<tr>
	<td colspan="2" align=center>
		<input type="button" value="Back" id="btnMain" onclick="javascript:document.location.href('BillingApprovalList.asp')">
	        &nbsp;&nbsp;<input type="button" value="Close this window" onclick="javascript:window.close();" name=btnclose>
        </td>
</tr>
</table>
</BODY>
</HTML>