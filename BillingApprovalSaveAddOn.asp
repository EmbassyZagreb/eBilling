<HTML>
<HEAD>
<!--#include file="connect.inc" -->
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 

<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
<CENTER>
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

   if Status_ = "A" then
	if (cdbl(TotalBillingAmount_) < cdbl(CeilingAmount_ )) or (cdbl(TotalBillingAmount_)=0 ) Then
		ProgressId_ = 7
	Else
		ProgressId_ = 4
	End if
   Else
	ProgressId_ = 3
   end if
%>

<%
	strsql = "Select * From vwMonthlyBilling Where EmpID='" & EmpID_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "'"
	'response.write BillType_ & "<Br>"  
	'response.write strsql & "<Br>"  
	set rsData = server.createobject("adodb.recordset") 
	set rsData = BillingCon.execute(strsql)
	if not rsData.eof then
		Period_ = rsData("MonthP") & rsData("YearP")
		EmpName_ = rsData("EmpName")
		Office_ = rsData("Agency") & " - " & rsData("Office")
		Position_ = rsData("WorkingTitle")
		OfficePhone_ = rsData("WorkPhone")
		HomePhone_ = rsData("HomePhone")
		MobilePhone_ = rsData("MobilePhone")	
		EmpEmail_ = rsData("EmailAddress")
		LoginID_ = rsData("LoginID")
		FiscalData_ = rsData("FiscalStripVAT")
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
	'objMail.CC = send_cc

'	if (Status_ = "A") or (Status_ = "P") Then
	if (Status_ = "A") and (ProgressId_ <> "7") Then
		objMail.Subject = "Info: eBilling System � Approval Notification"
		objMail.HTMLBody = "<html><head>"
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
				& "<tr align='center'><td><b>Billing Period</b></td><td><b>Bill Type</b></td><td><b>Description</b></td><td><b>Amount Due (USD)</b></td><td><b>Exchange Rate</b></td><td><b>Amount Due (kn)</b></td></tr>"
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

	Elseif (Status_ = "A") and (ProgressId_ ="7") Then
		objMail.Subject = "Info: eBilling System � Approval Notification"
		objMail.HTMLBody = "<html><head>"
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
		& "        <td>- Period : <b>" & Period_ & "</b></td></tr> "_         

		& "    <tr> "_       
		& "        <td>- Total Personal Usage : <b>Kn. " & formatnumber(TotalBillingAmount_,-1) & "</b>.&nbsp;&nbsp;<font color='blue'><b>Your personal usage amount is less than the threshold amount, no payment is required.</b></font></td></tr> "_ 

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
	
	Elseif Status_ = "C" Then

		objMail.Subject = "Action Required: eBilling System � Correction Notification"
		objMail.HTMLBody = "<html><head>"
		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
	
		& " <title>e-Billing Application</title> "_              
		& " <meta name='Microsoft Border' content='none, default'> "_     
		& " </head><body bgcolor='#ffffff'> "_              
		& " <p><table border='0' width='100%'> "_    

		& "    <tr> "_       
		& "        <td>Click <A HREF='" & WebSiteAddress & "/MonthlyBilling.asp'>here</A> to correct and resubmit your invoice</td> "_ 


		& "    <tr> "_           
		& "        <td> Your billing <b>has been returned to you for CORRECTIONS</b> by Your Supervisor. Please make the necessary changes and re-submit it.</td></tr> "_      

		& "    <tr> "_       
		& "        <td><br></td></tr> "_      
				
		& "    <tr> "_       
		& "        <td>Approver's Remarks/Corrections: <i>" & Remark_  & ".</i></td></tr> "_      
		
		& "    <tr> "_       
		& "        <td><br></td></tr> "_       

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
		& "        <td>- Total Personal Usage : <b>Kn. " & formatnumber(TotalBillingAmount_,-1) & ".</b></td></tr> "_      

		& "    <tr> "_       
		& "        <td>- Total Billing Amount : <b>kn. " & formatnumber(TotalCost_,-1) & ".</b></td></tr> "_ 

		& "    <tr> "_       
		& "        <td>- Remark/Correction  : <b>" & Remark_ & "</b></td></tr> "_  
     
		& "    <tr> "_       
		& "        <td>&nbsp; </td></tr> "_      

		& "    <tr> "_       
		& "        <td>&nbsp; </td></tr> "_      
	
		& "    <tr> "_       
		& "        <td align='middle'>NOTE: This e-mail was automatically generated. Please do not respond to this e-mail address.</td> "_       
		& "    </tr> "_ 

		& " </table></p>"_ 
		& "</body></html>"
	
	End If
'	response.write Status_ & "<br>"
'	response.write "send mail"
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
	strsql = "Update MonthlyBilling Set ProgressId=" & ProgressId_ & ", ProgressIdDate=GetDate(), SupervisorRemark='" & Remark_ & "', SupervisorApproveDate='" & Date() & "' Where EmpID='" & EmpID_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ &"'"
'	strsql = "Update BillingHd Set Status='" & Status_ & "', SpvRemark='" & Remark_ & "', SpvApprovalDate='" & Date() & "' Where Extension='" & Extension_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ & "'"
	BillingCon.execute(strsql)
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
		<input type="button" value="Back" id="btnMain" onclick="javascript:document.location.href('BillingApprovalListAddOn.asp')">
	        &nbsp;&nbsp;<input type="button" value="Close this window" onclick="javascript:window.close();" name=btnclose>
        </td>
</tr>
</table>
</BODY>
</HTML>