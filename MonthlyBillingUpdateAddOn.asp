<%@ Language=VBScript%>
<!--#include file="connect.inc" -->
<% ' VI 6.0 Scripting Object Model Enabled %>
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"> 

<head>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
<CENTER>
<%	
	Extension_ = Request.Form("txtExtension")
	EmpID_ = Request.Form("txtEmpID")
'	response.write Extension_ & "<br>"
	MonthP_ = Request.Form("txtMonthP")
'	response.write MonthP_ & "<br>"
	YearP_ = Request.Form("txtYearP")
'	response.write YearP_ & "<br>"
	SpvEmail_ = Request.Form("cmbSupervisor")
'	response.write SvpMail_ & "<br>"
	Notes_ = replace(Request.Form("txtNotes"),"'","''")
'	response.write Notes_ & "<br>"
	EmpName_ = replace(Request.Form("txtEmpName"),"'","''")
	Period_ = replace(Request.Form("txtPeriod"),"'","''")
	Office_ = replace(Request.Form("txtOffice"),"'","''")
	TotalCost_ = replace(Request.Form("txtTotalCost"),"'","''")
	TotalBillingPrsAmount_ = replace(Request.Form("txtTotalBillingPrsAmount"),"'","''")


	'Save Header
	strsql = "Update MonthlyBilling Set SupervisorEmail='" & SpvEmail_ & "', Notes='" & Notes_ & "' Where EmpID='" & EmpID_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ &"'"
	'response.write strsql & "<Br>"  
	BillingCon.execute(strsql)

	strsql = "Update MonthlyBilling Set ProgressId=2, ProgressIdDate=GetDate() Where EmpID='" & EmpID_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ &"'"
	'response.write strsql & "<Br>"  
	BillingCon.execute(strsql) 

	Dim send_from, send_to, send_cc, send_bcc
	send_from = BillingDL
	send_to = SpvEmail_

'		send_to = "kurniawane@state.gov"
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

	objMail.Subject = "Action Required: eBilling System � Approval Request"

	objMail.HTMLBody = "<html><head>"
	ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
	
	
	& " <title>Billing Application</title> "_              
	& " <meta name='Microsoft Border' content='none, default'> "_     
	& " </head><body bgcolor='#ffffff'> "_              
	& " <p><table border='0' width='100%'> "_    
	& "    <tr> "_       
	& "        <td>Click <font face='Arial'><A HREF='http://jakartaws01.eap.state.sbu/eBilling/BillingApproval.asp?EmpID=" & EmpID_ & "&Month=" & MonthP_ & "&Year=" & YearP_ & "'> Here </A>to review and approve this request."_
	& "		</font></td></tr> "_  
	& "    <tr> "_       
	& "        <td> " & EmpName_  & " has submitted his/her office phone billing and need you to approve.</td></tr> "_      

	& "    <tr> "_       
	& "        <td>The Request is currently Pending.</td></tr> "_      

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
	& "        <td>- Total Personal Usage : <b>Kn. " & formatnumber(TotalBillingPrsAmount_,-1) & "</b></td></tr> "_      

	& "    <tr> "_       
	& "        <td>- Total Billing Amount : <b>Kn. " & formatnumber(TotalCost_,-1) & "</b></td></tr> "_ 

	& "    <tr> "_       
	& "        <td>- Note  : <b>" & Notes_ & "</b></td></tr> "_  
     
	& "    <tr> "_       
	& "        <td>&nbsp; </td></tr> "_      
	
	& "    <tr> "_       
	& "        <td align='middle'>NOTE: This e-mail was automatically generated. Please do not respond to this e-mail address.</td> "_       
	& "    </tr> "_ 

	& " </table></p>"_ 
	& "</body></html>"

	objMail.Send 
	Set objMail = Nothing 
	Set objConfig = Nothing 

'	response.redirect("Default.asp")
%>
<table cellspadding="1" cellspacing="0" width="60%" bgColor="white">
<tr>
	<td><br></td>
</tr>
<tr>
	<td>Your data has already been saved.</td>
</tr>
<tr>
	<td>
<!--		<input type="button" value="Back" onClick="window.location.href('OfficePhoneDetail.asp?Extension=<%=Extension_ %>&MonthP=<%=MonthP_ %>&YearP=<%=YearP_ %>')" </input> -->
		&nbsp;&nbsp;<input type="button" value="Close" onClick="window.close()" </input>
	</td>
</tr>
</table>
</body> 

</html>