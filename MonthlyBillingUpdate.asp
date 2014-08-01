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
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Monthly Billing</TD>
   </TR>
   <tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
   </tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<%	
	MobilePhone_ = Request.Form("txtMobilePhone")
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
	Position_ = replace(Request.Form("txtPosition"),"'","''")
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

	objMail.Subject = "Action Required: eBilling System – Approval Request"

	objMail.HTMLBody = "<html><head>"
	ObjMail.HTMLBody = ObjMail.HTMLBody & " "_




& " <title>e-Billing Application</title> "_              
					& " <meta name='Microsoft Border' content='none, default'><style type='text/css'><!--.FontContent {font-family: verdana;font-size: 11px;color: black;}--></style> "_    
					& " </head><body bgcolor='#ffffff'> "_              
					& " <p><table cellspadding='1' cellspacing='0' width='80%' bgColor='white'>"_ 
					& "    <tr> "_           
					& "        <td colspan='6' align='center'><font face='Verdana, Arial, Helvetica' color='#999999' size='5'>eBilling System – Approval Request</font></td></tr> "_
					& "    <tr> "_       
					& "        <td colspan='6'>&nbsp; </td></tr> "_       
					& "    <tr> "_           
					& "        <td colspan='6' class='FontContent'> " & EmpName_  & " has submitted his/her cell phone billing and need you to approve it.</td></tr> "_
					& "    <tr> "_       
					& "        <td colspan='6'>&nbsp; </td></tr> "_   

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
		& "    	</table></td></tr> " _
		& "    <tr> "_ 
		& "        <td align='Left' colspan='6' class='FontContent'>&nbsp;<u><b>Billing Detail:<b></u></td></tr> "_
		& "    <tr> "_
		& "    <td align='Left' colspan='6'> "_
		& "    <table cellspadding='1' border='1' bordercolor='black' cellspacing='0' width='100%' bgColor='white'> "_  
		& "    	<tr align='center' height=26> "_
		& "    		<td width='20%' class='FontContent'><b>Action</b></td> "_
		& "    		<td width='20%' class='FontContent'><b>Billing Period</b></td> "_
		& "    		<td width='20%' class='FontContent'><b>Status</b></td> "_
		& "    		<td width='20%' class='FontContent'><b>Billing (Kn.)</b></td> "_
		& "    		<td width='20%' class='FontContent'><b>Personal Amount (Kn.)</b></td> "_
		& "    	</tr> "

		if cdbl(TotalCost_ ) > 0 Then

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    	<tr height=26> "_
		& "    	<td class='FontContent'>&nbsp;<a href='" & WebSiteAddress & "/BillingApproval.asp?EmpID=" & EmpID_ & "&Month=" & MonthP_ & "&Year=" & YearP_ & "' target='_blank'>Review and approve</a></td> "_
	        & "    	<TD align='right' class='FontContent'>&nbsp;" & MonthP_ & "-" & YearP_ & "</font>&nbsp;</TD> "_
	        & "    	<TD align='right' class='FontContent'>Waiting Approval from Supervisor&nbsp;</font></TD> "_
		& "    	<td align='right' class='FontContent'>" & formatnumber(TotalCost_  ,-1) & "&nbsp;</td> "_
		& "    	<td align='right' class='FontContent'>" & formatnumber(TotalBillingPrsAmount_ ,-1) & "&nbsp;</td> "_
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

		& "    	</table></td></tr> " _
		& "    <tr> "_           
		& "        <td colspan='6' align='Left' class='FontContent'>&nbsp;<u><b>Employee's Note:<b></u></td></tr> "_
		& "    <tr> "_ 
		& "        <td colspan='6' align='Left' class='FontContent'>" & Notes_ & "</td></tr> "_
      
		& "        <td colspan='6'>&nbsp; </td></tr> "_     
					& "    <tr> "_       
					& "        <td height=26 align='center' colspan='6' class='FontContent'>NOTE: This e-mail was automatically generated.</td> "_ 
					& "    </tr> "_       
		& "        <td colspan='6'>&nbsp; </td></tr> "_  
			
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
<!--<tr>
	<td align="center">
		<input type="button" value="Back" onClick="window.location.href('OfficePhoneDetail.asp?Extension=<%=Extension_ %>&MonthP=<%=MonthP_ %>&YearP=<%=YearP_ %>')" </input>
		&nbsp;&nbsp;<input type="button" value="Close" onClick="window.close()" </input> 
		<br><a href="AgencyList.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a>
	</td>
</tr> -->
<%
Response.AddHeader "REFRESH","1;URL=MonthlyBilling.asp?CellPhone=" & MobilePhone_ & "&MonthP=" & MonthP_ & "&YearP=" & YearP_ & ""
%>
</table>
</body> 

</html>
