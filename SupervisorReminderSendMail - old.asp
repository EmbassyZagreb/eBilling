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
  	<TD COLSPAN="4" ALIGN="center" Class="title">SUPERVISOR REMINDER</TD>
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
		EmpID_ = Left(loopIndex, X-6)
		'response.write EmpID_ & "<br>"
		Period = right(loopIndex,6)
		MonthP_ = left(Period,2)
		'response.write MonthP_ & "<br>"
		YearP_ = Right(Period,4)
		'response.write YearP_ & "<br>"
		strsql = "Select * From vwMonthlyBilling Where EmpID='" & EmpID_ & "' And MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "'"
'		response.write strsql & "<Br>"  
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
			TotalBillingPrsAmount_ = rsData("TotalBillingAmountPrsRp")
			TotalCost_ = rsData("TotalBillingRp")
			SPVEmail_ = rsData("SupervisorEmail")
		end if
		
		'Send mail
		send_to = SPVEmail_ 
		'send_to = "kurniawane@state.gov"
		'response.write send_to
		objMail.From = send_from
		objMail.To = send_to 	
		objMail.Subject = "Action Required: eBilling System – Approval Request Repeat Notice"
		objMail.HTMLBody = "<html><head>"
		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
	
		& " <title>Billing Application</title> "_              
		& " <meta name='Microsoft Border' content='none, default'> "_     
		& " </head><body bgcolor='#ffffff'> "_              
		& " <p><table border='0' width='100%'> "_    
		& "    <tr> "_       
		& "        <td>Click <font face='Arial'><A HREF='http://zagrebws03:8080/eBilling/BillingApproval.asp?EmpID=" & EmpID_ & "&Month=" & MonthP_ & "&Year=" & YearP_ & "'> Here </A>to review and approve this request."_
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
		& "        <td>- Phone Ext. : <b>" & OfficePhone_ & "</b></td></tr> "_      

		& "    <tr> "_       
		& "        <td>- Period : <b>" & Period_ & "</b></td></tr> "_      
	
		& "    <tr> "_       
		& "        <td>- Total Personal Usage : <b>Rp. " & formatnumber(TotalBillingPrsAmount_,-1) & "</b></td></tr> "_      

		& "    <tr> "_       
		& "        <td>- Total Billing Amount : <b>Rp. " & formatnumber(TotalCost_,-1) & "</b></td></tr> "_ 

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
	<td align="center"><br><a href="SupervisorReminder.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
</tr>
</table>
</body> 
</html>