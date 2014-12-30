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
			EmpID_ = rsData("EmpID")
			EmpName_ = rsData("EmpName")
			Office_ = rsData("Agency") & " - " & rsData("Office")
			Position_ = rsData("WorkingTitle")
			'OfficePhone_ = rsData("WorkPhone")
			'HomePhone_ = rsData("HomePhone")
			'MobilePhone_ = rsData("MobilePhone")
			TotalBillingPrsAmount_ = rsData("TotalBillingAmountPrsRp")
			TotalCost_ = rsData("TotalBillingRp")
			SPVEmail_ = rsData("SupervisorEmail")
			CellPhoneBillRp_ = rsData("CellPhoneBillRp")
			CellPhonePrsBillRp_ = rsData("CellPhonePrsBillRp")
			ProgressDesc_ = rsData("ProgressDesc")

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
	
					& " <title>e-Billing Application</title> "_              
					& " <meta name='Microsoft Border' content='none, default'><style type='text/css'><!--.FontContent {font-family: verdana;font-size: 11px;color: black;}--></style> "_    
					& " </head><body bgcolor='#ffffff'> "_              
					& " <p><table cellspadding='1' cellspacing='0' width='80%' bgColor='white'>"_ 
					& "    <tr> "_           
					& "        <td colspan='6' align='center'><font face='Verdana, Arial, Helvetica' color='#999999' size='5'>eBilling System – Supervisor Reminder</font></td></tr> "_
					& "    <tr> "_       
					& "        <td colspan='6'>&nbsp; </td></tr> "_       
					& "    <tr> "_           
					& "        <td colspan='6' class='FontContent'> " & EmpName_  & " has submitted his/her cell phone billing and need you to approve it.</td></tr> "_
					& "    <tr> "_       
					& "        <td colspan='6'>&nbsp; </td></tr> "_   

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
		& "    		<td width='20%' class='FontContent'><b>Personal Amount (Kn.)</b></td> "_
		& "    	</tr> "

		if cdbl(CellPhoneBillRp_ ) > 0 Then

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    	<tr height=26> "_
		& "    	<td class='FontContent'>&nbsp;<a href='" & WebSiteAddress & "/BillingApproval.asp?EmpID=" & EmpID_ & "&Month=" & MonthP_ & "&Year=" & YearP_ & "&MobilePhone=" & MobilePhone_ &"' target='_blank'>Review and approve</a></td> "_
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
					& "        <td height=26 align='center' colspan='6' class='FontContent'>NOTE: This e-mail was automatically generated.</td> "_ 
					& "    </tr> "_       
		& "        <td colspan='6'>&nbsp; </td></tr> "_  
			
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