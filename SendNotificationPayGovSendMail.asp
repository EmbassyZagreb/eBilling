<HTML>
<HEAD>
<!--#include file="connect.inc" -->
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 
<TITLE>U.S. Embassy Zagreb - zBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
   <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Billing Notification - Pay.gov</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<%
sMonthP_ = request.form("txtsMonthP")
sYearP_ =  request.form("txtsYearP")
eMonthP_ =  request.form("txteMonthP")
eYearP_ =  request.form("txteYearP")
ProgressID_ =  request.form("txtProgressID")

rbPeriod_ = request.form("txtPeriod")

'response.write "rbPeriod_  :" & rbPeriod_ 
if rbPeriod_ ="S" then
	sPeriod_ = sYearP_&sMonthP_
	ePeriod_ = eYearP_&eMonthP_
End if

curMonth_ = month(date())
curYear_ = year(date())
if len(curMonth_)= 1 then
	curMonth_ = "0" & curMonth_
end if

Agency_ =  request.form("txtAgency")
Section_ =  request.form("txtSection")
SectionGroup_ =  request.form("txtSectionGroup")
EmpID_ =  request.form("txtEmpID")
EmailAddress_ =  request.form("txtEmailAddress")
BillType_=  request.form("txtBillType")

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
	Dim send_from, send_to, send_cc, noMail, fileName
	send_from = BillingDL
	Dim ObjMail
	noMail=0
	For Each loopIndex in Request("cbApproval")
		'response.write loopIndex & "<br>"
		X = len(loopIndex)
		'response.write X & "<br>"
'		EmpID_ = Left(loopIndex, X-1)
		EmpID_ = loopIndex
		'response.write EmpID_ & "<br>"
		'BillType_ = Right(loopIndex,1)
		'SendStatus_ = Right(loopIndex,1)

		'response.write YearP_ & "<br>"
		if rbPeriod_ = "X" then
'			strsql = "Select * From vwMonthlyBilling Where ProgressID in(4,8) And EmpID='" & EmpID_ & "' And (ProgressID=" & ProgressID_ & " or " & ProgressID_ & " =0) and SendMailStatusID=" & SendStatus_ 
			strsql = "Select * From vwMonthlyBilling Where ProgressID in(4,8) And EmpID='" & EmpID_ & "' And (ProgressID=" & ProgressID_ & " or " & ProgressID_ & " =0)"
		else
			strsql = "Select * From vwMonthlyBilling Where ProgressID in(4,8) And EmpID='" & EmpID_ & "' And (ProgressID=" & ProgressID_ & " or " & ProgressID_ & " =0) and YearP+MonthP>='" & sPeriod_ & "' and YearP+MonthP<='" & ePeriod_ & "'"
		end if	
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

		if EmpEmail_ <>"" Then
			If Right(EmpID_,1)="N" then			
				LoginID_ = EmpID_
			End If
			Dim objBP
			Set ObjMail = Server.CreateObject("CDO.Message")
			Set objConfig = CreateObject("CDO.Configuration") 
			objConfig.Fields(cdoSendUsingMethod) = 2  
			objConfig.Fields(cdoSMTPServer) = SMTPServer
			objConfig.Fields.Update 
			Set objMail.Configuration = objConfig 
			'Send mail
			send_to = EmpEmail_ 
			'send_to = "kurniawane@state.gov"
			'response.write send_to

			'ObjMail.MimeFormatted = True
			objMail.From = send_from
			objMail.To = send_to 	

			objMail.Subject = "Action Required: zBilling System – Monthly Billing Reminder"
			objMail.HTMLBody = "<html><head>"
			ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
			& " <title>e-Billing Application</title> "_              
			& " <meta name='Microsoft Border' content='none, default'><style type='text/css'><!--.FontContent {font-size: 12px;color: blue;}--></style> "_     
			& " </head><body bgcolor='#ffffff'><center> "_              
			& " <IMG SRC='"& WebSiteAddress & "/images/embassytitle2.jpeg' WIDTH='100%' HEIGHT='80' BORDER='0'><br>"_
			& " <p><table cellspadding='1' cellspacing='0' width='100%' bgColor='white'>"_ 
			& " <tr><td colspan='2' width='50%'><table cellspacing='0' border='1' bordercolor='black'>"_
						& " <tr><td colspan='2' align='Center'><strong>Payee Info</strong></td></tr>"_
						& " <tr><td align='right'><i>Employee Name : </i></td><td>" & EmpName_ & "</td></tr>"_
						& " <tr><td align='right'><i>Department : </i></td><td>" & Office_ & "</td></tr></table></td> "_
				& "<td colspan='4' width='50%' align='right'><table cellspacing='0'>"_
					& " <tr><td colspan='4'><strong><font color='red' size='5'>Bill of Collection</font></td></tr>"_
					& " <tr><td colspan='2'><i>Bill of Collection Number : </i></td><td colspan='2'>" & LoginID_ & curMonth_ & curYear_ &"</td></tr>"_
					& " <tr><td colspan='2'><i>Bill of Collection Date : </i></td><td colspan='2'>" & Date() & "</td></tr>"_
			& " </table></td></tr>"_
			& "<tr><td colspan='6'>&nbsp;</td></tr>"_
			& "<tr><td colspan='6'>&nbsp;</td></tr>"_
			& "<tr><td colspan='6'><table cellspacing='0' width='100%' border='1' bordercolor='black'>"_
				& "<tr align='center'><td><strong>Billing Period</strong></td><td><strong>Bill Type</strong></td><td><strong>Description</strong></td><td><strong>Amount Due (USD)</strong></td><td><strong>Exchange Rate</strong></td><td><strong>Amount Due (Kn)</strong></td></tr>"
			TotalBillRp_ = 0
			TotalBillDlr_ = 0
			 Do while not rsData.eof
				TotalBillRp_ = cdbl(TotalBillRp_) + cdbl(rsData("TotalBillingAmountPrsRp"))
				TotalBillDlr_ = cdbl(TotalBillDlr_) + cdbl(rsData("TotalBillingAmountPrsDlr"))
				ObjMail.HTMLBody = ObjMail.HTMLBody & "<tr><td>&nbsp;" & rsData("YearP") & "-" & rsData("MonthP") & "</td><td>Mobile Phone</td><td>&nbsp;" & rsData("MobilePhone") & "</td><td align='right'>$ " & formatnumber(cdbl(rsData("TotalBillingAmountPrsDlr")),-1) & "&nbsp;</td><td align='right'>" & formatnumber(rsData("ExchangeRate"),-1) & "&nbsp;</td><td align='right'>" & formatnumber(cdbl(rsData("TotalBillingAmountPrsRp")),-1) & "&nbsp;</td></tr>"
			 	rsData.movenext
			 Loop 
			ObjMail.HTMLBody = ObjMail.HTMLBody & "<tr><td colspan='3' align='right'><strong>Total</strong>&nbsp;</td><td align='right'><strong>$ " & formatnumber(TotalBillDlr_,-1) & "</strong>&nbsp;</td><td>&nbsp;</td><td align='right'><strong>" & formatnumber(TotalBillRp_,-1) & "</strong>&nbsp;</td></tr></table></td></tr>"_			
			& " </table><div align='right'><i>*Please remit payment within 15 days of the Invoice Date</i></div>"_
			& "<br><br><br><div align='left'><strong>Payment Options</strong></div>"_
			& "<div align='left'>Pay at the embassy cashier:</div>"_
			& "<div align='left'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & CashierInfo & "</div>"_
			& "<br><br><br><div align='left'><strong>Fund Cite Details (For Cashier)</strong></div>"_
			& "<div align='left'>&nbsp;&nbsp;&nbsp;$ " & formatnumber(TotalBillDlr_,-1) & " 19-X45190001-00-5306-53063R3221-6195-2322 (ICASS Mobile Phone)</div>"_
			& "<br><br><br><br><br><div align='left'>Please contact <a href='mailto:zgbphonebill@state.gov'>zgbphonebill@state.gov</a> with any questions</div>"_
			& "</p></body></html>"
			objMail.Send

			strsql = "Execute spSendNotificationGroupUpdate '" & EmpID_ & "'"
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