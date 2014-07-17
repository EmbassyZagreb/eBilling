<HTML>
<HEAD>
<!--#include file="connect.inc" -->
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 
<TITLE>U.S. Mission Jakarta e-Billing</TITLE>

<STYLE TYPE="text/css"><!--
  A:ACTIVE { color:#003399; font-size:8pt; font-family:Verdana; }
  A:HOVER { color:#003399; font-size:8pt; font-family:Verdana; }
  A:LINK { color:#003399; font-size:8pt; font-family:Verdana; }
  A:VISITED { color:#003399; font-size:8pt; font-family:Verdana; }
  body {scrollbar-3dlight-color:#FFFFFF; scrollbar-arrow-color:#E3DCD5; scrollbar-base-color:#FFFFFF; scrollbar-darkshadow-color:#FFFFFF;	scrollbar-face-color:#FFFFFF; scrollbar-highlight-color:#E3DCD5; scrollbar-shadow-color:#E3DCD5; }
  p { font-family: verdana; font-size: 12px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; color: #003399; text-decoration: none}
  h3 { font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 16px; font-style: normal; line-height: normal; font-weight: bold; color: #003399; letter-spacing: normal; word-spacing: normal; font-variant: small-caps}
  td { font-family: verdana; font-size: 10px; font-style: normal; font-weight: normal; color: #000000}
  .title { font-size:14px; font-weight:bold; color:#000080; }
  .SubTitle { font-size:16px; font-weight:bold; color:#000080;  }
  A.menu { text-decoration:none; font-weight:bold; }
  A.mmenu { text-decoration:none; color:#FFFFFF; font-weight:bold; }
  .normal { font-family:Verdana,Arial; color:black}
  .style5 {color: #FFFFFF;}
  .ActivePage {color: red; font-weight:bold; }
--></STYLE>
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
  <Center><FONT COLOR=#009900><B>SENSITIVE BUT UNCLASSIFIED</Center></FONT></B>
  <BR>
<CENTER>
  <IMG SRC="images/embassytitle2.jpeg" WIDTH="661" HEIGHT="80" BORDER="0"> 
  <TABLE WIDTH="65%" BORDER="0" CELLPADDING="0" CELLSPACING="0">
  <CAPTION><H3 STYLE="font-size:17px;color:#000040">Mission Jakarta - Billing Application</H3></CAPTION>
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Billing Notification</TD>
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
	send_from = "JakartaCustomerSeviceCenter@state.gov"
	Dim ObjMail
	Set ObjMail = Server.CreateObject("CDO.Message")
	Set objConfig = CreateObject("CDO.Configuration") 
	objConfig.Fields(cdoSendUsingMethod) = 2  
	objConfig.Fields(cdoSMTPServer) = "10.4.16.170" 
	objConfig.Fields.Update 
	Set objMail.Configuration = objConfig 
	noMail=0
	For Each loopIndex in Request("cbApproval")
		'response.write loopIndex & "<br>"
		X = len(loopIndex)
		'response.write X & "<br>"
		EmpID_ = Left(loopIndex, X-7)
		'response.write EmpID_ & "<br>"
		Period = mid(loopIndex,X-6,6)
		MonthP_ = left(Period,2)
		'response.write MonthP_ & "<br>"
		YearP_ = Right(Period,4)
		BillType_ = Right(loopIndex,1)
		'response.write YearP_ & "<br>"
		strsql = "Select * From vwMonthlyBilling Where EmpID='" & EmpID_ & "' And MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "'"
		'response.write BillType_ & "<Br>"  
		'response.write strsql & "<Br>"  
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
		end if
		
		'Send mail
		send_to = EmpEmail_ 
		'send_to = "kurniawane@state.gov"
		'response.write send_to
		objMail.From = send_from
		objMail.To = send_to 	
		objMail.Subject = "e-Billing System - Monthly Billing Reminder"
		objMail.HTMLBody = "<html><head>"
		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
	
		& " <title>e-Billing Application</title> "_              
		& " <meta name='Microsoft Border' content='none, default'><style type='text/css'><!--.FontContent {font-size: 12px;color: blue;}--></style> "_     
		& " </head><body bgcolor='#ffffff'> "_              
		& " <p><table cellspadding='1' cellspacing='0' width='80%' bgColor='white'>"_    
		& "    <tr> "_           
		& "        <td colspan='6' align='center'><u>Billing Period (Month - Year) : <a class='FontContent'>" & Period_ & "</a></u></td></tr> "_
		& "    <tr> "_           
		& "        <td colspan='6' align='Left'><u><b>Personal Info<b></u></td></tr> "_
		& "    <tr> "_           
		& "        <td width='20%'>Employee Name</td><td width='1%'>:</td><td class='FontContent'>" & EmpName_ & "</td><td>Agency / Office</td><td width='1%'>:</td><td class='FontContent'>" & Office_ & "</td></tr> "_
		& "    <tr> "_           
		& "        <td>Position</td><td width='1%'>:</td><td class='FontContent'>" & Position_ & "</td><td>Office Phone/Ext.</td><td width='1%'>:</td><td class='FontContent'>" & OfficePhone_ & "</td></tr> "_
		& "    <tr> "_           
		& "        <td>Homephone</td><td width='1%'>:</td><td class='FontContent' colspan='4'>" & HomePhone_ & "</td></tr> "_
		& "    <tr> "_ 
		& "        <td>Mobile Phone</td><td width='1%'>:</td><td class='FontContent'>" & MobilePhone_ & "</td><td>Exchange Rate</td><td width='1%'>:</td><td class='FontContent'>Rp." & FormatNumber(ExchangeRate_,0) & "/ Dollar</td></tr> "_
		& "    <tr> "_ 
		& "        <td colspan='6'><hr></td></tr> "_
		& "    <tr> "_ 
		& "        <td align='Left' colspan='6'><u><b>Billing detail :<b></u></td></tr> "_
		& "    <tr> "_ 
		& "        <td align='Left' colspan='6'><table cellspadding='1' border='1' bordercolor='black' cellspacing='0' width='100%' bgColor='white'><tr align='center'><td rowspan='2'><b>Type</b></td><td rowspan='2'><b>Billing (Rp.)</b></td><td colspan='2'><b>Should be paid</b></td></tr>"_
		& "    <tr> "_ 
		& "        <td align='center'><b>In Rupiah (Rp.)</b></td><td align='center'><b>In US Dollar ($)</b></td></tr> "
		if (cdbl(OfficePhoneBillRp_) > 0) And ((BillType_ ="X") or (BillType_ ="O")) Then 
			ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
			& "<tr> <td><a href='http://jakartaws01.eap.state.sbu/eBilling/OfficePhoneDetail.asp?Extension=" & OfficePhone_ & "&MonthP=" & MonthP_ & " &YearP=" & YearP_ & "' target='_blank'>Office Phone</a></td> "_
			& "	<td width='20%' class='FontContent' align='right'>" & formatnumber(OfficePhoneBillRp_,0) & "&nbsp;</td>"_
			& "	<td width='20%' class='FontContent' align='right'>" & formatnumber(OfficePhonePrsBillRp_ ,0) & "&nbsp;</td> "_
			& "	<td width='20%' class='FontContent' align='right'>" & formatnumber(OfficePhonePrsBillDlr_,2) & "&nbsp;</td> "_		
			& " </tr>"
		elseIf (BillType_ ="X") then
			ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
			& " <tr>"_
			& " 	<td>Office Phone</td>"_
			& " 	<td class='FontContent' align='right'>- &nbsp;</td>"_
			& " 	<td class='FontContent' align='right'>- &nbsp;</td>"_
			& " 	<td class='FontContent' align='right'>- &nbsp;</td>"_
			& " </tr>"
		end if
		if (cdbl(HomePhoneBillRp_) > 0) And ((BillType_ ="X") or (BillType_ ="H")) Then 
			ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
			& " <tr>"_
			& " <td><a href='http://jakartaws01.eap.state.sbu/eBilling/HomePhoneDetail.asp?HomePhone=" & HomePhone_ & "&MonthP=" & MonthP_ & "&YearP=" & YearP_ & "' target='_blank'>Home Phone</a></td> "_
			& " <td class='FontContent' align='right'>" & formatnumber(HomePhoneBillRp_ ,0) & "&nbsp;</td> "_
			& " <td class='FontContent' align='right'>" & formatnumber(HomePhonePrsBillRp_ ,0) & "&nbsp;</td> "_
			& " <td class='FontContent' align='right'>" & formatnumber(HomePhonePrsBillDlr_ ,2) & "&nbsp;</td> "_
			& " </tr>"
		elseIf (BillType_ ="X") then
			ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
			& " <tr>"_
			& " <td>Home Phone</td>"_
			& " <td class='FontContent' align='right'>- &nbsp;</td>"_
			& " <td class='FontContent' align='right'>- &nbsp;</td>"_
			& " <td class='FontContent' align='right'>- &nbsp;</td>"_
			& " </tr>"
		end if
		if (cdbl(CellPhoneBillRp_ ) > 0) And ((BillType_ ="X") or (BillType_ ="C")) Then 
			ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
			& " <tr>"_
				& " <td><a href='http://jakartaws01.eap.state.sbu/eBilling/CellPhoneDetail.asp?CellPhone=" & MobilePhone_ & "&MonthP="& MonthP_ & "&YearP="& YearP_ &"' target='_blank'>CellPhone</a></td>"_
				& " <td class='FontContent' align='right'>"& formatnumber(CellPhoneBillRp_  ,0) & "&nbsp;</td>"_
				& " <td class='FontContent' align='right'>"& formatnumber(CellPhonePrsBillRp_ ,0) & "&nbsp;</td>"_
				& " <td class='FontContent' align='right'>"& formatnumber(CellPhonePrsBillDlr_ ,2) & "&nbsp;</td>"_
			& " </tr>"
		elseIf (BillType_ ="X") then
			ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
			& " <tr>"_
				& " <td>Mobile Phone</td>"_
				& " <td class='FontContent' align='right'>- &nbsp;</td>"_
				& " <td class='FontContent' align='right'>- &nbsp;</td>"_
				& " <td class='FontContent' align='right'>- &nbsp;</td>"_
			& " </tr>"
		end if
		if (cdbl(TotalShuttleBillRp_) > 0) And ((BillType_ ="X") or (BillType_ ="S"))  Then
			ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
			& " <tr>"_
				& " <td><a href='http://jakartaws01.eap.state.sbu/eBilling/ShuttleBusBillDetail.asp?Username=" & user1_ & "&MonthP="& MonthP_ & "&YearP=" & YearP_ &"' target='_blank'>Shuttle Bus</a></td>"_
				& " <td class='FontContent' align='right'>"& formatnumber(TotalShuttleBillRp_ ,0) &"&nbsp;</td>"_
				& " <td class='FontContent' align='right'>"& formatnumber(TotalShuttleBillRp_ ,0) &"&nbsp;</td>"_
				& " <td class='FontContent' align='right'>"& formatnumber(TotalShuttleBillDlr_,2) &"&nbsp;</td>"_
			& " </tr>"
		elseIf (BillType_ ="X") then
			ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
			& " <tr>"_
				& " <td>Shuttle Bus</td>"_
				& " <td class='FontContent' align='right'>- &nbsp;</td>"_
				& " <td class='FontContent' align='right'>- &nbsp;</td>"_
				& " <td class='FontContent' align='right'>- &nbsp;</td>"_
			& " </tr>"
		end if
		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
		& "        </table></td></tr> "

		if (BillType_ ="X") Then
		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
		& "    <tr> "_ 
		& "        <td colspan='6'><table cellspadding='1' cellspacing='0' width='100%' bgColor='white' border='0'><tr><td align='center'><b>Total</b></td><td width='20%' class='FontContent' align='right'>&nbsp;</td><td width='20%' class='FontContent' align='right'><b><u>" & formatnumber(TotalBillingRp_ , 0) & "</u></b>&nbsp;</td><td width='20%' class='FontContent' align='right'><b><u>" & formatnumber(TotalBillingDlr_ ,2) & "</u></b>&nbsp;</td></tr></table></td></tr> "
		end if

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_	
		& "    <tr> "_       
		& "        <td colspan='6'>&nbsp; </td></tr> "_      

		& "    <tr> "_       
		& "        <td colspan='6'> "_
		& "        <p>For more details, please click on each billing type above</p></td> "_  

		& "    <tr> "_       
		& "        <td colspan='6'>&nbsp; </td></tr> "_      

		& "    <tr> "_       
		& "        <td colspan='6'>&nbsp; </td></tr> "_      

		& "    <tr> "_       
		& "        <td align='middle' colspan='6'>NOTE: This e-mail was automatically generated. Please do not respond to this e-mail address.</td> "_ 
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
	<td align="center"><br><a href="javascript:history.go(-1)"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
</tr>
</table>
</body> 
</html>