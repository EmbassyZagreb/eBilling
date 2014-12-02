<%@ Language=VBScript %>
<%
'Option Explicit
On Error Resume Next
%>

<!--#include file="connect.inc" -->


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
	<TITLE>U.S. Embassy Zagreb - zBilling Application</TITLE>
	<script src="jquery-latest.js" type="text/javascript"></script>
	<script src="jquery.tablesorter.js" type="text/javascript"></script>
	<script src="menu.js" type="text/javascript"></script>
	<link rel="stylesheet" type="text/css" href="style-left-nav.css" />
	<link rel="stylesheet" type="text/css" href="style-top-nav.css" />
	<link rel="stylesheet" type="text/css" href="style-template.css" />
	<link rel="stylesheet" type="text/css" href="style-graph.css" />
	<link rel="stylesheet" type="text/css" href="style-tablesorter.css" />
	<meta http-equiv="Content-Type" content="text/html; charset=Utf-8" />
<script type="text/javascript">
function checkall(obj)
{
	var c = document.frmCellPhonzBilling.elements.length
	for (var x=0; x<frmCellPhonzBilling.elements.length; x++)
	{
		cbElement = frmCellPhonzBilling.elements[x]
		if (cbElement.type == "checkbox")
		{
			cbElement.checked= obj.checked?true:false
		}
	}
}
</script>
	<script type="text/javascript">
	$(function() {
		$("#myTable").tablesorter({headers: { 5:{sorter: false}}, widgets: ['zebra']});
	});
	</script>
	<!--[if eq IE 8]>
   	<style type="text/css">
	div#navigation{
	   	position: absolute;
    		top: 80px;
    		left: 180px;
    		right: 0;
    		margin: 0 auto;
		}
	div#container {
  		margin-top: 65px
		}
   	</style>
   	<![endif]-->
</head>
<body>
<%
SentStatus_ = request("SentStatus")
if SentStatus_ = "" Then SentStatus_ = Request.Form("cmbSentStatus")
if SentStatus_ = "" then
	SentStatus_ = 1
end if

Func = Request("Func")
if isempty(Func) Then
	Func = 1
End if
Select Case Func
Case 1


Const IMAGES_PATH = "images/"
Const NrOfMonths = 12  'Number of months on the graph
Const GraphHeight = 100	'Height of the graph
Const BarWidth = 20


BarColorReSent = IMAGES_PATH & "aa0000ff.png" 'red
BarColorNotSent = IMAGES_PATH & "ffcc00ff.png" 'yellow
BarColorSent = IMAGES_PATH & "00aa00ff.png" 'green
TransparentPix = IMAGES_PATH & "00000000.png" 'transparent pixel


Dim rsPeriod, Period_, y, m, i, j
Dim iNotSent, iSent, iReSent
Dim iHeightNotSent, iHeightPersonal, iHeightReSent
Dim iStatus, iTotal

Dim user_
Dim rst
Dim strsql, arrResultSet, rs, rsempty, arrNumberList

user_ = request.servervariables("remote_user")

strsql = "Select Max(YearP+MonthP) As Period From vwMonthlyBilling"
set rsPeriod = server.createobject("adodb.recordset")
set rsPeriod = BillingCon.execute(strsql)
if not rsPeriod.eof Then
	Period_ = rsPeriod("Period")
end if


If Period_ <> "" Then
	curMonth_ = Right(Period_, 2)
	curYear_ = Left(Period_, 4)
Else
	curMonth_ = month(date())
	curYear_ = year(date())
End If

eYearP = curYear_
eMonthP = curMonth_
ePeriod = eYearP&eMonthP

sMonthP = Month(DateAdd("m", - NrOfMonths + 1, CDate(eMonthP& "/01/" &eYearP)))
If sMonthP < 10 Then sMonthP = "0" & CStr(sMonthP) Else sMonthP = CStr(sMonthP)
sYearP = CStr(Year(DateAdd("m", - NrOfMonths + 1, CDate(eMonthP& "/01/" &eYearP))))
sPeriod = sYearP&sMonthP










if len(curMonth_)= 1 then
	curMonth_ = "0" & curMonth_
end if

MonthP = Request("MonthP")
if MonthP ="" then
	MonthP = Request.Form("txtMonthP")
	if MonthP ="" then
		MonthP = curMonth_
	end if
end if

YearP = Request("YearP")
if YearP ="" then
	YearP = Request.Form("txtYearP")
	if YearP ="" then
		YearP = curYear_
	end if
end if





MobilePhone_ = trim(Request("CellPhone"))
'response.write "HomePhone_  :" & HomePhone_ & "<br>"
'MonthP_ = Request("MonthP")
'response.write MonthP_ & "<br>"
'YearP_ = Request("YearP")
'response.write YearP_ & "<br>"
AlternateEmailFlag_ = trim(Request("AlternateEmailFlag"))

SortBy_ = Request.Form("SortList")
'response.write "SortBy" & SortBy_
if (SortBy_ ="") then
	if Request("SortBy")<>"" then
		SortBy_ = Request("SortBy")
	Else
		SortBy_ = "DialedDatetime"
	end if
end if

Order_ = Request("OrderList")
if (Order_ ="") then
	if Request.Form("OrderList")<>"" then
		Order_ = Request.Form("OrderList")
	Else
		Order_ = "Asc"
	end if
end if





strsql = "Exec spNavigatorNotification '" & sPeriod & "','" & ePeriod & "','" & GraphHeight & "'"
set rs = server.createobject("adodb.recordset")
set rs = BillingCon.execute(strsql)

' Official, Personal, HeightOfficial, HeightPersonal, MonthP, YearP, ProgressId, AccumulatedDebt, HeightAccumulatedDebt
If NOT rs.EOF Then
	arrResultSet = rs.GetRows()
End If

'Close the connection with the database and free all database resources
'Set rs = Nothing
'BillingCon.Close
'Set BillingCon = Nothing


NotSend_ = 0
Send_ = 0
Resent_ = 0
ProgressStatus_ = "Bills Not Generated"
j = UBound(arrResultSet,2)
For i = 0 To j
	If (arrResultSet (6,i) = MonthP AND arrResultSet (7,i) = YearP) Then
		NotSend_ = arrResultSet (0,i)
		Send_ = arrResultSet (1,i)
		Resent_ = arrResultSet (2,i)
		ProgressStatus_ = "Bills Generated"
	End If
Next



%>


<div id="container">

	<div id="navigation">

						<form method="post" action="1SendNotification.asp" name="frmSendNotification"">
						<div class="selector_header">Sending Notifications<br><br></div>
						<div class="selector_title">Billing Period</div>
						<div class="selector_info"><%if ProgressStatus_ = "Bills Generated" Then %><%= MonthName(Cint(MonthP))%>&nbsp;<%= YearP%><%else%>- &nbsp;<%end if%></div>
						<div class="selector_title">Status</div>
						<div class="selector_info"><%=ProgressStatus_%></div>
						<div class="selector_title">Notifications Not Sent</div>
						<div class="selector_info"><%if ProgressStatus_ = "Bills Generated" Then %><%=NotSend_%><%else%>- &nbsp;<%end if%></div>
						<div class="selector_title">Notifications Sent</div>
						<div class="selector_info"><%if ProgressStatus_ = "Bills Generated" Then %><%=Send_%><%else%>- &nbsp;<%end if%></div>
						<div class="selector_title">Notifications Resent</div>
						<div class="selector_info"><%if ProgressStatus_ = "Bills Generated" Then %><%=Resent_%><%else%>- &nbsp;<%end if%></div>

						<%

							Response.Write "<table border=""0"" cellspacing=""0"" cellpadding=""0""  id=""chart3_table"">"
							Response.Write "<tr><td colspan=""" & (8)  & """ class=""selector_title"">Number of Sent / Not Sent Notifications</td><td colspan=""" & (4)  & """ class=""selector_graph_top""><img src=""" & IMAGES_PATH & "asc.gif" & """>" & eYearP & "<img src=""" & IMAGES_PATH & "desc.gif" & """></td></tr>"
							Response.Write "<tr>"

							j = 0
							For i = 0 To (NrOfMonths - 1)
								m = Month(DateAdd("m", i, CDate(sMonthP& "/01/" &sYearP)))
								y = Year(DateAdd("m", i, CDate(sMonthP& "/01/" &sYearP)))
								iMonth = MonthName(m ,True)
								iNotSent = ""
								iSent = ""
								iResent = ""
								iHeightNotSent = 0
								iHeightSent= 0
								iHeightReSent = 0
								iTotal = ""
								If (CInt(arrResultSet (6,j)) = m AND CInt(arrResultSet (7,j)) = y) Then
									iNotSent = CLng(arrResultSet (0,j))
									iSent = CLng(arrResultSet (1,j))
									iReSent = CLng(arrResultSet (2,j))
									iHeightNotSent = CLng(arrResultSet (3,j))
									iHeightSent= CLng(arrResultSet (4,j))
									iHeightReSent = CLng(arrResultSet (5,j))
									iTotal = iNotSent + iSent
									j = j + 1
								End If
								If m < 10 Then m = "0" & CStr(m) Else m = CStr(m)

								Response.Write "<td valign=""top"" class=""barcell"">"
								If iTotal <> "" Then
									Response.Write "<a href=""1SendNotification.asp?MonthP=" & m & "&YearP=" & y & "&SentStatus=" & SentStatus_ & """ style=""display:block; text-decoration: none;"">"
								Else
									Response.Write "<a href=""#"" style=""display:block; text-decoration: none;"">"
								End If
								Response.Write "<img src=""" & TransparentPix & """ width=""0"" height=""" & _
													GraphHeight - iHeightNotSent - iHeightSent& """ alt="""" title="""" />" & _
												"<br />" & _
												iTotal & "<br /><img src=""" & BarColorNotSent & """ width=""" & BarWidth & """ height=""" & _
													iHeightNotSent & """ alt="""" title=""" & iNotSent & """ />" & _
												"<br /><img src=""" & BarColorSent & """ width=""" & BarWidth & """ height=""" & _
													iHeightSent& """ alt="""" title=""" & iSent & """ />"
								If m = MonthP Then
									Response.Write "<div class=""chart3_labels_active"">" & iMonth & "</div>"
								Else
									Response.Write "<div class=""chart3_labels"">" & iMonth & "</div>"
								End If
								Response.Write "<div valign=""top"" class=""chart3_barcell_bottom""><img src=""" & BarColorReSent & """ width=""" & BarWidth & """ height=""" & _
													iHeightReSent & """ alt="""" title=""Resent"" />" & _
												"<br />" & iReSent & "<br><img src=""" & TransparentPix & """ width=""0"" height=""" & _
													GraphHeight - iHeightReSent & """ alt="""" title=""Resent"" /></div></a></td>"
							Next

							Response.Write "</tr>"
							Response.Write "<tr><td colspan=""" & (4) & """ class=""selector_graph_bottom"" align=""left"">About Graph</td><td colspan=""" & (8) & """ class=""selector_graph_bottom"" align=""right"">Number of Resent Notifications</td></tr>"
							Response.Write "</table>"
							%>

							<div class="selector_title">Sent Status</div>

						<%
							strsql ="select SendMailStatusID, SendMailStatusDesc from SendMailStatus"
							set SentStatusRS = server.createobject("adodb.recordset")
							set SentStatusRS = BillingCon.execute(strsql)
			'				response.write strStr
			%>
							<Select name="cmbSentStatus">
								<Option value='0'>--All--</Option>
			<%				Do While not SentStatusRS.eof %>
								<Option value='<%=SentStatusRS("SendMailStatusID")%>' <%if trim(SentStatus_) = trim(SentStatusRS("SendMailStatusID")) then %>Selected<%End If%> ><%=SentStatusRS("SendMailStatusDesc")%></Option>

			<%					SentStatusRS.MoveNext
							Loop%>
							</select>
							<div class="selector_title"> </div>

							<input type="submit" name="Submit" value="Search">
							<input type="hidden" name="txtMonthP" value='<%=MonthP%>' />
							<input type="hidden" name="txtYearP" value='<%=YearP%>' />
						</form>
	</div>

	<div id="wrapper">

		<div id="content">


		<%
if ProgressStatus_ = "Bills Generated" Then

%>
		<div class="details_header">USAGE DETAIL</div>
		<form method="post" action="1SendNotification.asp?Func=2" name="frmCellPhonzBilling">
		<table id="myTable" class="tablesorter">
		<thead>
		<tr>
		    <th>Name</th>
			<th>Section</th>
			<th>Number</th>
			<th>Sent Status</th>
			<th>Sent Date</th>
			<th>Check all
				<input type="checkbox" name="cbAll" value="true" onclick="checkall(this)" />
			</th>
		</tr>
		</thead>
		<tbody>
		<%



		strsql = "Select * From vwMonthlyBilling Where MonthP='" & MonthP & "' and YearP='" & YearP & "'"
		strFilter=""
		If SentStatus_ <> "0" then
			strFilter =strFilter & " and SendMailStatusID=" & SentStatus_
		End If
		strsql = strsql  & strFilter & " Order By EmpName Asc"
		set rsData = BillingCon.execute(strsql)

		no_ = 1
		separator_ = Chr(31)
		do while not rsData.eof
   			'if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4"
'			if (ProgressID_ = 4) then
			'if (cdbl(rsData("Cost")) <> cdbl(DetailRecordAmount_ )) then
		%>
			<tr>
			        <td>&nbsp;<%=rsData("EmpName")%></td>
		        	<td>&nbsp;<%=rsData("Office")%></td>
		        	<td>&nbsp;<%=rsData("MobilePhone")%></td>
		        	<td>&nbsp;<%=rsData("SendMailStatusDesc")%></td>
			        <td align="right"><%=rsData("SendMailDate")%></td>
					<td align="center">
			<%		If len(rsData("EmailAddress"))>5 then %>
						<Input type="Checkbox" name="cbApproval" Value='<%=rsData("EmailAddress")%><%=separator_%><%=rsData("MobilePhone")%><%=separator_%><%=rsData("MonthP")%><%=separator_%><%=rsData("YearP")%><%=separator_%><%=rsData("EmpName")%><%=separator_%><%=rsData("Office")%><%=separator_%><%=rsData("CellPhoneBillRp")%><%=separator_%><%=rsData("CellPhonePrsBillRp")%><%=separator_%><%=rsData("AlternateEmailFlag")%><%=separator_%><%=rsData("DummyFlag")%><%=separator_%><%=rsData("ProgressID")%><%=separator_%><%=rsData("ProgressDesc")%><%=separator_%><%=rsData("BillFlag")%>'>
			<%		Else%>
						&nbsp;
			<%		End If%>
					</td>
			</tr>
		<%     ' end if
			rsData.movenext
			no_ = no_ + 1
		loop
		%>
		</tbody>
		<%
		if ((ProgressID_< 4 and no_ >1) or (ProgressID_ = 4 and AlternateEmailFlag_="Y")) then%>
				<input type="submit" name="btnSubmit" Value="Send Notification(s)" />&nbsp;&nbsp;
				<input type="button" value="Cancel" onClick="javascript:location.href='1MonthlyBilling.asp?CellPhone=<%=MobilePhone_%>&MonthP=<%=MonthP%>&YearP=<%=YearP%>'">

				<input type="hidden" name="txtMobilePhone" value='<%=MobilePhone_ %>' />
				<input type="hidden" name="txtMonthP" value='<%=MonthP%>' />
				<input type="hidden" name="txtYearP" value='<%=YearP%>' />
				<input type="hidden" name="txtEmpID" value='<%=EmpID_ %>' />
				<input type="hidden" name="cmbSentStatus" value='<%=SentStatus_ %>' />
		<%end if%>
		</table>
		</form>
<%else%>
<table width="100%">
<tr>
	<td align="center">&nbsp;</td>
</tr>
<tr>
	<td align="center">there is not data.</td>
</tr>
</table>
<%end if%>

		</div>

	</div>

<!--#include file="1NavigationAlerts.asp" -->


<%
Case 2
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

strsql = " select * from PaymentDueDate"
set rst1 = server.createobject("adodb.recordset")
set rst1 = BillingCon.execute(strsql)
if not rst1.eof then
	CashierMinimumAmount_ = rst1("CashierMinimumAmount")
	CeilingAmount_ = rst1("CeilingAmount")
	DetailRecordAmount_ = rst1("DetailRecordAmount")
end if

'response.write "Test : " & Request("cbApproval")
If Request("cbApproval") <> "" then

	Set fso = CreateObject("Scripting.FileSystemObject")

	Dim send_from, send_to, send_cc, noMail, fileName, arrParams
	send_from = BillingDL

	fileName = "Files\BillingDetail.xls"

	Dim ObjMail
	noMail=0
	separator_ = Chr(31)
	For Each loopIndex in Request("cbApproval")
		'response.write loopIndex & "<br>"
	 	arrParams = Split(loopIndex, separator_)

	 	EmpEmail_ = arrParams(0)
	 	MobilePhone_  = arrParams(1)
   	 	MonthP_ = arrParams(2)
   	 	YearP_ = arrParams(3)
   	 	EmpName_ = arrParams(4)
		Office_	=  arrParams(5)
		CellPhoneBillRp_ = arrParams(6)
		CellPhonePrsBillRp_ = arrParams(7)
		AlternateEmailFlag_ = arrParams(8)
		DummyFlag_ = arrParams(9)
		ProgressID_ = arrParams(10)
		ProgressDesc_ = arrParams(11)
		BillFlag_ =  arrParams(12)

		Period_ = MonthP_ & " - " & YearP_

		if EmpEmail_ <>"" Then

			Set ObjMail = Server.CreateObject("CDO.Message")
			ObjMail.Configuration.Fields.Item _
				("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
			ObjMail.Configuration.Fields.Item _
				("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer
			ObjMail.Configuration.Fields.Item _
				("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
			ObjMail.Configuration.Fields.Update
			send_to = EmpEmail_
			'send_to = "zivkom@state.gov"
			'response.write send_to
			objMail.From = send_from
			objMail.To = send_to


			if ProgressID_ ="7" then
				objMail.Subject = "Info: zBilling System - No Invoice This Period"
				objMail.HTMLBody = "<html><head>"
				ObjMail.HTMLBody = ObjMail.HTMLBody & " "_

					& " <title>e-Billing Application</title> "_
					& " <meta name='Microsoft Border' content='none, default'><style type='text/css'><!--.FontContent {font-family: verdana;font-size: 11px;color: black;}--></style> "_
					& " </head><body bgcolor='#ffffff'> "_
					& " <p><table cellspadding='1' cellspacing='0' width='80%' bgColor='white'>"_
					& "    <tr> "_
					& "        <td colspan='6' align='center'><font face='Verdana, Arial, Helvetica' color='#999999' size='5'>zBilling System - No Invoice This Period</font></td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6'>&nbsp; </td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6' class='FontContent'>Your invoice has been processed and your usage did not meet the threshold to require review.  You have no further action.</td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6' class='FontContent'><br></td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6' class='FontContent'>Please reply if the phone number assignment below is not accurate.</td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6' class='FontContent'>&nbsp;</td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6' align='Left' class='FontContent'><u><strong>&nbsp;Personal Info:<strong></u></td></tr> "_
					& "    <tr> "_
					& "    <td colspan='6' align='Left'> "_
					& "    	<table cellspadding='1' border='2' bordercolor='black' cellspacing='3' width='100%' bgColor='#999999' border='0'>   "_
					& "    		<tr BGCOLOR='#999999'> "_
					& "    			<td colspan='3' style='border: none;' class='FontContent'><FONT color=#FFFFFF><strong>Employee Name : " & EmpName_ & "</strong></font></td> "_
					& "    			<td colspan='3' style='border: none;' align='right' class='FontContent'><FONT color=#FFFFFF><strong>Phone Number : " & MobilePhone_ & "&nbsp;</strong></font></td> "_
					& "    		</tr> "_
					& "    		<tr BGCOLOR='#999999'> "_
					& "    			<td colspan='6' style='border: none;' class='FontContent'><FONT color=#FFFFFF><strong>Agency / Office : " & Office_ & "</strong></font></td> "_
					& "    		</tr> "_
					& "    	</table></td></tr> " _
					& "    <tr> "_
					& "        <td align='Left' colspan='6' class='FontContent'><u><strong>&nbsp;Billing Detail:<strong></u></td></tr> "_
					& "    <tr> "_
					& "    <td align='Left' colspan='6'> "_
					& "    <table cellspadding='1' border='1' bordercolor='black' cellspacing='0' width='100%' bgColor='white'> "_
					& "    	<tr align='center' height=26> "_
					& "    		<td width='20%' class='FontContent'><strong>Action</strong></td> "_
					& "    		<td width='20%' class='FontContent'><strong>Billing Period</strong></td> "_
					& "    		<td width='20%' class='FontContent'><strong>Status</strong></td> "_
					& "    		<td width='20%' class='FontContent'><strong>Billing (Kn.)</strong></td> "_
					& "    		<td width='20%' class='FontContent'><strong>Personal Amount (Kn.)</strong></td> "_
					& "    	</tr> "

					if cdbl(CellPhoneBillRp_ ) > 0 Then

					ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
					& "    	<tr height=26> "_
					& "    	<td class='FontContent'>&nbsp;<a href='" & WebSiteAddress & "/1MonthlyBilling.asp?CellPhone=" & MobilePhone_ & "&MonthP=" & MonthP_ & "&YearP=" & YearP_ & "' target='_blank'>View Your Bill</a></td> "_
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
					& " </table></p>"_
					& "</body></html>"

			else

				objMail.Subject = "Action Required: zBilling System – Monthly Billing Notification"
'				objMail.Subject = "e-Billing System - Monthly Billing Reminder for period " & Period_
				objMail.HTMLBody = "<html><head>"

				if AlternateEmailFlag_ ="N" and DummyFlag_="N" Then
					ObjMail.HTMLBody = ObjMail.HTMLBody & " "_

					& " <title>e-Billing Application</title> "_
					& " <meta name='Microsoft Border' content='none, default'><style type='text/css'><!--.FontContent {font-family: verdana;font-size: 11px;color: black;}--></style> "_
					& " </head><body bgcolor='#ffffff'> "_
					& " <p><table cellspadding='1' cellspacing='0' width='80%' bgColor='white'>"_
					& "    <tr> "_
					& "        <td colspan='6' align='center'><font face='Verdana, Arial, Helvetica' color='#999999' size='5'>zBilling System – Monthly Billing Notification</font></td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6'>&nbsp; </td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6' class='FontContent'>Your invoice has been processed for this billing period and action is required.</td></tr> "

					if BillFlag_ = "P" Then

					ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
					& "    <tr> "_
					& "        <td colspan='6' class='FontContent'>Please follow the instructions below:</td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6' class='FontContent'>1) Click <a href='"& WebSiteAddress & "/1MonthlyBilling.asp?CellPhone=" & MobilePhone_ & "&MonthP=" & MonthP_ & "&YearP=" & YearP_ &"' target='_blank'>here </a> to access the zBilling application.</td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6' class='FontContent'>2) In the application, this cell phone is registered as a personal one.</td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6' class='FontContent'>3) Proceed with the payment if your total accumulated debt is greater than " & CashierMinimumAmount_ & " Kuna.</td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6' class='FontContent'><br></td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6' class='FontContent'>Please reply if the cell phone is approved for official use.</td></tr> "

					else

					ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
					& "    <tr> "_
					& "        <td colspan='6' class='FontContent'><strong>Do NOT</strong> make a payment yet - Please follow the instructions below:</td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6' class='FontContent'>1) Click <a href='"& WebSiteAddress & "/1MonthlyBilling.asp?CellPhone=" & MobilePhone_ & "&MonthP=" & MonthP_ & "&YearP=" & YearP_ &"' target='_blank'>here </a> to access the zBilling application.</td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6' class='FontContent'>2) Uncheck any official calls.</td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6' class='FontContent'>3) Click ""update"" to subtotal remaining personal calls.</td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6' class='FontContent'>4) Submit your invoice to your supervisor for approval.</td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6' class='FontContent'>5) Make payment if necessary - only <strong>AFTER</strong> your supervisor has approved.  You will receive a confirmation email informing you if you need to make a payment.</td></tr> "

					end if

					ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
					& "    <tr> "_
					& "        <td colspan='6' align='center'>&nbsp;</td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6' align='Left' class='FontContent'><u><strong>&nbsp;Personal Info:<strong></u></td></tr> "_
					& "    <tr> "_
					& "    <td colspan='6' align='Left'> "_
					& "    	<table cellspadding='1' border='2' bordercolor='black' cellspacing='3' width='100%' bgColor='#999999' border='0'>   "_
					& "    		<tr BGCOLOR='#999999'> "_
					& "    			<td colspan='3' style='border: none;' class='FontContent'><FONT color=#FFFFFF><strong>Employee Name : " & EmpName_ & "</strong></font></td> "_
					& "    			<td colspan='3' style='border: none;' align='right' class='FontContent'><FONT color=#FFFFFF><strong>Phone Number : " & MobilePhone_ & "&nbsp;</strong></font></td> "_
					& "    		</tr> "_
					& "    		<tr BGCOLOR='#999999'> "_
					& "    			<td colspan='6' style='border: none;' class='FontContent'><FONT color=#FFFFFF><strong>Agency / Office : " & Office_ & "</strong></font></td> "_
					& "    		</tr> "_
					& "    	</table></td></tr> " _
					& "    <tr> "_
					& "        <td align='Left' colspan='6' class='FontContent'><u><strong>&nbsp;Billing Detail:<strong></u></td></tr> "_
					& "    <tr> "_
					& "    <td align='Left' colspan='6'> "_
					& "    <table cellspadding='1' border='1' bordercolor='black' cellspacing='0' width='100%' bgColor='white'> "_
					& "    	<tr align='center' height=26> "_
					& "    		<td width='20%' class='FontContent'><strong>Action</strong></td> "_
					& "    		<td width='20%' class='FontContent'><strong>Billing Period</strong></td> "_
					& "    		<td width='20%' class='FontContent'><strong>Status</strong></td> "_
					& "    		<td width='20%' class='FontContent'><strong>Billing (Kn.)</strong></td> "_
					& "    		<td width='20%' class='FontContent'><strong>Personal Amount (Kn.)</strong></td> "_
					& "    	</tr> "

					if cdbl(CellPhoneBillRp_ ) > 0 Then

					ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
					& "    	<tr height=26> "_
					& "    	<td class='FontContent'>&nbsp;<a href='" & WebSiteAddress & "/1MonthlyBilling.asp?CellPhone=" & MobilePhone_ & "&MonthP=" & MonthP_ & "&YearP=" & YearP_ & "' target='_blank'>Review your invoice</a></td> "_
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

				else
					If fso.FileExists (fileName) THEN
						set objFile = fso.GetFile (fileName)
						objFile.Delete
					end If

					Set objFile = fso.CreateTextFile(Server.MapPath(fileName))

					objFile.Writeline "<HTML>"
					objFile.Writeline "<HEAD><TITLE>Billing</TITLE>"
					objFile.Writeline "<style type='text/css'>"
					objFile.Writeline "<!--"
					objFile.Writeline ".style4 {color: #FFFFFF; font-weight: bold;}"
					objFile.Writeline ".smallfont{font-size: x-small;}"
					objFile.Writeline "-->"
					objFile.Writeline "</style>"
					objFile.Writeline "</HEAD>"
					objFile.Writeline "<BODY>"
					objFile.Writeline "<form>"
					objFile.Writeline "   <table cellpadding='0' cellspacing='0' border='0' width='100%'>"
					objFile.Writeline "      <tr>"
					objFile.Writeline "		<td><strong>Personal Usage Detail for Period <strong>" & MonthP_ & " - " & YearP_ & "</strong> :</strong></td>"
					objFile.Writeline "      </tr>"
					objFile.Writeline "     <tr>"
					objFile.Writeline "		<td>&nbsp;</td>"
					objFile.Writeline "     </tr>"

					objFile.Writeline "     <tr>"
					objFile.Writeline "     	<td align='Center'>"
					objFile.Writeline "     		<table cellspadding='0' cellspacing='0' bordercolor='black' border='1' width='90%' bgColor='white'>"
					objFile.Writeline "     		<tr align='center' cellpadding='0' cellspacing='0'>"
					objFile.Writeline "     			<TD width='5%'><strong>No</strong></TD>"
					objFile.Writeline "     		     	<TD><strong>Dialed Date/time</strong></TD>"
					objFile.Writeline "     			<TD width='20%'><strong>Dialed Number</strong></TD>"
					objFile.Writeline "     			<TD><strong>Call Type</strong></TD>"
					objFile.Writeline "     			<TD><strong>Duration</strong></TD>"
					objFile.Writeline "     			<TD width='10%'><strong>Amount (Kn)</strong></TD>"
					objFile.Writeline "     		</tr>"


								'strsql = "Select * from CellPhoneDt Where PhoneNumber='" & MobilePhone_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "' and cost>" & DetailRecordAmount_
								strsql = "Select * from CellPhoneDt Where PhoneNumber='" & MobilePhone_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "' and cost>'" & DetailRecordAmount_ & "' Order by DialedDatetime Asc"
								'response.write strsql & "<br>"
								set rsCellPhone = BillingCon.execute(strsql)
								No_ = 1
								do while not rsCellPhone.eof
   								if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4"
					objFile.Writeline "     		<tr bgcolor='" & bg & "'>"
					objFile.Writeline "     			<td align='right'>" & No_ & "&nbsp;</td>"
					objFile.Writeline "     			<td><FONT color=#330099 size=2>&nbsp;" & rsCellPhone("DialedDatetime") & "</font></td>"
					objFile.Writeline "     			<td><FONT color=#330099 size=2>&nbsp;" & rsCellPhone("DialedNumber") & "</font></td>"
					objFile.Writeline "     			<td><FONT color=#330099 size=2>&nbsp;" & rsCellPhone("CallType") & "</font></td>"
					objFile.Writeline "     			<td><FONT color=#330099 size=2>&nbsp;" & rsCellPhone("CallDuration") & "</font></td>"
					objFile.Writeline "     			<td align='right'><FONT color=#330099 size=2>" & formatnumber(rsCellPhone("Cost"),-1) & "</font></td>"
					objFile.Writeline "			</tr>"

								rsCellPhone.movenext
								No_ = No_ + 1
								loop
					objFile.Writeline "     		<tr>"
					objFile.Writeline "     			<td align='center' colspan='5'><strong>Total (Kn.) </strong>&nbsp;</td>"
					objFile.Writeline "				<td align='right'><strong><u>" & formatnumber(CellPhonePrsBillRp_ ,-1) & "</u></strong>&nbsp;</td>"
					objFile.Writeline "			</tr>"
					objFile.Writeline "			</table>"
					objFile.Writeline "		</td>"
					objFile.Writeline "	</tr>"
					objFile.Writeline "	</table>"
					objFile.Writeline "</form>"
					objFile.Writeline "</BODY>"
					objFile.Writeline "</HTML>"
					objFile.close

					'response.write MobilePhone_ & MonthP_ & YearP_
					strsql = "Select * From vwCellphoneHd Where PhoneNumber='" & MobilePhone_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "'"
					'response.write strsql & "<br>"
					set rsCellPhone = BillingCon.execute(strsql)
					if not rsCellPhone.eof then
						PreviousBalance_= rsCellPhone("PreviousBalance")
						Payment_= rsCellPhone("Payment")
						Adjustment_= rsCellPhone("Adjustment")
						BalanceDue_= rsCellPhone("BalanceDue")
						SubscriptionFee_= rsCellPhone("SubscriptionFee")
						LocalCall_= rsCellPhone("LocalCall")
						Interlocal_= rsCellPhone("SLJJ")
						IDD_= rsCellPhone("SLI")
						SMS_= rsCellPhone("SMS")
						IRL_= rsCellPhone("IRL")
						Prepaid_= rsCellPhone("Prepaid")
						FARIDA_= rsCellPhone("FARIDA")
						MobileBanking_= rsCellPhone("MobileBanking")
						DetailedCallRecord_= rsCellPhone("DetailedCallRecord")
						GPRS_= rsCellPhone("GPRS")
						IPHONE_= rsCellPhone("IPHONE")
						'FARIDA_= rsCellPhone("FARIDA")
						'DataRoam_= rsCellPhone("DataRoam")
						MinUsage_= rsCellPhone("MinUsage")
						DiskonBicara_= rsCellPhone("DiskonBicara")
						GPRS_= rsCellPhone("GPRS")
						DiskonSMS_= rsCellPhone("DiskonSMS")
						DiskonGPRS_= rsCellPhone("DiskonGPRS")
						DiskonMMS_= rsCellPhone("DiskonMMS")
						DiskonPenggunaan_= rsCellPhone("DiskonPenggunaan")
						SubTotalTKP_= rsCellPhone("SubTotalTKP")
						SubTotalKP_= rsCellPhone("SubTotalKP")
						PPN_= rsCellPhone("PPN")
						StampFee_= rsCellPhone("StampFee")
						CurrentBalance_= rsCellPhone("CurrentBalance")
						Total_= rsCellPhone("Total")
					end if
					'response.write Total_
					ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
					& " <title>e-Billing Application</title> "_
					& " <meta name='Microsoft Border' content='none, default'><style type='text/css'><!--.FontContent {font-size: 12px;color: blue;}--></style> "_
					& " </head><body bgcolor='#ffffff'> "_
					& " <p><table cellspadding='1' cellspacing='0' width='80%' bgColor='white'>"_
					& "    <tr> "_
					& "        <td colspan='6'>You have received this email because you have been identified as the supervisor for the phone number below, which is assigned to a group, an employee without open-net access, or an employee who is responsible for multiple phones.  Please follow the instructions :</td> "_
					& "    </tr> "_
					& "    <tr> "_
					& "        <td colspan='6'>1) Review the summary below and the Usage Detail in the attached MS Excel file.</td> "_
					& "    </tr> "_
					& "    <tr> "_
					& "        <td colspan='6'>2) Work with the users of the phone to determine if any of the calls are personal.</td> "_
					& "    </tr> "_
					& "    <tr> "_
					& "        <td colspan='6'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;a.If personal calls amount to <strong>less than or equal to " & formatnumber(CeilingAmount_,-1) & " kuna</strong>, reply to this email and write “No Payment”.</td> "_
					& "    </tr> "_
					& "    <tr> "_
					& "        <td colspan='6'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;in the email content.</td> "_
					& "    </tr> "_
					& "    <tr> "_
					& "        <td colspan='6'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;b.If the personal calls amount to <strong>more than " & formatnumber(CeilingAmount_,-1) & " kuna</strong>, print this email, write in the personal call amount</td> "_
					& "    </tr> "_
					& "    <tr> "_
					& "        <td colspan='6'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;at the bottom, sign at the bottom, and instruct the employee to make payment with the cashier.</td> "_
					& "    </tr> "_
					& "    <tr> "_
					& "        <td colspan='6'>&nbsp; </td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6' align='center'><u>Billing Period (Month - Year) : <a class='FontContent'>" & Period_ & "</a></u></td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6' align='Left'><u><strong>Personal Info<strong></u></td></tr> "_
					& "    <tr> "_
					& "        <td width='20%'>Employee Name</td><td width='1%'>:</td><td class='FontContent'>" & EmpName_ & "</td><td>Agency / Office</td><td width='1%'>:</td><td class='FontContent'>" & Office_ & "</td></tr> "_
					& "    <tr> "_
					& "        <td>Position</td><td width='1%'>:</td><td class='FontContent' colspan='4'>" & Position_ & "</td></tr> "_
					& "    <tr> "_
					& "        <td>Mobile Phone</td><td width='1%'>:</td><td class='FontContent' colspan='4'>" & MobilePhone_ & "</td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6'><hr></td></tr> "_
					& "    <tr> "_
					& "        <td align='Left' colspan='6'><u><strong>Billing Detail :<strong></u></td></tr> "_
					& "    <tr> "_
					& "        <td align='Left' colspan='6'> "_
					& "		<table cellspadding='0' border='1' bordercolor='black' cellspacing='0' width='100%' bgColor='white'> "_
					& "		<tr><td colspan='4' align='center' class='SubTitle'>USAGE SUMMARY</td></tr> "_
					& "		<tr><td colspan='4'>&nbsp;<u><strong>Monthly Fees</strong> / <i>Mjesecne pretplate:<i/></u></td></tr>"_
					& "		<tr><td colspan='4'><table cellspadding='0' cellspacing='0' bordercolor='black' width='100%' bgColor='white'>"_
					& "		    <tr><td width='70%'>&nbsp;<strong>Subscription Monthly Fee</strong> / <i>Mjesecna naknada za pretplatnicki broj<i/></td><td width='3%'>&nbsp;Kn.</td><td align='right'>" & formatnumber(SubscriptionFee_,-1) & "&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<strong>Data Monthly Fee</strong> / <i>Mjesecna naknada za mobilni prijenos podataka<i/></td><td>&nbsp;Kn.</td><td align='right'>"& formatnumber(FARIDA_,-1) & "&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<strong>Other Charges</strong> / <i>Ostale usluge<i/></td><td>&nbsp;Kn.</td><td align='right'>"& formatnumber(DetailedCallRecord_,-1) & "&nbsp;&nbsp;&nbsp;</td></tr></table></td></tr>"_
					& "		<tr><td colspan='4'>&nbsp;<u><strong>Usage Charges</strong> / <i>Pozivi i prijenos podataka:<i/></u></td></tr>"_
					& "		<tr><td colspan='4'><table cellspadding='0' cellspacing='0' bordercolor='black' width='100%' bgColor='white'>"_
					& "		    <tr><td>&nbsp;<strong>VPN Network Calls</strong> / <i>Pozivi unutar VPN mreže<i/></td><td>&nbsp;Kn.</td><td align='right'>"& formatnumber(LocalCall_,-1) &"&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<strong>Calls to VIP Network</strong> / <i>Pozivi prema VIP mobilnoj mreži<i/></td><td>&nbsp;Kn.</td><td align='right'>"& formatnumber(BalanceDue_,-1) &"&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<strong>Calls to Landlines in Croatia</strong> / <i>Pozivi prema fiksnim mrežama u Hrvatskoj<i/></td><td>&nbsp;Kn.</td><td align='right'>"& formatnumber(Interlocal_,-1) & "&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<strong>Calls to Other Mobile Networks</strong> / <i>Pozivi prema ostalim mobilnim mrežama<i/></td><td>&nbsp;Kn.</td><td align='right'>"& formatnumber(IDD_,-1) & "&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<strong>SMS</strong> / <i>SMS poruke<i/></td><td>&nbsp;Kn.</td><td align='right'>"& formatnumber(SMS_,-1) & "&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<strong>MMS</strong> / <i>MMS Poruke<i/></td><td>&nbsp;Kn.</td><td align='right'>"& formatnumber(GPRS_,-1) & "&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<strong>International Calls from Croatia</strong> / <i>Medunarodni pozivi iz Hrvatske<i/></td><td>&nbsp;Kn.</td><td align='right'>"& formatnumber(IRL_,-1) & "&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<strong>Incoming Calls in Roaming</strong> / <i>Dolazni pozivi u roamingu<i/></td><td>&nbsp;Kn.</td><td align='right'>"& formatnumber(PreviousBalance_,-1) &"&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<strong>Outgoing Calls in Roaming</strong> / <i>Odlazni pozivi u roamingu<i/></td><td>&nbsp;Kn.</td><td align='right'>"& formatnumber(Adjustment_,-1) &"&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<strong>GPRS/EDGE/UMTS Data Transfer</strong> / <i>GPRS/EDGE/UMTS prijenos podataka<i/></td><td>&nbsp;Kn.</td><td align='right'>"& formatnumber(IPHONE_,-1) &"&nbsp;&nbsp;&nbsp;</td></tr></table></td></tr>"_
					& "		    <tr><td colspan='4'><table cellspadding='0' cellspacing='0' bordercolor='black' width='100%' bgColor='white'>"_
					& "		    <tr><td width='70%'>&nbsp;<strong>Neto Total</strong> / <i>Neto Total<i/></td><td width='3%'>&nbsp;Kn.</td><td align='right'>"& formatnumber(Payment_,-1) &"&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<strong>VAT</strong> / <i>PDV<i/></td><td>&nbsp;Kn.</td><td align='right'>"& formatnumber(PPN_,-1) &"&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<strong>Services Exempted from VAT</strong> / <i>Usluge na koje se ne obracunava PDV<i/></td><td>&nbsp;Kn.</td><td align='right'>"& formatnumber(StampFee_,-1) &"&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		    <tr><td>&nbsp;<strong>Grand Total</strong> / <i>Bruto Total<i/></td><td>&nbsp;Kn.</td><td align='right'>"& formatnumber(CurrentBalance_,-1)&"&nbsp;&nbsp;&nbsp;</td></tr>"_
					& "		   </table></td></tr>"_
					& "        </table></td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6'>&nbsp; </td></tr> "_
					& "    <tr> "_
					& "        <td colspan='6'>&nbsp; </td></tr> "_

					& "    <tr> "_
					& "        <td colspan='6'>&nbsp;&nbsp;Amount to be paid for personal calls: _______________________  Supervisor Signature:________________________</td> "_
					& "    </tr> "_
					& " </table></p>"_
					& "</body></html>"

					ObjMail.AddAttachment Server.MapPath(fileName)
				end if
			end if

'response.write ObjMail.HTMLBody
				objMail.Send

'			strsql = "Execute spSendNotificationUpdate '" & EmpID_ & "','" & MonthP_ & "','" & YearP_ & "','" & AlternateEmailFlag_ & "'"
			strsql = "Execute spSendNotificationUpdate '" & MobilePhone_ & "','" & MonthP_ & "','" & YearP_ & "'"
			'response.write strsql & "<Br>"
			set rsData = server.createobject("adodb.recordset")
			set rsData = BillingCon.execute(strsql)

			noMail=noMail+1
		end if
	next
	Set objMail = Nothing
	Set objConfig = Nothing
End If

	Response.AddHeader "REFRESH","0;URL=1SendNotification.asp?CellPhone=" & MobilePhone_ & "&MonthP=" & MonthP_ & "&YearP=" & YearP_ & "&SentStatus=" & SentStatus_ & ""

End Select
%>
</BODY>
</html>
