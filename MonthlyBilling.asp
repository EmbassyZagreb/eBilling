<%@ Language=VBScript %>
<%
'Option Explicit
On Error Resume Next
%>
'test comment
<!--#include file="connect.inc" -->


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
	<TITLE>U.S. Embassy Zagreb - zBilling rocks!Application</TITLE>
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
	var c = document.frmCellPhoneBilling.elements.length
	for (var x=0; x<frmCellPhoneBilling.elements.length; x++)
	{
		cbElement = frmCellPhoneBilling.elements[x]
		if (cbElement.type == "checkbox")
		{
			cbElement.checked= obj.checked?true:false
		}
	}
}

function validate_form()
{
	valid = true;
	msg = ""

	if (document.frmMonthlyBilling.cmbSupervisor.value == "" )
	{
		msg = "Please fill in your supervisor mail !!!\n"
		valid = false;
	}

	if (valid == false)
	{
		alert(msg)
	}
	return valid;
}
function show(ele) {
         var srcElement = document.getElementById(ele);
         if(srcElement != null) {
	   if(srcElement.style.display == "block") {
     		  srcElement.style.display= 'none';
   	    }
            else {
                   srcElement.style.display='block';
            }
            return false;
       }
  }

EnableSubmit = function(val)
{
    var sbmt = document.getElementById("Accept");

    if (val.checked == true)
    {
        sbmt.disabled = true;
    }
    else
    if (val.checked == false)
    {
        sbmt.disabled = true;
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
Dim user_
user_ = request.servervariables("remote_user")

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

Dim m_arrBarColor (2,7)
'// Official (Status: 0-Pending, 1-Waiting Approval from Supervisor, 2-Need Correction, etc)
m_arrBarColor (0,0) = IMAGES_PATH & "aa0000ff.png"
m_arrBarColor (0,1) = IMAGES_PATH & "ffcc00ff.png"
m_arrBarColor (0,2) = IMAGES_PATH & "aa0000ff.png"
m_arrBarColor (0,3) = IMAGES_PATH & "00aa00ff.png"
m_arrBarColor (0,4) = IMAGES_PATH & "00aa00ff.png"
m_arrBarColor (0,5) = IMAGES_PATH & "555555ff.png"
m_arrBarColor (0,6) = IMAGES_PATH & "555555ff.png"
m_arrBarColor (0,7) = IMAGES_PATH & "ffcc00ff.png"

'// Personal (Status: 0-Pending, 1-Waiting Approval from Supervisor, 2-Need Correction, etc)
m_arrBarColor (1,0) = IMAGES_PATH & "aa0000ffstriped.png"
m_arrBarColor (1,1) = IMAGES_PATH & "ffcc00ffstriped.png"
m_arrBarColor (1,2) = IMAGES_PATH & "aa0000ffstriped.png"
m_arrBarColor (1,3) = IMAGES_PATH & "00aa00ffstriped.png"
m_arrBarColor (1,4) = IMAGES_PATH & "00aa00ffstriped.png"
m_arrBarColor (1,5) = IMAGES_PATH & "555555ffstriped.png"
m_arrBarColor (1,6) = IMAGES_PATH & "555555ffstriped.png"
m_arrBarColor (1,7) = IMAGES_PATH & "ffcc00ffstriped.png"

'// Accumulated Debt (Status: 0-Pending, 1-Waiting Approval from Supervisor, 2-Need Correction, etc)
m_arrBarColor (2,0) = IMAGES_PATH & "55aaffff.png"
m_arrBarColor (2,1) = IMAGES_PATH & "55aaffff.png"
m_arrBarColor (2,2) = IMAGES_PATH & "55aaffff.png"
m_arrBarColor (2,3) = IMAGES_PATH & "55aaffff.png"
m_arrBarColor (2,4) = IMAGES_PATH & "55aaffff.png"
m_arrBarColor (2,5) = IMAGES_PATH & "55aaffff.png"
m_arrBarColor (2,6) = IMAGES_PATH & "55aaffff.png"
m_arrBarColor (2,7) = IMAGES_PATH & "55aaffff.png"

TransparentPix = IMAGES_PATH & "00000000.png" 'transparent pixel

Dim rsPeriod, Period_, y, m, i, j
Dim iOfficial, iPersonal, iAccumulatedDebt
Dim iHeightOfficial, iHeightPersonal, iHeightAccumulatedDebt
Dim iStatus, iTotal


Dim rst
Dim strsql, arrResultSet, rs, rsempty, arrNumberList


BillingCon.Open
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

AlternateEmailFlag_ = trim(Request("AlternateEmailFlag"))

SortBy_ = Request.Form("SortList")
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


strsql = "Select Distinct MobilePhone From vwMonthlyBilling Where LoginID='" & user_ & "' order by MobilePhone"
set rs = server.createobject("adodb.recordset")
set rs = BillingCon.execute(strsql)

' MobilePhone
rsempty = false
If NOT rs.EOF Then
    arrNumberList = rs.GetRows()
	If MobilePhone_="" Then MobilePhone_ = arrNumberList(0,0)
Else
	rsempty = true
End If

EmpID_ = ""

strsql = "Exec spNavigator '" & EmpID_ & "','" & user_ & "','" & MobilePhone_ & "','" & sPeriod & "','" & ePeriod & "','" & GraphHeight & "'"
set rs = server.createobject("adodb.recordset")
set rs = BillingCon.execute(strsql)

' Official, Personal, HeightOfficial, HeightPersonal, MonthP, YearP, ProgressId, AccumulatedDebt, HeightAccumulatedDebt
If NOT rs.EOF Then
	arrResultSet = rs.GetRows()
End If




CellPhonePrsBillRp_ = 0
MaxAccumulatedDebt_ = 0
EmpName_ = ""
CellPhoneBillRp_ = 0
ProgressStatus_ = "Not assigned for this month"
AgencyFundingDesc_ = ""
EmailAddress_ = ""
SupervisorEmail_ = ""
Notes_ = ""
SpvRemark_ = ""
Office_ = ""
j = UBound(arrResultSet,2)
For i = 0 To j
	If (arrResultSet (6,i) = MonthP AND arrResultSet (7,i) = YearP) Then
		CellPhonePrsBillRp_ = arrResultSet (1,i)
		MaxAccumulatedDebt_ = arrResultSet (2,j)
		EmpName_ = arrResultSet (9,i)
		CellPhoneBillRp_ = arrResultSet (10,i)
		ProgressStatus_ = arrResultSet (11,i)
		AgencyFundingDesc_ = arrResultSet (12,j)
		EmailAddress_ = arrResultSet (13,i)
		SupervisorEmail_ = arrResultSet (14,i)
		Notes_ = arrResultSet (15,i)
		SpvRemark_ = arrResultSet (16,i)
		Office_ = arrResultSet (17,i)
		EmpID_ = arrResultSet (18,i)
		ProgressID_ = arrResultSet (19,i)
	End If
Next
Period_ = MonthP & " - " & YearP
If SupervisorEmail_ = "" Then
	SupervisorEmail_ = EmailAddress_
End If

%>

<div id="container">

	<div id="navigation">

						<form method="post" action="MonthlyBilling.asp?Func=3" name="frmMonthlyBilling"">
						<div class="selector_header">Name : <%=EmpName_%><br>Phone Number : <%=MobilePhone_ %>&nbsp;<br>Funded : <%=AgencyFundingDesc_%></div>
						<div class="selector_title">Billing Period</div>
						<div class="selector_info"><%if cdbl(CellPhoneBillRp_ ) > 0 Then %><%= MonthName(Cint(MonthP))%>&nbsp;<%= YearP%><%else%>- &nbsp;<%end if%></div>
						<div class="selector_title">Status</div>
						<div class="selector_info"><%if cdbl(CellPhoneBillRp_ ) => 0 Then %><%=ProgressStatus_%><%else%>- &nbsp;<%end if%></div>
						<div class="selector_title">Total Bill Amount</div>
						<div class="selector_info"><%if cdbl(CellPhoneBillRp_ ) > 0 Then %><%=formatnumber(CellPhoneBillRp_  ,-1) %>&nbsp;Kn<%else%>- &nbsp;<%end if%></div>
						<div class="selector_title">Personal Amount Due</div>
						<div class="selector_info"><%if cdbl(CellPhoneBillRp_ ) > 0 Then %><%=formatnumber(CellPhonePrsBillRp_  ,-1) %>&nbsp;Kn<%else%>- &nbsp;<%end if%></div>
						<div class="selector_title">Accumulated Debt for this number</div>
						<div class="selector_info"><%if cdbl(MaxAccumulatedDebt_ ) > 0 Then %><%=formatnumber(MaxAccumulatedDebt_  ,-1) %>&nbsp;Kn<%else%>- &nbsp;<%end if%></div>

						<%

							Response.Write "<table border=""0"" cellspacing=""0"" cellpadding=""0""  id=""chart3_table"">"
							Response.Write "<tr><td colspan=""" & (8)  & """ class=""selector_title"">Total Bill / Personal Amount Due</td><td colspan=""" & (4)  & """ class=""selector_graph_top""><img src=""" & IMAGES_PATH & "asc.gif" & """>" & eYearP & "<img src=""" & IMAGES_PATH & "desc.gif" & """></td></tr>"
							Response.Write "<tr>"

							j = 0
							For i = 0 To (NrOfMonths - 1)
								m = Month(DateAdd("m", i, CDate(sMonthP& "/01/" &sYearP)))
								y = Year(DateAdd("m", i, CDate(sMonthP& "/01/" &sYearP)))
								iMonth = MonthName(m ,True)
								iOfficial = ""
								iPersonal = ""
								iAccumulatedDebt = ""
								iHeightOfficial = 0
								iHeightPersonal = 0
								iHeightAccumulatedDebt = 0
								iStatus = 0
								iTotal = ""
								If (CInt(arrResultSet (6,j)) = m AND CInt(arrResultSet (7,j)) = y) Then
									iOfficial = CLng(arrResultSet (0,j))
									iPersonal = CLng(arrResultSet (1,j))
									iAccumulatedDebt = CLng(arrResultSet (2,j))
									iHeightOfficial = CLng(arrResultSet (3,j))
									iHeightPersonal = CLng(arrResultSet (4,j))
									iHeightAccumulatedDebt = CLng(arrResultSet (5,j))
									iStatus = arrResultSet (8,j) - 1
									iTotal = iOfficial + iPersonal
									j = j + 1
								End If
								If m < 10 Then m = "0" & CStr(m) Else m = CStr(m)

								Response.Write "<td valign=""top"" class=""barcell"">"
								If iTotal <> "" Then
									Response.Write "<a href=""MonthlyBilling.asp?CellPhone=" & MobilePhone_ & "&MonthP=" & m & "&YearP=" & y & """ style=""display:block; text-decoration: none;"">"
								Else
									Response.Write "<a href=""#"" style=""display:block; text-decoration: none;"">"
								End If
								Response.Write "<img src=""" & TransparentPix & """ width=""0"" height=""" & _
													GraphHeight - iHeightOfficial - iHeightPersonal & """ alt="""" title="""" />" & _
												"<br />" & _
												iTotal & "<br /><img src=""" & m_arrBarColor(0,iStatus) & """ width=""" & BarWidth & """ height=""" & _
													iHeightOfficial & """ alt="""" title=""" & iOfficial & """ />" & _
												"<br /><img src=""" & m_arrBarColor(1,iStatus) & """ width=""" & BarWidth & """ height=""" & _
													iHeightPersonal & """ alt="""" title=""" & iPersonal & """ />"
								If m = MonthP Then
									Response.Write "<div class=""chart3_labels_active"">" & iMonth & "</div>"
								Else
									Response.Write "<div class=""chart3_labels"">" & iMonth & "</div>"
								End If
								Response.Write "<div valign=""top"" class=""chart3_barcell_bottom""><img src=""" & m_arrBarColor(2,iStatus) & """ width=""" & BarWidth & """ height=""" & _
													iHeightAccumulatedDebt & """ alt="""" title=""Accumulated Debt"" />" & _
												"<br />" & iAccumulatedDebt & "<br><img src=""" & TransparentPix & """ width=""0"" height=""" & _
													GraphHeight - iHeightAccumulatedDebt & """ alt="""" title=""Accumulated Debt"" /></div></a></td>"
							Next

							Response.Write "</tr>"
							Response.Write "<tr><td colspan=""" & NrOfMonths & """ class=""selector_graph_bottom"" align=""right"">Accumulated Debt</td></tr>"
							Response.Write "</table>"
							%>

							<div class="selector_title">Supervisor</div>
							<select name="cmbSupervisor" <% If ((ProgressID_ <> 1) and (ProgressID_ <> 3)) then %>Disabled<%End If%> >
								<option value="">--Select--</option>
						<%
							strsql = "Select EmailAddress, EmpName, Office, WorkingTitle From vwDirectReport Where len(EmailAddress)>5 and Type = 'AMER' Order by EmpName"
							set rsSPV = server.createobject("adodb.recordset")
							set rsSPV = BillingCon.execute(strsql)
							do while not rsSPV.eof
								Ename_ = rsSPV("EmpName")
						%>
								<OPTION value=<%=rsSPV("EmailAddress")%> <%if trim(SupervisorEmail_) = trim(rsSPV("EmailAddress")) then %> Selected<%End If%>  >  <%= EName_  %>
						<%
								rsSPV.MoveNext
							Loop%>
							</select>
							<div class="selector_title">Note</div>
							<TextArea name="txtNotes" Rows="4" style="width:290px" Wrap <% if (ProgressID_  <> 1) and (ProgressID_ <> 3) then%>ReadOnly<%End If%> ><%=Notes_%></textarea>
							<div class="selector_title">Supervisor's Remarks / Corrections</div>
							<TextArea name="txtRemark" Rows="4" style="width:290px" Wrap <% if (ProgressID_  <> 1) or (ProgressID_ <> 3) then%>ReadOnly<%End If%>><%=SpvRemark_ %></textarea>

					<%		if (ProgressID_ = 1) or (ProgressID_ = 3) then%>
								<input type="submit" id="Accept" name="btnSubmit" Value="Submit to Supervisor" <%if isempty(request.querystring ("updated"))=true then response.write " disabled/><div><tr><td colspan=' & NrOfMonths & ' align='left'>Update change(s) to enable button!</td></tr></div><tr>" :end if%>
								<input type="hidden" name="txtMobilePhone" value='<%=MobilePhone_%>' />
								<input type="hidden" name="txtMonthP" value='<%=MonthP%>' />
								<input type="hidden" name="txtYearP" value='<%=YearP%>' />
								<input type="hidden" name="txtEmpID" value='<%=EmpID_ %>' />
								<input type="hidden" name="txtEmpName" value='<%=EmpName_%>' />
								<input type="hidden" name="txtFunded" value='<%=AgencyFundingDesc_%>' />
								<input type="hidden" name="txtPeriod" value='<%=Period_%>' />
								<input type="hidden" name="txtOffice" value='<%=Office_%>' />
								<input type="hidden" name="txtTotalCost" value='<%=CellPhoneBillRp_ %>' />
								<input type="hidden" name="txtTotalBillingPrsAmount" value='<%=CellPhonePrsBillRp_ %>' />
					<%		end if	%>
						</form>
	</div>

	<div id="wrapper">

		<div id="content">


		<%
if ProgressStatus_ <> "Not assigned for this month" Then
'if not rsempty or ProgressStatus_ <> "Not assigned for this month" Then
'if not rsCellPhone.eof Then

			strsql = "Select * From vwCellphoneHd Where PhoneNumber='" & MobilePhone_ & "' and MonthP='" & MonthP & "' and YearP='" & YearP & "'"
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
		%>
		<div class="details_header"><a href="#" onclick="show('Summary')">BILL SUMMARY</a></div>		<DIV ID="Summary" style="display:none">		<table class="details">
		<tr class="details_title">
			<td  colspan="3">Monthly Fees</strong> / <i>Mjesecne pretplate:</i></td>
		</tr>
		<tr>
			<td width="90%"><strong>Subscription Monthly Fee</strong> / <i>Mjesecna naknada za pretplatnicki broj<i/></td>
			<td width="3%">&nbsp;Kn.</td>
			<td width="7%" align="right"><%=formatnumber(SubscriptionFee_,-1) %>&nbsp;&nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;<strong>Data Monthly Fee</strong> / <i>Mjesecna naknada za mobilni prijenos podataka<i/></td>
			<td>&nbsp;Kn.</td>
			<td align="right"><%=formatnumber(FARIDA_,-1) %>&nbsp;&nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;<strong>Other Charges</strong> / <i>Ostale usluge<i/></td>
			<td>&nbsp;Kn.</td>
			<td align="right"><%=formatnumber(DetailedCallRecord_,-1) %>&nbsp;&nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td class="details_title" colspan="3">Usage Charges / <i>Pozivi i prijenos podataka</i></td>
		</tr>
		<tr>
			<td>&nbsp;<strong>VPN Network Calls</strong> / <i>Pozivi unutar VPN mreže<i/></td>
			<td>&nbsp;Kn.</td>
			<td align="right"><%=formatnumber(LocalCall_,-1) %>&nbsp;&nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;<strong>Calls to VIP Network</strong> / <i>Pozivi prema VIP mobilnoj mreži<i/></td>
			<td>&nbsp;Kn.</td>
			<td align="right"><%=formatnumber(BalanceDue_,-1) %>&nbsp;&nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;<strong>Calls to Landlines in Croatia</strong> / <i>Pozivi prema fiksnim mrežama u Hrvatskoj<i/></td>
			<td>&nbsp;Kn.</td>
			<td align="right"><%=formatnumber(Interlocal_,-1) %>&nbsp;&nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;<strong>Calls to Other Mobile Networks</strong> / <i>Pozivi prema ostalim mobilnim mrežama<i/></td>
			<td>&nbsp;Kn.</td>
			<td align="right"><%=formatnumber(IDD_,-1) %>&nbsp;&nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;<strong>SMS</strong> / <i>SMS poruke<i/></td>
			<td>&nbsp;Kn.</td>
			<td align="right"><%=formatnumber(SMS_,-1) %>&nbsp;&nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;<strong>MMS</strong> / <i>MMS Poruke<i/></td>
			<td>&nbsp;Kn.</td>
			<td align="right"><%=formatnumber(GPRS_,-1) %>&nbsp;&nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;<strong>International Calls from Croatia</strong> / <i>Medunarodni pozivi iz Hrvatske<i/></td>
			<td>&nbsp;Kn.</td>
			<td align="right"><%=formatnumber(IRL_,-1) %>&nbsp;&nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;<strong>Incoming Calls in Roaming</strong> / <i>Dolazni pozivi u roamingu<i/></td>
			<td>&nbsp;Kn.</td>
			<td align="right"><%=formatnumber(PreviousBalance_,-1) %>&nbsp;&nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;<strong>Outgoing Calls in Roaming</strong> / <i>Odlazni pozivi u roamingu<i/></td>
			<td>&nbsp;Kn.</td>
			<td align="right"><%=formatnumber(Adjustment_,-1) %>&nbsp;&nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;<strong>GPRS/EDGE/UMTS Data Transfer</strong> / <i>GPRS/EDGE/UMTS prijenos podataka<i/></td>
			<td>&nbsp;Kn.</td>
			<td align="right"><%=formatnumber(IPHONE_,-1) %>&nbsp;&nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td class="details_title">&nbsp;<strong>Neto Total</strong> / <i>Neto Total<i/></td>
			<td class="details_title">&nbsp;Kn.</td>
			<td class="details_title" align="right"><%=formatnumber(Payment_,-1) %>&nbsp;&nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;<strong>VAT</strong> / <i>PDV<i/></td>
			<td>&nbsp;Kn.</td>
			<td align="right"><%=formatnumber(PPN_,-1) %>&nbsp;&nbsp;&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;<strong>Services Exempted from VAT</strong> / <i>Usluge na koje se ne obracunava PDV<i/></td>
			<td>&nbsp;Kn.</td>
			<td align="right"><%=formatnumber(StampFee_,-1) %>&nbsp;&nbsp;&nbsp;</td>
		</tr>
		<tr class="details_title">
			<td>&nbsp;<strong>Grand Total</strong> / <i>Bruto Total<i/></td>
			<td>&nbsp;Kn.</td>
			<td align="right"><u><strong><%=formatnumber(CurrentBalance_,-1) %></strong></u>&nbsp;&nbsp;&nbsp;</td>
		</tr>
		</table>



</div>		<div class="details_header">USAGE DETAIL</div>
		<form method="post" action="MonthlyBilling.asp?Func=2" name="frmCellPhoneBilling">
		<table id="myTable" class="tablesorter">
		<thead>
		<tr>
		    <th>Dialed Date/time</th>
			<th>Dialed Number</th>
			<th>Call Type</th>
			<th>Call Duration</th>
			<th>Amount (Kn.)</th>
			<th>Check if<br>personal
<%			if (ProgressID_ = 1) or (ProgressID_ = 3) then %>
				<input type="checkbox" name="cbAll" value="true" onclick="checkall(this);EnableSubmit(this);" />
<%			end if %>
			</th>
		</tr>
		</thead>
		<tbody>
		<%
strsql = "Select DetailRecordAmount From PaymentDueDate"
'response.write strsql & "<br>"
set rsDetailRecord = server.createobject("adodb.recordset")
set rsDetailRecord = BillingCon.execute(strsql)
if not rsDetailRecord.eof then
	DetailRecordAmount_ = rsDetailRecord("DetailRecordAmount")
	'response.write "DetailRecordAmount :" & DetailRecordAmount_
end if


		strsql = "Select * from CellPhoneDt Where PhoneNumber='" & MobilePhone_ & "' and MonthP='" & MonthP & "' and YearP='" & YearP & "' Order by DialedDatetime Asc"
		set rsCellPhone = BillingCon.execute(strsql)

		no_ = 1
		do while not rsCellPhone.eof
   			'if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4"
'			if (ProgressID_ = 4) then
			if (cdbl(rsCellPhone("Cost")) <> cdbl(DetailRecordAmount_ )) then
		%>
			<tr>
			        <td>&nbsp;<%=rsCellPhone("DialedDatetime")%></td>
		        	<td>&nbsp;<%=rsCellPhone("DialedNumber")%></td>
		        	<td>&nbsp;<%=rsCellPhone("CallType")%></td>
		        	<td>&nbsp;<%=rsCellPhone("CallDuration")%></td>
			        <td align="right"><%=formatnumber(rsCellPhone("Cost"),-1)%></td>
<%'			if (ProgressID_ = 1) or (ProgressID_ = 3) then %>
<%'			if (cdbl(ProgressID_)<4 or (ProgressID_ = 4 and AlternateEmailFlag_="Y")) then %>

<%			if (((cdbl(ProgressID_) < 4 or cdbl(ProgressID_) = 8) and (InStr(1,rsCellPhone("CallType"),ExemptedIfOfficialCallType_,1) = 0 and InStr(1,rsCellPhone("CallType"),AlwaysExemptedCallType_,1) = 0))) then %>
		 	       <td align="center">
				<%if rsCellPhone("isPersonal") = "Y" then%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" onclick="EnableSubmit(this);" Value='<%=rsCellPhone("CallRecordID")%>' Checked>
				<%else%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" onclick="EnableSubmit(this);" Value='<%=rsCellPhone("CallRecordID")%>' >
				<%end if%>
				</td>
<%			else%>
		 	       <td align="center">
				<%if rsCellPhone("isPersonal") = "Y" then%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsCellPhone("CallRecordID")%>' Checked disabled>
				<%else%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsCellPhone("CallRecordID")%>'  disabled>
				<%end if%>
				</td>
<%			end if %>

			</tr>
		<%      end if
			rsCellPhone.movenext
			no_ = no_ + 1
		loop
		%>
		</tbody>
		<%
		if ((ProgressID_< 4 and no_ >1) or (ProgressID_ = 4 and AlternateEmailFlag_="Y")) then%>
				<input type="submit" name="btnSubmit" Value="Update Change(s)" />&nbsp;&nbsp;
				<input type="button" value="Cancel" onClick="javascript:location.href='MonthlyBilling.asp?CellPhone=<%=MobilePhone_%>&MonthP=<%=MonthP%>&YearP=<%=YearP%>'">

				<input type="hidden" name="txtMobilePhone" value='<%=MobilePhone_ %>' />
				<input type="hidden" name="txtMonthP" value='<%=MonthP%>' />
				<input type="hidden" name="txtYearP" value='<%=YearP%>' />
				<input type="hidden" name="txtEmpID" value='<%=EmpID_ %>' />
		<%end if%>
		</table>
		</form>
<%else%>
<table width="100%">
<tr>
	<td align="center">&nbsp;</td>
</tr>
<tr>
	<td align="center">No Data.  Please select a phone from the left menu.</td>
</tr>
</table>
<%end if

'Close the connection with the database and free all database resources
Set rs = Nothing
BillingCon.Close
Set BillingCon = Nothing

%>

		</div>

	</div>

<!--#include file="1NavigationHome.asp" -->


<%
Case 2
	MobilePhone_ = Request.Form("txtMobilePhone")
	EmpID_ = Request.Form("txtEmpID")
	MonthP_ = Request.Form("txtMonthP")
	YearP_ = Request.Form("txtYearP")

	'Save Detail
	strsql = "Update CellPhoneDt Set isPersonal='' Where PhoneNumber='" & MobilePhone_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ & "'"
	BillingCon.execute(strsql)

	For Each loopIndex in Request.Form("cbPersonal")
	'	response.write loopIndex

		strsql = "Update CellPhoneDt Set isPersonal='Y' Where PhoneNumber='" & MobilePhone_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ &"' And CallRecordID= " & loopIndex
		'response.write strsql & "<Br>"
		BillingCon.execute(strsql)
	next

	'Update table MonthlyBilling
	strsql = "spUpdateTotPersonalCall '3','" & EmpID_ & "','" & MobilePhone_ & "','" & MonthP_ & "','" & YearP_ & "'"
	'response.write strsql & "<Br>"
	BillingCon.execute(strsql)

'Close the connection with the database and free all database resources
BillingCon.Close
Set BillingCon = Nothing

Response.AddHeader "REFRESH","0;URL=MonthlyBilling.asp?CellPhone=" & MobilePhone_ & "&MonthP=" & MonthP_ & "&YearP=" & YearP_ & "&updated"


Case 3

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
	Funded_ = replace(Request.Form("txtFunded"),"'","''")
	Period_ = replace(Request.Form("txtPeriod"),"'","''")
	Office_ = replace(Request.Form("txtOffice"),"'","''")
	TotalCost_ = replace(Request.Form("txtTotalCost"),"'","''")
	TotalBillingPrsAmount_ = replace(Request.Form("txtTotalBillingPrsAmount"),"'","''")

	'Save Header
	strsql = "Update MonthlyBilling Set SupervisorEmail='" & SpvEmail_ & "', Notes='" & Notes_ & "' Where EmpID='" & EmpID_ & "' And PhoneNumber='" & MobilePhone_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ &"'"
	BillingCon.execute(strsql)

	strsql = "Update MonthlyBilling Set ProgressId=2, ProgressIdDate=GetDate() Where EmpID='" & EmpID_ & "' And PhoneNumber='" & MobilePhone_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ &"'"
	response.write strsql
	BillingCon.execute(strsql)

	Dim send_from, send_to, send_cc, send_bcc
	send_from = BillingDL
	send_to = SpvEmail_

	Dim ObjMail
	Set ObjMail = Server.CreateObject("CDO.Message")
	ObjMail.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	ObjMail.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer
	ObjMail.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	ObjMail.Configuration.Fields.Update
	objMail.From = send_from
	objMail.To = send_to
	'objMail.CC = send_cc

	objMail.Subject = "Action Required: eBilling System – Approval Request"

	objMail.HTMLBody = "<html><head>"
	ObjMail.HTMLBody = ObjMail.HTMLBody & " "_

& " <title>z-Billing Application</title> "_
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
		& "        <td colspan='6' align='Left' class='FontContent'>&nbsp;<u><strong>Personal Info:<strong></u></td></tr> "_
		& "    <tr> "_
		& "    <td colspan='6' align='Left'> "_
		& "    	<table cellspadding='1' border='2' bordercolor='black' cellspacing='3' width='100%' bgColor='#999999' border='0'>   "_
		& "    		<tr BGCOLOR='#999999'> "_
		& "    			<td colspan='3' style='border: none;' class='FontContent'><FONT color=#FFFFFF><strong>Employee Name : " & EmpName_ & "</strong></font></td> "_
		& "    			<td colspan='3' style='border: none;' align='right' class='FontContent'><FONT color=#FFFFFF><strong>Phone Number : " & MobilePhone_ & "&nbsp;</strong></font></td> "_
		& "    		</tr> "_
		& "    		<tr BGCOLOR='#999999'> "_
		& "    			<td colspan='6' style='border: none;' class='FontContent'><FONT color=#FFFFFF><strong>Office : " & Office_ & "</strong></font></td> "_
		& "    		</tr> "_
		& "    		<tr BGCOLOR='#999999'> "_
		& "    			<td colspan='6' style='border: none;' class='FontContent'><FONT color=#FFFFFF><strong>Funded : " & Funded_ & "</strong></font></td> "_
		& "    		</tr> "_
		& "    	</table></td></tr> " _
		& "    <tr> "_
		& "        <td align='Left' colspan='6' class='FontContent'>&nbsp;<u><strong>Billing Detail:<strong></u></td></tr> "_
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

		if cdbl(TotalCost_ ) > 0 Then

		ObjMail.HTMLBody = ObjMail.HTMLBody & " "_
		& "    	<tr height=26> "_
		& "    	<td class='FontContent'>&nbsp;<a href='" & WebSiteAddress & "/BillingApproval.asp?Cellphone=" & MobilePhone_ & "&LoginID=" & user_ & "&MonthP=" & MonthP_ & "&YearP=" & YearP_ & "&Nav=1' target='_blank'>Review and approve</a></td> "_
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
		& "        <td colspan='6' align='Left' class='FontContent'>&nbsp;<u><strong>Employee's Note:<strong></u></td></tr> "_
		& "    <tr> "_
		& "        <td colspan='6' align='Left' class='FontContent'>" & Notes_ & "</td></tr> "_
		& "        <td colspan='6'>&nbsp; </td></tr> "_
		& "        <td height=26 align='center' colspan='6' class='FontContent'>NOTE: This e-mail was automatically generated.</td> "_
		& "    </tr> "_
		& "        <td colspan='6'>&nbsp; </td></tr> "_
		& " </table></p>"_
		& "</body></html>"

		'response.write ObjMail.HTMLBody
	objMail.Send
	Set objMail = Nothing
	Set objConfig = Nothing

	'Close the connection with the database and free all database resources
	BillingCon.Close
	Set BillingCon = Nothing

'	Response.AddHeader "REFRESH","0;URL=MonthlyBilling.asp?CellPhone=" & MobilePhone_ & "&MonthP=" & MonthP_ & "&YearP=" & YearP_ & ""
'
End Select
%>
</BODY>
</html>
