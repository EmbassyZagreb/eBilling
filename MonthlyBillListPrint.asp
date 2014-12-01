<HTML>
<HEAD>
<!--#include file="connect.inc" -->
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 
<TITLE>U.S. Embassy Zagreb - zBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
nRec = 0
'response.write "Test : " & Request("cbPrint")
   If Request("cbPrint") <> "" then
	For Each loopIndex in Request("cbPrint")
		nRec = nRec + 1
		'response.write loopIndex & "<br>"
		X = len(loopIndex)
		'response.write X & "<br>"
		EmpID_ = Left(loopIndex, X-6)
		'response.write EmpID_ & "<br>"
		Period = right(loopIndex,6)
		MonthP = left(Period,2)
		'response.write MonthP_ & "<br>"
		YearP = Right(Period,4)
		'response.write YearP_ & "<br>"
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
		
		strsql = "Select * from vwMonthlyBilling Where EmpID='" & EmpID_ & "' And MonthP='" & MonthP & "' And YearP='" & YearP & "'"
		'response.write strsql & "<br>"
		set rsData = server.createobject("adodb.recordset") 
		set rsData = BillingCon.execute(strsql) 
		Period_ = MonthP & " - " & YearP
		'response.write Period_  & "<br>"
		if not rsData.eof then
			EmpID_ = rsData("EmpID")
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
			ProgressID_ = rsData("ProgressID")
			ProgressStatus_ = rsData("ProgressDesc")
			SupervisorEmail_ = rsData("SupervisorEmail")
			If SupervisorEmail_ = "" Then
				SupervisorEmail_ = rsData("EmailAddress")
			End If
			Notes_ = rsData("Notes")
			SpvRemark_ = rsData("SupervisorRemark")
		'response.write Period_  & "<br>"
		statusPageBreak_ = nRec mod 2
'		response.write "mod:"  & statusPageBreak_ 
		end If
		If statusPageBreak_ = 1 and nRec >1 then
		%>		
			<P Class="PageBreak">&nbsp;</P>
		<%End If%>
		<table cellspadding="1" cellspacing="0" width="60%" bgColor="white" align="center">  
		<tr>
			<td  colspan="6"><br><br></td>
		</tr>
		<tr>
			<td colspan="6" align="center"><h3>Monthly Bill</h3></td>
		</tr>
		<tr>
			<td colspan="6" align="center"><u>Billing Period (Month - Year) : <a class="FontContent"><%=Period_%></a></u></td>
		</tr>
		<tr>
		          <td align="Left"><u><b>Personal Info<b></u></TD>
		</tr>  
		<tr>
			<td width="20%">Employee Name</td>
			<td width="1%">:</td>
			<td class="FontContent"><%=EmpName_%></td>
			<td>Agency / Office</td>
			<td width="1%">:</td>
			<td class="FontContent"><%=Office_%></td>
		</tr>
		<tr>
			<td>Position</td>
			<td width="1%">:</td>
			<td class="FontContent"><%=Position_ %></td>
			<td>Office Phone/Ext.</td>
			<td width="1%">:</td>
			<td class="FontContent"><%=OfficePhone_ %></td>
		</tr>
		<tr>
			<td>Homephone</td>
			<td width="1%">:</td>
			<td class="FontContent"><%=HomePhone_ %></td>
		</tr>
		<tr>
			<td>Mobile Phone</td>
			<td width="1%">:</td>
			<td class="FontContent"><%=MobilePhone_ %></td>
			<td>Exchange Rate</td>
			<td width="1%">:</td>
			<td class="FontContent">Kn. <%= FormatNumber(ExchangeRate_,-1 %> / Dollar</td>
		
		</tr>
		<tr>
			<td>Payment Status</td>
			<td width="1%">:</td>
			<td class="FontContent" colspan="4"><%=ProgressStatus_%></td>
		</tr>
		<tr>
			<td colspan="6"><hr></td>
		</tr>
		
		<tr>
			<td align="Left" colspan="5"><u><b>Billing detail :<b></u></TD>
		</tr>
		<tr>
			<td colspan="6">*Click on each billing type for more detail</td>
		</tr>
		<tr>
			<td align="Left" colspan="6">
			<table cellspadding="1" border="1" bordercolor="black" cellspacing="0" width="100%" bgColor="white" border="0">  
			<tr align="center">
				<td rowspan="2"><b>Type</b></td>
				<td rowspan="2"><b>Billing (Kn.)</b></td>
				<td colspan="2"><b>Should be paid</b></td>
			</tr>
			<tr>
				<td align="center"><b>In Kuna (Kn.)</b></td>
				<td align="center"><b>In US Dollar ($)</b></td>
			</tr>
		<%if cdbl(OfficePhoneBillRp_) > 0 Then %>
			<tr>
				<td>Office Phone</td>
				<td width="20%" class="FontContent" align="right"><%=formatnumber(OfficePhoneBillRp_,-1 %>&nbsp;</td>
				<td width="20%" class="FontContent" align="right"><%=formatnumber(OfficePhonePrsBillRp_ ,-1 %>&nbsp;</td>
				<td width="20%" class="FontContent" align="right"><%=formatnumber(OfficePhonePrsBillDlr_,-1 %>&nbsp;</td>		
			</tr>
		<%else%>
			<tr>
				<td>Office Phone</td>
				<td class="FontContent" align="right">- &nbsp;</td>
				<td class="FontContent" align="right">- &nbsp;</td>
				<td class="FontContent" align="right">- &nbsp;</td>
			</tr>
		<%end if%>
		<%if cdbl(HomePhoneBillRp_) > 0 Then %>
			<tr>
				<td>Home Phone</td>
				<td class="FontContent" align="right"><%=formatnumber(HomePhoneBillRp_ ,-1 %>&nbsp;</td>
				<td class="FontContent" align="right"><%=formatnumber(HomePhonePrsBillRp_ ,-1 %>&nbsp;</td>
				<td class="FontContent" align="right"><%=formatnumber(HomePhonePrsBillDlr_ ,-1 %>&nbsp;</td>
			</tr>
		<%else%>
			<tr>
				<td>Home Phone</td>
				<td class="FontContent" align="right">- &nbsp;</td>
				<td class="FontContent" align="right">- &nbsp;</td>
				<td class="FontContent" align="right">- &nbsp;</td>
			</tr>
		<%end if%>
		<%if cdbl(CellPhoneBillRp_ ) > 0 Then %>
			<tr>
				<td>Mobile Phone</td>
				<td class="FontContent" align="right"><%=formatnumber(CellPhoneBillRp_  ,-1 %>&nbsp;</td>
				<td class="FontContent" align="right"><%=formatnumber(CellPhonePrsBillRp_ ,-1 %>&nbsp;</td>
				<td class="FontContent" align="right"><%=formatnumber(CellPhonePrsBillDlr_ ,-1 %>&nbsp;</td>
			</tr>
		<%else%>
			<tr>
				<td>Mobile Phone</td>
				<td class="FontContent" align="right">- &nbsp;</td>
				<td class="FontContent" align="right">- &nbsp;</td>
				<td class="FontContent" align="right">- &nbsp;</td>
			</tr>
		<%end if%>
		<%if cdbl(TotalShuttleBillRp_) > 0 Then %>
			<tr>
				<td>Shuttle Bus</td>
				<td class="FontContent" align="right"><%=formatnumber(TotalShuttleBillRp_ ,-1 %>&nbsp;</td>
				<td class="FontContent" align="right"><%=formatnumber(TotalShuttleBillRp_ ,-1 %>&nbsp;</td>
				<td class="FontContent" align="right"><%=formatnumber(TotalShuttleBillDlr_,-1 %>&nbsp;</td>
			</tr>
		<%else%>
			<tr>
				<td>Shuttle Bus</td>
				<td class="FontContent" align="right">- &nbsp;</td>
				<td class="FontContent" align="right">- &nbsp;</td>
				<td class="FontContent" align="right">- &nbsp;</td>
			</tr>
		<%end if%>
			</table>
			</TD>
		</tr>
		<tr>
			<td colspan="6">
			<table cellspadding="1" cellspacing="0" width="100%" bgColor="white" border="0">
			<tr>
				<td align="center"><b>Total</b></td>
				<td width="20%" class="FontContent" align="right">&nbsp;</td>
				<td width="20%" class="FontContent" align="right"><b><u><%=formatnumber(TotalBillingRp_ ,-1) %></u></b>&nbsp;</td>
				<td width="20%" class="FontContent" align="right"><b><u><%=formatnumber(TotalBillingDlr_ ,-1 %></u></b>&nbsp;</td>
			</tr>
			</table>	
			</td>
		</tr>
		</table>
<%
	next
  End If

'  response.redirect("BillingApprovalList.asp")
%>
			<script language="JavaScript">
					print();
				</script>
</body> 
</html>