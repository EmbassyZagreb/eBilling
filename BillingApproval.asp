<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>

<script type="text/javascript">
function validate_form()
{
	valid = true;
	msg = ""

	if (document.frmBillingApproval.SupervisorSign.value == "" )
	{
		msg = msg + "Please take a decision for this billing !!! "
		valid = false;
	}


	if (valid == false)
	{
		alert(msg)
	}
	return valid;
}
</script>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Billing Approval</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<% 
EmpID = request("EmpID") 
MonthP = request("Month")
YearP = request("Year")

%>  
<form method="post" name="frmBillingApproval" action="BillingApprovalSave.asp" onsubmit="return validate_form();">	
<table cellspadding="1" cellspacing="0" width="65%" bgColor="white">  
<%
'strsql = "Exec spGetMonthlyBill '" & user1_ & "','" & MonthP & "','" & YearP & "'"
strsql = "Select * From vwMonthlyBilling Where EmpID='" & EmpID & "' and MonthP='" & MonthP & "' and YearP='" & YearP & "'"
'response.write strsql & "<br>"
set rsData = server.createobject("adodb.recordset") 
set rsData = BillingCon.execute(strsql) 
Period_ = MonthP & " - " & YearP
'response.write Period_  & "<br>"
if not rsData.eof then
	EmpName_ = rsData("EmpName")
	Office_ = rsData("Agency") & " - " & rsData("Office")
	Position_ = rsData("WorkingTitle")
	OfficePhone_ = rsData("WorkPhone")
	HomePhone_ = rsData("HomePhone")
	MobilePhone_ = rsData("MobilePhone")
	EmpEmail_ = rsData("EmailAddress")	
	ExchangeRate_ = rsData("ExchangeRate")
	Note_ =  rsData("Notes")
	SupervisorRemark_ = rsData("SupervisorRemark")
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
	TotalBillingAmount_ = rsData("TotalBillingAmountPrsRp")
	TotalBillingAmountPrsDlr_ = rsData("TotalBillingAmountPrsDlr")
'response.write TotalBillingRp_ & "<br>"
'response.write TotalBillingAmount_ & "<br>"
%>
<!-- <tr>
	<td colspan="6" align="center"><u>Billing Period (Month - Year) : <a class="FontContent"><%=Period_%></a></u></td>
</tr> -->
<tr>
          <td  colspan="6" align="Left"><u><b>Personal Info<b></u></TD>
</tr>




<tr>
	<td colspan="6" align="Left">
	<table cellspadding="1" border="2" bordercolor="black" cellspacing="3" width="100%" bgColor="#999999" border="0">  

		<tr BGCOLOR="#999999">
			<td colspan="3" style="border: none;"><FONT color=#FFFFFF><b>Employee Name : <%=EmpName_%></b></font></td>
			<td colspan="3" style="border: none;" align="right"><FONT color=#FFFFFF><b>Phone Number : <%=MobilePhone_ %>&nbsp;</b></font></td>
		</tr>
		<tr BGCOLOR="#999999">
			<td colspan="6" style="border: none;"><FONT color=#FFFFFF><b>Position : <%=Position_%></b></font></td>
		</tr>
		<tr BGCOLOR="#999999">
			<td colspan="6" style="border: none;"><FONT color=#FFFFFF><b>Agency / Office : <%=Office_%></b></font></td>
		</tr>
	</table>
	</td>
</tr>




  
<!-- <tr>
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
	<td>Mobile Phone</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=MobilePhone_ %></td>
</tr>
<tr>
	<td>Exchange Rate</td>
	<td width="1%">:</td>
	<td class="FontContent">Kn. <%= FormatNumber(ExchangeRate_,-1) %> / Dollar</td>
	<td>Payment Status</td>
	<td width="1%">:</td>
	<td class="FontContent">
	<%'If (ProgressID_ = 1) or (ProgressID_ = 3) then %>
		<a href="OfficePhoneDetail.asp?Username=<%=user1_ %>&MonthP=<%=MonthP%>&YearP=<%=YearP%>" target="_blank"><%=ProgressStatus_%></a>
	<%'Else%>
		<%=ProgressStatus_%>
	<%'End if%>
	</td>
</tr>
<tr>
	<td colspan="6"><hr></td>
</tr> -->
<tr>
	<td align="Left" colspan="6"><u><b>Approval<b></u></TD>
</tr>
<tr>
	<td colspan="6">
	<table class="FontComment" width="100%">
	<tr>
		<td>Your decision</td>
		<td width="1%">:</td>
		<td width="20%" class="FontContent">
	<%if ProgressID_ = 2 then%>
			<input type="radio" name="SupervisorSign" value="A" checked>Approve</input>
			<input type="radio" name="SupervisorSign" value="C" >Need Correction</input>&nbsp;&nbsp;&nbsp;&nbsp;		

	        	<input type="submit" value="Submit">
        		&nbsp;<input type="button" value="Cancel" onClick="javascript:location.href='BillingApprovalList.asp'">
			<input type="hidden" name="txtExtension" value='<%=OfficePhone_ %>' />
			<input type="hidden" name="txtMonthP" value='<%=MonthP%>' />
			<input type="hidden" name="txtYearP" value='<%=YearP%>' />
			<input type="hidden" name="txtEmpEmail" value='<%=EmpEmail_ %>' />

			<input type="hidden" name="txtEmpID" value='<%=EmpID %>' />
			<input type="hidden" name="txtEmpName" value='<%=EmpName_%>' />
			<input type="hidden" name="txtPeriod" value='<%=Period_%>' />
			<input type="hidden" name="txtOffice" value='<%=Office_%>' />
			<input type="hidden" name="txtTotalCost" value='<%=TotalBillingRp_ %>' />
			<input type="hidden" name="txtTotalBillingAmount" value='<%=TotalBillingAmount_ %>' />
			<input type="hidden" name="txtNote" value='<%=Note_%>' />
	<%Else%>
			<input type="radio" name="SupervisorSign" value="A" <%if cdbl(ProgressID_) >= 4 Then%>checked<%end if%> disabled>Approve</input>
			<input type="radio" name="SupervisorSign" value="C" <%if ProgressID_ = 3 Then%>checked<%end if%> disabled>Need Correction</input>&nbsp;&nbsp;&nbsp;&nbsp;		
	<%End If%>
		
		</td>
	</tr>


	<tr>
		<td width="12%" valign="top">Remark/Correction(s)</td>
		<td width="1%" valign="top">:</td>
		<td><TextArea name="txtRemark" Rows="3" Cols="65" Wrap maxlength="500"><%=SupervisorRemark_ %></textarea></td>
  	</tr>
	</table>	
	</td>
</tr>









<tr>
	<td align="Left" colspan="6"><u><b>Billing detail :<b></u></TD>
</tr>
<tr>
	<td colspan="6">*Click on the bill for more detail</td>
</tr>
<tr>
	<td align="Left" colspan="6">
	<table cellspadding="1" border="1" bordercolor="black" cellspacing="0" width="100%" bgColor="white" border="0">  
	<tr align="center" height=26>
		<td width=20%><b>Action</b></td>
		<td width=20%><b>Billing Period</b></td>
		<td width=20%><b>Status</b></td>
		<td width=20%><b>Billing (Kn.)</b></td>
		<td width=20%><b>Should be paid (Kn.)</b></td>
	</tr>
<!-- <%if cdbl(OfficePhoneBillRp_) > 0 Then %>
	<tr>
		<td><a href="OfficePhoneDetail.asp?Extension=<%=OfficePhone_ %>&MonthP=<%=MonthP%>&YearP=<%=YearP%>" target="_blank">Office Phone</a></td>
		<td width="20%" class="FontContent" align="right"><%=formatnumber(OfficePhoneBillRp_,-1) %>&nbsp;</td>
		<td width="20%" class="FontContent" align="right"><%=formatnumber(OfficePhonePrsBillRp_ ,-1) %>&nbsp;</td>
		<td width="20%" class="FontContent" align="right"><%=formatnumber(OfficePhonePrsBillDlr_,-1) %>&nbsp;</td>		
	</tr>
<%else%>
	<tr>
		<td>Office Phone</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
	</tr>
<%end if%> -->
<!-- <%if cdbl(HomePhoneBillRp_) > 0 Then %>
	<tr>
		<td><a href="HomePhoneDetail.asp?HomePhone=<%=HomePhone_%>&MonthP=<%=MonthP%>&YearP=<%=YearP%>" target="_blank">Home Phone</a></td>
		<td class="FontContent" align="right"><%=formatnumber(HomePhoneBillRp_ ,-1) %>&nbsp;</td>
		<td class="FontContent" align="right"><%=formatnumber(HomePhonePrsBillRp_ ,-1) %>&nbsp;</td>
		<td class="FontContent" align="right"><%=formatnumber(HomePhonePrsBillDlr_ ,-1) %>&nbsp;</td>
	</tr>
<%else%>
	<tr>
		<td>Home Phone</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
	</tr>
<%end if%> -->
<%if cdbl(CellPhoneBillRp_ ) > 0 Then %>
	<tr height=26>
		<td>&nbsp;<a href="CellPhoneDetail.asp?CellPhone=<%=MobilePhone_ %>&MonthP=<%=MonthP%>&YearP=<%=YearP%>" target="_blank">View Submitted Bill</a></td>
	        <TD align="right">&nbsp;<%= MonthP%>-<%= YearP%></font>&nbsp;</TD>
	        <TD align="right"><%=ProgressStatus_%>&nbsp;</font></TD>
		<td align="right"><%=formatnumber(CellPhoneBillRp_  ,-1) %>&nbsp;</td>
		<td align="right"><%=formatnumber(CellPhonePrsBillRp_ ,-1) %>&nbsp;</td>



	</tr>
<%else%>
	<tr>
		<td>Mobile Phone</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
	</tr>
<%end if%>
<!-- <%if cdbl(TotalShuttleBillRp_) > 0 Then %>
	<tr>
		<td><a href="ShuttleBusBillDetail.asp?Username=<%=user1_ %>&MonthP=<%=MonthP%>&YearP=<%=YearP%>" target="_blank">Shuttle Bus</a></td>
		<td class="FontContent" align="right"><%=formatnumber(TotalShuttleBillRp_ ,-1) %>&nbsp;</td>
		<td class="FontContent" align="right"><%=formatnumber(TotalShuttleBillRp_ ,-1) %>&nbsp;</td>
		<td class="FontContent" align="right"><%=formatnumber(TotalShuttleBillDlr_,-1) %>&nbsp;</td>
	</tr>
<%else%>
	<tr>
		<td>Shuttle Bus</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
	</tr>
<%end if%> -->
	</table>
	</TD>
</tr>
<tr>
	<td width="32%" valign="top">Employee's Note</td>
	<td width="1%" valign="top">:</td>
	<td>
		<TextArea name="txtNotes" Rows="3" Cols="65" Wrap readonly><%=Note_%></textarea>
	</td>
</tr>

<%Else%>
<tr>
	<td colspan="6" align="center">there is no data.</td>	
</tr>
<% end if %>
</table>
</form>
</BODY>
</html>