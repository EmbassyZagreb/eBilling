<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="connect.inc" -->
<TITLE>U.S. Embassy Zagreb - zBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<style type="text/css">
<!--
.style1 {
	font-size: large;
	font-weight: bold;
	}

.FontText {
	font-size: small;
}

.FontContent {
	font-size: small;
        color: blue;
}

.FontJudul {
	font-size: 24px;
	font-weight: bold;
}

.FontComment {
	font-size: 18px;
	font-weight: bold;
}

-->
</style>
<link href="style.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function validate_form()
{
	valid = true;

	if (document.frmAgencyEdit.txtAgencyName.value == "" )
	{
		alert("Please fill the agency name !!!");
		valid = false;
	}

	if (document.frmAgencyEdit.txtAgencyStripe.value == "" )
	{
		alert("Please fill the agency stripe !!!");
		valid = false;
	}

	if (document.frmAgencyNew.txtAgencyStripe.value == "" )
	{
		alert("Please fill the Fiscal Strip VAT !!!");
		valid = false;
	}

	if (document.frmAgencyNew.txtAgencyStripeNonVAT.value == "" )
	{
		alert("Please fill the Fiscal Strip Non VAT !!!");
		valid = false;
	}

	return valid;
}
</script>
<%

Dim user_ , user1_, UserRole_

user_ = request.servervariables("remote_user")
user1_ = user_  'user1_ = right(user_,len(user_)-4)

ID_ = request.querystring("ID")
'response.write "ServiceRecordID : " & serviceRecordId & "<br>"

%>
</head>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="SubTitle">Agency Edit</TD>
  </TR>
  <tr>
	<td colspan="3" align="Left" width="20%"><A HREF="Default.asp">Home</A></td>
	<td align="Right" width="20%"><A HREF="AgencyList.asp">Back</A></td>
  </tr>
  <tr>
  	<td COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></td>
   </tr>
  </TABLE>
<form method="post" name="frmAgencyEdit" id="frmAgencyEdit" action="AgencySave.asp?Mode=U" onsubmit="return validate_form()">
<%
strsql = "select * from Users where loginId='" & user1_ & "'"
set RS_Query = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set RS_Query = BillingCon.execute(strsql)
'response.write RS_Query("RoleID") & "<br>"


if (trim(RS_Query("RoleID")) = "Admin") or (trim(RS_Query("RoleID")) = "Voucher") or (trim(RS_Query("RoleID")) = "FMC") then
%>
	<table align="center">
	<%
	   dim rsAgency
	   strsql = "Select * From AgencyFunding Where AgencyID =" & ID_
	   set rsAgency= server.createobject("adodb.recordset")
	'response.write strsql & "<br>"
	   set rsAgency= BillingCon.execute(strsql)
'	   Response.write rsAgency("Disabled")
	%>
	<tr>
		<td align="right">Agency Code:</td>
		<td align="left"><input name="txtAgencyCode" type="Input" size="8" value='<%=rsAgency("AgencyFundingCode")%>' ></td>
	</tr>
	<tr>
		<td align="right">Agency Name :</td>
		<td align="left"><input name="txtAgencyName" type="Input" size="60" value='<%=rsAgency("AgencyDesc")%>'></td>
	</tr>
	<tr>
		<td align="right">Fiscal Strip  VAT :</td>
		<td align="left"><input name="txtAgencyStripe" type="Input" size="100" value='<%=rsAgency("FiscalStripVAT")%>'></td>
	</tr>
	<tr>
		<td align="right">Fiscal Strip Non VAT :</td>
		<td align="left"><input name="txtAgencyStripeNonVAT" type="Input" size="100" value='<%=rsAgency("FiscalStripNonVAT")%>'></td>
	</tr>
	<tr>
		<td align="right">Disabled :</td>
		<td align="left">
		     <select name="txtAgencyType">
			<option value="">--Select--</option>
			<% if rtrim(rsAgency("Disabled"))="Y" then%>
				<option value="Y" selected>Yes</option>
			<%Else%>
				<option value="Y">Yes</option>
			<%End If%>
			<% if rtrim(rsAgency("Disabled"))="N" then%>
				<option value="N" Selected>No</option>
			<%Else%>
				<option value="N">No</option>
			<%End If%>
		     </select>
	</tr>
  	 <tr>
		<td colspan=2 align="center">
        		<input type="submit" value="Submit">
			&nbsp;<input type="button" value="Cancel" onClick="javascript:location.href='AgencyList.asp'">
			<INPUT TYPE="HIDDEN" NAME="txtID" value=<%=ID_%>>
    		</td>
  	</tr>
	<tr>
		<td colspan=2>&nbsp;</td></tr>
	</table>


<table border="0" bordercolor="#FFFFFF" cellpadding="2" cellspacing="0" width="80%"  class="FontText">
	<tr>
		<td><u><strong>Historical assignment of Funding Agency: <%=rsAgency("AgencyDesc")%><strong></u></td>
	</tr>
	<tr>
		<td class="Hint" align="left">*To alter historical data 'Generate Monthly Billing' procedure must be executed. Procedure sets bill to 'Pending' status.</td>
	</tr>
</table>









<%

strsql = "Select Distinct YearP+MonthP, YearP, MonthP, AgencyFundingCode, AgencyFundingDesc, FiscalStripVAT, FiscalStripNonVAT From vwMonthlyBilling Where AgencyID = '" & ID_ & "' Order by (YearP+MonthP) Desc"
set DataRS = server.createobject("adodb.recordset")
DataRS.CursorLocation = 3
DataRS.Open strsql,BillingCon
'set DataRS=BillingCon.execute(strsql)

dim intPageSize,PageIndex,TotalPages
dim RecordCount,RecordNumber,Count
intpageSize=10
PageIndex=request("PageIndex")

if PageIndex ="" then PageIndex=1

if not DataRS.eof then
	RecordCount = DataRS.RecordCount
	'response.write RecordCount & "<br>"
	RecordNumber=(intPageSize * PageIndex) - intPageSize
	'response.write RecordNumber
	DataRS.PageSize =intPageSize
	DataRS.AbsolutePage = PageIndex
	TotalPages=DataRS.PageCount
	'response.write TotalPages & "<br>"
End If
'response.write strsql

dim intPrev,intNext
intPrev=PageIndex - 1
intNext=PageIndex +1


if not DataRS.eof Then

%>
<!-- <div align="right"><input type="submit" name="btnApproval" value="Approve" /></div> -->
<!-- <div align="right"><input type="button" name="btnExport" value="Export to Excel" onClick="javascript:document.location.href('ARBillingReportAllPrint.asp?sMonthP=<%=sMonthP%>&sYearP=<%=sYearP%>&eMonthP=<%=eMonthP%>&eYearP=<%=eYearP%>&Agency=<%=Agency_%>&Section=<%=Section_%>&EmpID=<%=EmpID_%>&Status=<%=Status%>');"/></div> -->
<table border="1" bordercolor="#EEEEEE" cellpadding="2" cellspacing="0" width="80%"  class="FontText">
    <TR BGCOLOR="#330099" align="center">
         <TD width="8%"><strong><label STYLE=color:#FFFFFF>Billing<br>Period</label></strong></TD>
         <TD width="8%"><strong><label STYLE=color:#FFFFFF>Agency<br>Code</label></strong></TD>
         <TD width="20%"><strong><label STYLE=color:#FFFFFF>Agency Name</label></strong></TD>
         <TD width="32%"><strong><label STYLE=color:#FFFFFF>Fiscal Strip VAT</label></strong></TD>
         <TD width="32%"><strong><label STYLE=color:#FFFFFF>Fiscal Strip Non VAT</label></strong></TD>


    </TR>
<%
   dim no_
   no_ = 1 + ((PageIndex*intPageSize)-intPageSize)
   Count=1
   do while not DataRS.eof   and Count<=intPageSize
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4"
%>

	   <TR bgcolor="<%=bg%>">
<!--	    <TD align="right">&nbsp;<%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font>&nbsp;</TD> -->
		<TD align="right">&nbsp;<a href="AgencyMembers.asp?AgencyID=<%=ID_%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>&AgencyFundingCode=<%=DataRS("AgencyFundingCode")%>&AgencyFundingDesc=<%= DataRS("AgencyFundingDesc")%>&FiscalStripVAT=<%= DataRS("FiscalStripVAT")%>&FiscalStripNonVAT=<%= DataRS("FiscalStripNonVAT")%>" target="_blank"><%= DataRS("MonthP")%>-<%= DataRS("YearP")%></a></font>&nbsp;</TD>
	    <TD>&nbsp;<%=DataRS("AgencyFundingCode") %></TD>
	    <TD>&nbsp;<%=DataRS("AgencyFundingDesc") %> </font></TD>
		<TD>&nbsp;<%= DataRS("FiscalStripVAT") %></font></TD>
		<TD>&nbsp;<%= DataRS("FiscalStripNonVAT") %></font></TD>
	   </TR>

<%
		Count=Count +1
	   DataRS.movenext
	   no_ = no_ + 1
   loop
	PageNo=1
%>
</table>
<table width="80%">
	<tr>
		<td align="right">
<%
		Do while PageNo<=TotalPages
			if trim(pageNo) = trim(PageIndex) Then
%>
				<label class="ActivePage"><%=PageNo%></label>&nbsp;
			<%Else%>
				<a href="AgencyEdit.asp?PageIndex=<%=PageNo%>&ID=<%=ID_%>&State=E"><%=PageNo%></a>&nbsp;
<%
			End If
			PageNo=PageNo+1
		Loop
%>
		</td>
	</tr>
</table>
<%
else
%>
	<table cellspadding="1" cellspacing="0" width="100%">
	<tr>
        	<td><br></TD>
	</tr>
	<tr>
		<td align="center">not data.</td>
	</tr>
	<tr>
        	<td><br></TD>
	</tr>
	<tr>
		<td align="center"><a href="Default.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
	</tr>
	</table>
<% end if %>










<%Else %>
	<table>
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
		</tr>
	</table>

<% end if %>
</form>
</body>
</html>
