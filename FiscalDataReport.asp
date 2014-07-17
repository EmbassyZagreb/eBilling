<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>
 
 
<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!--#include file="connect.inc" -->


<script type="text/javascript">

function ClearFilter()
{
	document.forms['frmSearch'].elements['cmbAgency'].value="X";
	document.forms['frmSearch'].elements['cmbSection'].value="X";
	document.forms['frmSearch'].elements['cmbEmp'].value="X";
}

</script>
<%
Dim user_ , user1_

user_ = request.servervariables("remote_user")
user1_ = right(user_,len(user_)-4)
'response.write user1_ & "<br>"

strsql = "select RoleID from Users where loginId ='" & user1_ & "'"
set UserRS = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set UserRS = BillingCon.execute(strsql)
if not UserRS.eof then
	UserRole_ = UserRS("RoleID")
Else
	UserRole_ = ""
end if

strsql = "Select Max(YearP+MonthP) As Period From vwMonthlyBilling"
'response.write strsql & "<br>"
set rsPeriod = server.createobject("adodb.recordset") 
set rsPeriod = BillingCon.execute(strsql)
if not rsPeriod.eof Then
	Period_ = rsPeriod("Period")
end if
'response.write Period_ & "<br>"

If Period_ <> "" Then
	curMonth_ = Right(Period_, 2)
	curYear_ = Left(Period_, 4)
Else
	curMonth_ = month(date())
	curYear_ = year(date())
End If

'curMonth_ = month(date())
'curYear_ = year(date())
if len(curMonth_)= 1 then
	curMonth_ = "0" & curMonth_
end if

sMonthP = request("sMonthP")

if sMonthP = "" Then sMonthP = Request.Form("sMonthList")
if sMonthP = "" then
	sMonthP = curMonth_ 
end if
'response.write sMonthP

sYearP = request("sYearP")
if sYearP ="" Then sYearP = Request.Form("sYearList")
if sYearP ="" then
	sYearP = curYear_ 
end if

eMonthP = request("eMonthP")
if eMonthP = "" Then eMonthP = Request.Form("eMonthList")
if eMonthP = "" then
	eMonthP = curMonth_ 
end if

eYearP = request("eYearP")
if eYearP = "" Then eYearP = Request.Form("eYearList")
if eYearP = "" then
	eYearP = curYear_ 
end if

Agency_ = request("Agency")
if Agency_ = "" Then Agency_ = Request.Form("cmbAgency")
if Agency_ = "" then
	Agency_ = "X"
end if

Section_ = request("Section")
if Section_ = "" Then Section_ = Request.Form("cmbSection")
if Section_ = "" then
	Section_ = "X"
end if
%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">FISCAL DATA REPORT</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>

<%

dim rs 
dim strsql
If (UserRole_ = "Admin") or (UserRole_ = "Voucher") or (UserRole_ = "FMC") Then
%>
<form method="post" name="frmSearch" Action="FiscalDataReport.asp">
<table cellspadding="1" cellspacing="0" width="70%" border="1" align="center">
<tr align="Center">
	<td colspan="2" align="center">
		<table  width="100%">
		<tr bgcolor="#000099">
			<td height="25" colspan=6"><strong>&nbsp;<span class="style5">Criteria(s): </span></strong></td>
		</tr>
		<tr>
			<td width="15%" align="right">&nbsp;Period&nbsp;</td>				
			<td>:</td>
			<td colspan="4">
				<Select name="sMonthList">
					<Option value="01" <%if sMonthP ="01" then %>Selected<%End If%>>January</Option>
					<Option value="02" <%if sMonthP ="02" then %>Selected<%End If%>>February</Option>
					<Option value="03" <%if sMonthP ="03" then %>Selected<%End If%>>March</Option>
					<Option value="04" <%if sMonthP ="04" then %>Selected<%End If%>>April</Option>
					<Option value="05" <%if sMonthP ="05" then %>Selected<%End If%>>May</Option>
					<Option value="06" <%if sMonthP ="06" then %>Selected<%End If%>>June</Option>
					<Option value="07" <%if sMonthP ="07" then %>Selected<%End If%>>July</Option>
					<Option value="08" <%if sMonthP ="08" then %>Selected<%End If%>>August</Option>
					<Option value="09" <%if sMonthP ="09" then %>Selected<%End If%>>September</Option>
					<Option value="10" <%if sMonthP ="10" then %>Selected<%End If%>>October</Option>
					<Option value="11" <%if sMonthP ="11" then %>Selected<%End If%>>November</Option>
					<Option value="12" <%if sMonthP ="12" then %>Selected<%End If%>>December</Option>
				</Select>&nbsp;
<%
				Year_ = Year(Date()) - 1
%>

				<Select name="sYearList">
<% 				Do While Year_ <= Year(Date()) %>
				<Option value='<%=Year_%>' <%if trim(Year_) = trim(sYearP) then %>Selected<%End If%> ><%=Year_%></Option>
<% 
			Year_ = Year_ + 1
			Loop %>	
				</Select>&nbsp;to&nbsp;
				<Select name="eMonthList">
					<Option value="01" <%if eMonthP ="01" then %>Selected<%End If%>>January</Option>
					<Option value="02" <%if eMonthP ="02" then %>Selected<%End If%>>February</Option>
					<Option value="03" <%if eMonthP ="03" then %>Selected<%End If%>>March</Option>
					<Option value="04" <%if eMonthP ="04" then %>Selected<%End If%>>April</Option>
					<Option value="05" <%if eMonthP ="05" then %>Selected<%End If%>>May</Option>
					<Option value="06" <%if eMonthP ="06" then %>Selected<%End If%>>June</Option>
					<Option value="07" <%if eMonthP ="07" then %>Selected<%End If%>>July</Option>
					<Option value="08" <%if eMonthP ="08" then %>Selected<%End If%>>August</Option>
					<Option value="09" <%if eMonthP ="09" then %>Selected<%End If%>>September</Option>
					<Option value="10" <%if eMonthP ="10" then %>Selected<%End If%>>October</Option>
					<Option value="11" <%if eMonthP ="11" then %>Selected<%End If%>>November</Option>
					<Option value="12" <%if eMonthP ="12" then %>Selected<%End If%>>December</Option>
				</Select>&nbsp;
<%
				Year_ = Year(Date()) - 1
%>

				<Select name="eYearList">
<% 				Do While Year_ <= Year(Date()) %>
				<Option value='<%=Year_%>' <%if trim(Year_) = trim(eYearP) then %>Selected<%End If%> ><%=Year_%></Option>
<% 
			Year_ = Year_ + 1
			Loop %>	
				</Select>
			</td>
		</tr>
		<tr>
			<td align="right">&nbsp;Funding Agency&nbsp;</td>
			<td>:</td>
			<td>
<%
 				strsql ="select AgencyID, AgencyFundingCode, AgencyDesc from AgencyFunding order by AgencyDesc"
				set AgencyRS = server.createobject("adodb.recordset")
				set AgencyRS = BillingCon.execute(strsql)
'				response.write strStr 
%>	
				<Select name="cmbAgency">
					<Option value='X'>--All--</Option>
<%				Do While not AgencyRS.eof %>
					<Option value='<%=AgencyRS("AgencyID")%>' <%if trim(Agency_) = trim(AgencyRS("AgencyID")) then %>Selected<%End If%> ><%=AgencyRS("AgencyDesc")%></Option>
					
<%					AgencyRS.MoveNext
				Loop%>
				</select>

			</td>	
			<td align="right">Section&nbsp;</td>
			<td>:</td>
			<td>
<%
 				strsql ="select distinct Office from vwPhoneCustomerList Where Office<>'' order by Office"
				set SectionRS = server.createobject("adodb.recordset")
				set SectionRS = BillingCon.execute(strsql)
'				response.write strStr 
%>	
				<Select name="cmbSection">
					<Option value='X'>--All--</Option>
<%				Do While not SectionRS.eof %>
					<Option value='<%=SectionRS("Office")%>' <%if trim(Section_) = trim(SectionRS("Office")) then %>Selected<%End If%> ><%=SectionRS("Office")%></Option>
					
<%					SectionRS.MoveNext
				Loop%>
				</select>

			</td>	
		</tr>
		<tr>
			<td colspan="2">&nbsp;</td>
			<td align="left">
				<input type="Button" name="btnReset" value="Reset" onClick="Javascript:ClearFilter();">	
			</td>
			<td align="Left" colspan="3">
				<input type="submit" name="Submit" value="Search">
			</td>
		</tr>
		</table>
	</td>
</tr>	
</table>
</form>
<%
sPeriod = sYearP&sMonthP
ePeriod = eYearP&eMonthP
'response.write sPeriod & ePeriod 
strsql = "Select * From vwFundingDataReport Where YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "'"
strFilter=""
If Agency_ <> "X" then
	strFilter=strFilter & " and AgencyID='" & Agency_ & "'"
End If

If Section_ <> "X" then
	strFilter=strFilter & " and Office='" & Section_ & "'"
End If

strsql = strsql  & strFilter & " order by AgencyFunding, EmpName, MobilePhone, YearP+MonthP"
'response.write strsql & "<br>"

set DataRS = server.createobject("adodb.recordset")
DataRS.CursorLocation = 3
DataRS.Open strsql,BillingCon
'set DataRS=BillingCon.execute(strsql)

'response.write strsql

dim intPrev,intNext 	
intPrev=PageIndex - 1 
intNext=PageIndex +1 


if not DataRS.eof Then

%>
<table width="90%">
<tr>
	<td align="right"><input type="button" name="btnExport" value="Export to Excel" onClick="javascript:document.location.href('FiscalDataReportPrint.asp?sMonthP=<%=sMonthP%>&sYearP=<%=sYearP%>&eMonthP=<%=eMonthP%>&eYearP=<%=eYearP%>&Agency=<%=Agency_%>&Section=<%=Section_%>');"/></td>
</tr>
</table>
<table border="1" bordercolor="#EEEEEE" cellpadding="2" cellspacing="0" width="90%"  class="FontText">
    <TR BGCOLOR="#330099" align="center">
         <TD width="3%" align="right"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Funding Agency</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Phone Number</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
         <TD width="10%"><strong><label STYLE=color:#FFFFFF>Billing Period</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Section</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>VAT</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Non-VAT</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Total</label></strong></TD>
         <TD BGCOLOR="#333333"><strong><label STYLE=color:#FFFFFF>Accounting</label></strong></TD>
    </TR>

<% 
   dim no_  
   no_ = 1
   Count=1
   PrevAgency_ =""
   SubTotalVal_ = 0
   SubTotalNonVal_ = 0
   SubTotal_ = 0
   SubTotalVip_ = 0
   TotalVal_ = 0
   TotalNonVal_ = 0
   Total_ = 0
   TotalVip_ = 0
   do while not DataRS.eof
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
	    if PrevAgency_ <> DataRS("AgencyFunding") Then
		SubTotalVal_ = 0
		SubTotalNonVal_ = 0
		SubTotal_ = 0
   		SubTotalVip_ = 0
		no_ = 1
%>
		<tr BGCOLOR="#999999">
			<td colspan="10"><FONT color=#FFFFFF><b><%=DataRS("AgencyFunding") %></b></font></td>
		</tr>		
<%
	    end if
%>
      
	   <TR bgcolor="<%=bg%>">
	        <TD align="right"><%=Count%>~<%=no_ %>&nbsp;</font></TD>
	        <TD>&nbsp;<%=DataRS("AgencyFunding") %></TD>
	        <TD>&nbsp;<%=DataRS("MobilePhone") %></TD>
	        <TD>&nbsp;<a href="CellPhoneDetail.asp?CellPhone=<%=DataRS("MobilePhone") %>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%=DataRS("EmpName") %></a></TD>
	        <TD align="right">&nbsp;<%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font>&nbsp;</TD>
	        <TD>&nbsp;<%=DataRS("Office") %> </font></TD>
	        <TD align="right"><%= formatnumber(DataRS("VAT"),-1) %>&nbsp;</font></TD>
		<TD align="right">&nbsp;<%= formatnumber(DataRS("NonVAT"),-1) %></font></TD>
		<TD align="right">&nbsp;<%= formatnumber(DataRS("Total"),-1) %></font></TD>
		<TD align="right" BGCOLOR="#999999">&nbsp;<label STYLE=color:#FFFFFF>
<%		If CDbl(DataRS("TotalVip")) > 0 Then %>
			<%= formatnumber(DataRS("TotalVip"),-1) %>
<%		Else %>
			-
<%		End If %>
		</label></font></TD>
	   </TR>

<%   
		SubTotalVal_ = cdbl(SubTotalVal_) + cdbl(DataRS("VAT"))
		SubTotalNonVal_ = cdbl(SubTotalNonVal_) + cdbl(DataRS("NonVAT"))
		SubTotal_ = cdbl(SubTotal_) + cdbl(DataRS("Total"))
		SubTotalVip_ = cdbl(SubTotalVip_) + cdbl(DataRS("TotalVip"))

		TotalVal_ = cdbl(TotalVal_) + cdbl(DataRS("VAT"))
		TotalNonVal_ = cdbl(TotalNonVal_) + cdbl(DataRS("NonVAT"))
		Total_ = cdbl(Total_) + cdbl(DataRS("Total"))
		TotalVip_ = cdbl(TotalVip_) + cdbl(DataRS("TotalVip"))

		PrevAgency_ = DataRS("AgencyFunding")
		PrevAgencyName_ = DataRS("AgencyFunding")
		FiscalStripVAT_ = DataRS("FiscalStripVAT")
		FiscalStripNonVAT_ = DataRS("FiscalStripNonVAT")
		Count=Count +1
	   DataRS.movenext
'		if (PrevAgency_ <> DataRS("AgencyFunding")) or DataRS.eof Then 
		if DataRS.eof Then 
%>
		<tr>
			<td colspan="6"><b><%=PrevAgencyName_ %> SubTotal</b></td>			
			<td align="right"><b>&nbsp;<%=formatnumber(cdbl(SubTotalVal_),-1) %></font></b></td>			
			<td align="right"><b>&nbsp;<%=formatnumber(cdbl(SubTotalNonVal_),-1) %></font></b></td>			
			<td align="right"><b>&nbsp;<%=formatnumber(cdbl(SubTotal_),-1) %></font></b></td>
			<td align="right" BGCOLOR="#666666"><b>&nbsp;<label STYLE=color:#FFFFFF><%=formatnumber(cdbl(SubTotalVip_),-1) %></label></b></td>
		</tr>
		<tr><td colspan="10"><br></td></tr>
<%
			if FiscalStripVAT_ <> FiscalStripNonVAT_ then
%>
				<tr>
					<td colspan="8">&nbsp;<%=FiscalStripVAT_ %></td>
					<td align="right">&nbsp;<%=formatnumber(cdbl(SubTotalVal_),-1) %></font></td>
					<td>&nbsp;</td>
				</tr>	
				<tr>
					<td colspan="8">&nbsp;<%=FiscalStripNonVAT_ %></b></td>
					<td align="right">&nbsp;<%=formatnumber(cdbl(SubTotalNonVal_),-1) %></font></td>
					<td>&nbsp;</td>
				</tr>	
<%
			else
%>
				<tr>
					<td colspan="8">&nbsp;<%=FiscalStripVAT_ %></td>
					<td align="right">&nbsp;<%=formatnumber(cdbl(SubTotal_),-1) %></font></td>
					<td>&nbsp;</td>
				</tr>	
<%
			end if
		elseif (PrevAgency_ <> DataRS("AgencyFunding"))Then 
%>
		<tr>
			<td colspan="6"><b><%=PrevAgencyName_ %> SubTotal</b></td>			
			<td align="right"><b>&nbsp;<%=formatnumber(cdbl(SubTotalVal_),-1) %></font></b></td>			
			<td align="right"><b>&nbsp;<%=formatnumber(cdbl(SubTotalNonVal_),-1) %></font></b></td>			
			<td align="right"><b>&nbsp;<%=formatnumber(cdbl(SubTotal_),-1) %></font></b></td>
			<td align="right" BGCOLOR="#666666"><b>&nbsp;<label STYLE=color:#FFFFFF><%=formatnumber(cdbl(SubTotalVip_),-1) %></label></b></td>
		</tr>
		<tr><td colspan="9"><br></td></tr>
<%
			if FiscalStripVAT_ <> FiscalStripNonVAT_ then
%>
				<tr>
					<td colspan="8">&nbsp;<%=FiscalStripVAT_ %></td>
					<td align="right"><b>&nbsp;<%=formatnumber(cdbl(SubTotalVal_),-1) %></font></b></td>
					<td>&nbsp;</td>
				</tr>	
				<tr>
					<td colspan="8">&nbsp;<%=FiscalStripNonVAT_ %></td>
					<td align="right"><b>&nbsp;<%=formatnumber(cdbl(SubTotalNonVal_),-1) %></font></b></td>
					<td>&nbsp;</td>
				</tr>	
<%
			else
%>
				<tr>
					<td colspan="8">&nbsp;<%=FiscalStripVAT_ %></td>
					<td align="right"><b>&nbsp;<%=formatnumber(cdbl(SubTotal_),-1) %></font></b></td>
					<td>&nbsp;</td>
				</tr>	
<%
			end if
		end if
	   no_ = no_ + 1 
   loop 
	PageNo=1
%>

		<tr  BGCOLOR="#999999" height=26>
			<td colspan="6"><FONT color=#FFFFFF><b>&nbsp;Total</b></font></td>			
			<td align="right"><FONT color=#FFFFFF><b>&nbsp;<%=formatnumber(cdbl(TotalVal_),-1) %>&nbsp;</font></b></td>			
			<td align="right"><FONT color=#FFFFFF><b>&nbsp;<%=formatnumber(cdbl(TotalNonVal_),-1) %>&nbsp;</font></b></td>
			<td align="right"><FONT color=#FFFFFF><b>&nbsp;<%=formatnumber(cdbl(Total_),-1) %>&nbsp;</font></b></td>
			<td align="right" BGCOLOR="#333333"><FONT color=#FFFFFF><b>&nbsp;<%=formatnumber(cdbl(TotalVip_),-1) %>&nbsp;</font></b></td>			
		</tr>

</table>
<table width="90%">
	<tr>
		<td align="right">
<%
		Do while PageNo<=TotalPages 
			if trim(pageNo) = trim(PageIndex) Then
%>		
				<label class="ActivePage"><%=PageNo%></label>&nbsp;
			<%Else%>
				<a href="FiscalDataReport.asp?PageIndex=<%=PageNo%>&sMonthP=<%=sMonthP%>&sYearP=<%=sYearP%>&eMonthP=<%=eMonthP%>&eYearP=<%=eYearP%>&Agency=<%=Agency_%>&Section=<%=Section_%>"><%=PageNo%></a>&nbsp;
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
		<td align="center">There is no data.</td>
	</tr>
	<tr>
        	<td><br></TD>
	</tr>
	<tr>
		<td align="center"><a href="Default.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
	</tr>	
	</table>
<% end if %>
<%Else%>
	<table>
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
		</tr>
	</table>
<% end if %>
</body> 

</html>


