<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>
  <%
Response.ContentType ="application/vnd.ms-excel" 
Response.Buffer  =  True 
Response.Clear() 
%> 
 
<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!--#include file="connect.inc" -->


<%
Dim user_ , user1_

user_ = request.servervariables("remote_user")
user1_ = user_  'user1_ = right(user_,len(user_)-4)
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

sMonthP = request("sMonthP")
'response.write sMonthP

sYearP = request("sYearP")
eMonthP = request("eMonthP")

eYearP = request("eYearP")

Agency_ = request("Agency")

Section_ = request("Section")
%>
<TITLE>U.S. Embassy Zagreb - zBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
<%

dim rs 
dim strsql
If (UserRole_ = "Admin") or (UserRole_ = "Voucher") or (UserRole_ = "FMC") Then
%>
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

if not DataRS.eof Then

%>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="90%"  class="FontText">
    <TR align="center" BGCOLOR="#330099">
         <TD class="style5" width="3%" align="right"><strong>No.</strong></TD>
         <TD class="style5"><strong>Funding Agency</strong></TD>
         <TD class="style5"><strong>Phone Number</strong></TD>
         <TD class="style5"><strong>Employee Name</strong></TD>
         <TD class="style5" width="10%"><strong>Billing Period</strong></TD>
         <TD class="style5"><strong>Section</strong></TD>
         <TD class="style5"><strong>VAT</strong></TD>
         <TD class="style5"><strong>Non-VAT</strong></TD>
         <TD class="style5"><strong>Total</strong></TD>
	 <TD class="style5" BGCOLOR="#333333"><strong>Accounting</strong></TD>
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
	    if PrevAgency_ <> DataRS("AgencyFundingCode") Then
		SubTotalVal_ = 0
		SubTotalNonVal_ = 0
		SubTotal_ = 0
   		SubTotalVip_ = 0
		no_ = 1
%>
		<tr>
			<td colspan="10" BGCOLOR="#999999" class="style5"><b><%=DataRS("AgencyFunding") %></b></td>
		</tr>		
<%
	    end if
%>
      
	   <TR bgcolor="<%=bg%>">
	        <TD align="right"><%=Count%>~<%=no_ %>&nbsp;</font></TD>
	        <TD>&nbsp;<%=DataRS("AgencyFunding") %></TD>
	        <TD>&nbsp;<%=DataRS("MobilePhone") %></TD>
	        <TD>&nbsp;<%=DataRS("EmpName") %></TD>
	        <TD align="right">&nbsp;<%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font>&nbsp;</TD>
	        <TD>&nbsp;<%=DataRS("Office") %> </font></TD>
	        <TD align="right"><%= formatnumber(DataRS("VAT"),-1) %></font></TD>
		<TD align="right"><%= formatnumber(DataRS("NonVAT"),-1) %></font></TD>
		<TD align="right"><%= formatnumber(DataRS("Total"),-1) %></font></TD>
		<TD align="right" BGCOLOR="#999999" class="style5">
<%		If CDbl(DataRS("TotalVip")) > 0 Then %>
			<%= formatnumber(DataRS("TotalVip"),-1) %>
<%		Else %>
			-
<%		End If %>
		</font></TD>
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

		PrevAgency_ = DataRS("AgencyFundingCode")
		PrevAgencyName_ = DataRS("AgencyFunding")
		FiscalStripVAT_ = DataRS("FiscalStripVAT")
		FiscalStripNonVAT_ = DataRS("FiscalStripNonVAT")
		Count=Count +1
	   DataRS.movenext
'		if (PrevAgency_ <> DataRS("AgencyFundingCode")) or DataRS.eof Then 
		if DataRS.eof Then 
%>
		<tr>
			<td colspan="6"><b><%=PrevAgencyName_ %> SubTotal</b></td>			
			<td align="right"><b><%=formatnumber(cdbl(SubTotalVal_),-1) %></font></b></td>			
			<td align="right"><b><%=formatnumber(cdbl(SubTotalNonVal_),-1) %></font></b></td>			
			<td align="right"><b><%=formatnumber(cdbl(SubTotal_),-1) %></font></b></td>
			<td align="right" BGCOLOR="#666666" class="style5"><b><%=formatnumber(cdbl(SubTotalVip_),-1) %></b></td>
		</tr>
		<tr><td colspan="9"><br></td></tr>
<%
			if FiscalStripVAT_ <> FiscalStripNonVAT_ then
%>
				<tr>
					<td colspan="8"><%=FiscalStripVAT_ %></td>
					<td align="right"><%=formatnumber(cdbl(SubTotalVal_),-1) %></font></td>
					<td>&nbsp;</td>
				</tr>	
				<tr>
					<td colspan="8"><%=FiscalStripNonVAT_ %></td>
					<td align="right"><%=formatnumber(cdbl(SubTotalNonVal_),-1) %></font></td>
					<td>&nbsp;</td>
				</tr>	
<%
			else
%>
				<tr>
					<td colspan="8">&nbsp;<%=FiscalStripVAT_ %></td>
					<td align="right"><%=formatnumber(cdbl(SubTotal_),-1) %></font></td>
					<td>&nbsp;</td>
				</tr>	
<%
			end if
		elseif (PrevAgency_ <> DataRS("AgencyFundingCode"))Then 
%>
		<tr>
			<td colspan="6"><b><%=PrevAgencyName_ %> SubTotal</b></td>			
			<td align="right"><b><%=formatnumber(cdbl(SubTotalVal_),-1) %></font></b></td>			
			<td align="right"><b><%=formatnumber(cdbl(SubTotalNonVal_),-1) %></font></b></td>			
			<td align="right"><b><%=formatnumber(cdbl(SubTotal_),-1) %></font></b></td>
			<td align="right" BGCOLOR="#666666" class="style5"><b><%=formatnumber(cdbl(SubTotalVip_),-1) %></b></td>
		</tr>
		<tr><td colspan="9"><br></td></tr>
<%
			if FiscalStripVAT_ <> FiscalStripNonVAT_ then
%>
				<tr>
					<td colspan="8">&nbsp;<%=FiscalStripVAT_ %></td>
					<td align="right"><b><%=formatnumber(cdbl(SubTotalVal_),-1) %></font></b></td>
					<td>&nbsp;</td>
				</tr>	
				<tr>
					<td colspan="8">&nbsp;<%=FiscalStripNonVAT_ %></td>
					<td align="right"><b><%=formatnumber(cdbl(SubTotalNonVal_),-1) %></font></b></td>
					<td>&nbsp;</td>
				</tr>	
<%
			else
%>
				<tr>
					<td colspan="8">&nbsp;<%=FiscalStripVAT_ %></td>
					<td align="right"><b><%=formatnumber(cdbl(SubTotal_),-1) %></font></b></td>
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
			<td align="right"><FONT color=#FFFFFF><b><%=formatnumber(cdbl(TotalVal_),-1) %></font></b></td>			
			<td align="right"><FONT color=#FFFFFF><b><%=formatnumber(cdbl(TotalNonVal_),-1) %></font></b></td>
			<td align="right"><FONT color=#FFFFFF><b><%=formatnumber(cdbl(Total_),-1) %></font></b></td>
			<td align="right" BGCOLOR="#333333"><FONT color=#FFFFFF><b><%=formatnumber(cdbl(TotalVip_),-1) %></font></b></td>			
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


