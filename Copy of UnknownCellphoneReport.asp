<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>
 
 
<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!--#include file="connect.inc" -->

<style type="text/css">
<!--

.FontText {
	font-size: small;
}


.Hint{
	color: gray;
	font-size: x-small;
}

-->
</style>
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

curMonth_ = month(date())
curYear_ = year(date())
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

%>
<TITLE>U.S. Mission Jakarta e-Billing</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">

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
<!--#include file="Header.inc" -->
 <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">UNKNOWN CELLPHONE BILL REPORT</TD>
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
If (UserRole_ = "Admin") or (UserRole_ = "Voucher") or (UserRole_ = "FMC") or (UserRole_ = "Cashier")  Then
%>
<form method="post" name="frmSearch" Action="UnknownCellphoneReport.asp">
<table cellspadding="1" cellspacing="0" width="40%" border="1" align="center">
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
			<td colspan="2">&nbsp;</td>
			<td align="left">
				<input type="Button" name="btnReset" value="Reset" onClick="Javascript:ClearFilter();">	
			</td>
			<td align="Left" colspan="2">
				<input type="submit" name="Submit" value="Search">
			</td>
			<td align="Right">
				<input type="button" name="btnExport" value="Export to Excel" onClick="javascript:document.location.href('UnknownCellphoneReportPrint.asp?sMonthP=<%=sMonthP%>&sYearP=<%=sYearP%>&eMonthP=<%=eMonthP%>&eYearP=<%=eYearP%>');"/>
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
strsql = "Select * From vwUnknownCellphoneBill Where YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "'"

strsql = strsql  & strFilter & " order by EmpName Desc"
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
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="60%"  class="FontText">
    <TR BGCOLOR="#330099" align="center">
         <TD width="3%" align="right"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
         <TD width="100px"><strong><label STYLE=color:#FFFFFF>Billing Period</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Phone Number</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Office</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Current Balance</label></strong></TD>
    </TR>
<% 
   dim no_  
   no_ = 1
   Count=1
   do while not DataRS.eof
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>    
	   <TR bgcolor="<%=bg%>">
	        <TD align="right"><FONT color=#330099 size=2><%=no_ %>&nbsp;</font></TD>
	        <TD align="right"><FONT color=#330099 size=2>&nbsp;<%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font>&nbsp;</TD>
	        <TD><FONT color=#330099 size=2>&nbsp;<%=DataRS("PhoneNumber") %> </font></TD>
	        <TD><FONT color=#330099 size=2>&nbsp;<%=DataRS("EmpName") %> </font></TD>
	        <TD><FONT color=#330099 size=2>&nbsp;<%=DataRS("Office") %> </font></TD>
	        <TD align="right"><FONT color=#330099 size=2><%= formatnumber(DataRS("CurrentBalance"),0) %>&nbsp;</font></TD>
	   </TR>

<%   
		Count=Count +1
	   DataRS.movenext
	   no_ = no_ + 1 
   loop 
%>
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
			<td>Please <a href="http://jakartaws01.eap.state.sbu/CSC">Submit Request </a> or contact Jakarta CSC Helpdesk at ext.9111.</td>
		</tr>
	</table>
<% end if %>
</body> 

</html>

