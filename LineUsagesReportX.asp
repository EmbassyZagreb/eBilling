<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>

<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!--#include file="connect.inc" -->

<script language="JavaScript" src="calendar.js"></script>
<script language="JavaScript">
function ClearFilter()
{
	document.forms['frmSearch'].elements['PostList'].value ='';
	document.forms['frmSearch'].elements['SortList'].value ='TotalCost';
}
</script>


<%
if (session("Month") = "") or (session("Year") = "") then
	strsql = "Select MonthP, YearP From Period"
	'response.write strsql & "<br>"
	set rsData = server.createobject("adodb.recordset") 
	set rsData = BillingCon.execute(strsql)
	if not rsData.eof then
		session("Month") = rsData("MonthP")
		session("Year") = rsData("YearP")
	end if
end if

Post_ = Request.Form("PostList")
if Post_ ="" then
	Post_ = Request("Post")	
end if

PhoneNumber_ = Request.Form("PhoneNumberList")
if PhoneNumber_ ="" then
	PhoneNumber_ = Request("PhoneNumber")	
end if

StartDate_ = Request.Form("txtStartDate")
'response.write "StartDate :" & StartDate_
if StartDate_ ="" then
	if Request("StartDate")<>"" then
		StartDate_ = Request("StartDate")
	else
		StartDate_ = Date()-30
	end if
end if

EndDate_ = Request.Form("txtEndDate")
'response.write "EndDate :" & EndDate_ 
if EndDate_ ="" then
	if Request("EndDate")<>"" then
		EndDate_ = Request("EndDate")
	else
		EndDate_ = Date()
	end if
end if

SortBy_ = Request.Form("SortList")
'response.write "SortBy" & SortBy_
if (SortBy_ ="") then
	if Request("SortBy")<>"" then
		SortBy_ = Request("SortBy")	
	Else
		SortBy_ = "TotalCost"
	end if
end if

Order_ = Request.Form("OrderList")
if (Order_ ="") then
	if Request("Order")<>"" then
		Order_ = Request("OrderList")	
	Else
		Order_ = "Asc"
	end if
end if
%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
  <Center><FONT COLOR=#009900><B>SENSITIVE BUT UNCLASSIFIED</Center></FONT></B>
  <BR>
<CENTER>
  <IMG SRC="images/embassytitle2.jpeg" WIDTH="661" HEIGHT="80" BORDER="0"> 
  <TABLE WIDTH="65%" BORDER="0" CELLPADDING="0" CELLSPACING="0">
  <CAPTION><H3 STYLE="font-size:17px;color:#000040">Mission Zagreb - Billing Application</H3></CAPTION>
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">LINE USAGES REPORT</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<table align="center" cellspadding="1" cellspacing="0" width="100%">  
<tr>
	<td colspan="2" align="center">Billing Period : <Label style="color:blue"><%= StartDate_ %> - <%= EndDate_ %></lable></td>
</tr>
</table>
<%

dim rs 
dim strsql
dim tombol
dim hlm
%>

<%
Dim user_ , user1_

dim intPageSize,PageIndex,TotalPages 
dim RecordCount,RecordNumber,Count 
intpageSize=20 
PageIndex=request("PageIndex")

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

If (UserRole_ <> "") Then
	strsql = "spLineUsagesReport '" & Post_ & "','" & StartDate_ & "','" & EndDate_ & "','" & SortBy_ & "','" & Order_ & "'"
	set DataRS = server.createobject("adodb.recordset")
	'response.write strsql & "<br>"
	DataRS.CursorLocation = 3
	DataRS.open strsql,BillingCon

	if ((PageIndex ="") or (request.form("btnSearch")="Search")) then PageIndex=1 
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

%>
	<table cellspadding="1" cellspacing="0" width="60%" border="1" align="center">
	<tr align="Center">
		<td colspan="2" align="center">
			<form method="post" name="frmSearch" Action="LineUsagesReport.asp">
			<table  width="100%">
			<tr bgcolor="#000099">
				<td height="25" colspan="7"><strong>&nbsp;<span class="style5">Search &amp; Sort By </span></strong></td>
			</tr>
			<tr>
				<td width="15%">&nbsp;Post&nbsp;</td>
				<td>:</td>
		        	<td>
				<Select name="PostList">
					<Option value="">-- All --</Option>
					<Option value="ZAGREB" <%if Post_ ="ZAGREB" then %>Selected<%End If%> >ZAGREB</Option>
					<Option value="PODGORICA" <%if Post_ ="PODGORICA" then %>Selected<%End If%> >PODGORICA</Option>
				</Select>
				</td>  					
				<td>&nbsp;Sort By&nbsp;</td>
				<td>:</td>
				<td>
					<Select name="SortList">
						<Option value="B.Extension" <%if SortBy_ ="B.Extension" then %>Selected<%End If%> >Extension</Option>
						<Option value="EmpName" <%if SortBy_ ="EmpName" then %>Selected<%End If%> >Employee Name</Option>
						<Option value="TotalCall" <%if SortBy_ ="TotalCall" then %>Selected<%End If%> >Total Call</Option>
						<Option value="TotalDuration" <%if SortBy_ ="TotalDuration" then %>Selected<%End If%> >Total Duration</Option>
						<Option value="TotalCost" <%if SortBy_ ="TotalCost" then %>Selected<%End If%> >Total Cost</Option>
					</Select>&nbsp;
					<Select name="OrderList">
						<Option value="Asc" <%if Order_ ="Asc" then %>Selected<%End If%> >Asc</Option>
						<Option value="Desc" <%if Order_ ="Desc" then %>Selected<%End If%> >Desc</Option>
					</Select>
				</td
			</tr>
			<tr>
				<td>&nbsp;Phone Ext.&nbsp;</td>
				<td>:</td>
		        	<td>
				<%
					strsql = "Select PhoneNumber From MsOfficePhoneNumber order by phonenumber"
					set PhoneRS = server.createobject("adodb.recordset")
					set PhoneRS = BillingCon.execute(strsql)
				%>
					<select name="PhoneNumberList"> 
						<Option value="A">-- All --</option>
				<%	do while not PhoneRS.eof 
					if trim(PhoneNumber_) = trim(PhoneRS("PhoneNumber")) then%>
						<OPTION value=<%=PhoneRS("PhoneNumber")%> Selected /><%=PhoneRS("PhoneNumber")%>
				<%	else	%>
						<OPTION value=<%=PhoneRS("PhoneNumber")%> /><%=PhoneRS("PhoneNumber")%>			
				<% end if%>
				<%		PhoneRS.movenext	
					loop%>
					</Select>
				</td>	
			</tr>
			<tr>
				<td>&nbsp;Date Report&nbsp;</td>				
				<td>:</td>
				<td colspan="4">
					<input name="txtStartDate" type="Input" size="10" value='<%=StartDate_ %>' maxlength="10">
					<a href="javascript:cal0.popup();"><img src="images/calendar.gif" width="34" height="18" border="0" alt="Calendar"></a>
					&nbsp;To&nbsp;
					<input name="txtEndDate" type="Input" size="10" value='<%=EndDate_ %>' maxlength="10">
					<a href="javascript:cal1.popup();"><img src="images/calendar.gif" width="34" height="18" border="0" alt="Calendar"></a>											

				</td>
			</tr>			
               		<tr>
			       <td>&nbsp;&nbsp;<a href="javascript:ClearFilter();">Clear Filter</a></td>
		               <td height="30" colspan="6" align="center">
<!--					<input type="Button" name="btnBack" value="Back" onClick="Javascript:document.location.href('BillingReportList.asp');"> -->
					<input type="Button" name="btnHome" value="Home" onClick="javascript:document.location.href('Default.asp');"/>
					<input type="submit" name="Submit" value="Search  / Show All">
				</td>
        		</tr>
			</table>
			</form>
		</td>
	</tr>	
	</table>
	<form method="post" name="frmLineUsagesList" action="LineUsagesPrint.asp?Post=<%=Post_%>&StartDate=<%=StartDate_ %>&EndDate=<%=EndDate_%>&SortBy=<%=SortBy_%>&Order=<%=Order_%>">
	<table cellpadding="1" cellspacing="0" width="100%">
	<tr>
		<td width="50%">&nbsp;<input type="submit" value="Export to Excel" /></td>
	</tr>
	</table>
	<table align="center" cellpadding="1" cellspacing="0" width="100%" border="1" bordercolor="black"> 
	<TR BGCOLOR="#000099" align="center" cellpadding="0" cellspacing="0" >
		<TD width="4%"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
	       	<TD width="14%"><strong><label STYLE=color:#FFFFFF>Extension</label></strong></TD>
		<TD><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
	       	<TD width="14%"><strong><label STYLE=color:#FFFFFF>Total Call</label></strong></TD>
		<TD width="15%"><strong><label STYLE=color:#FFFFFF>Total Duration (Second)</label></strong></TD>
		<TD width="10%"><strong><label STYLE=color:#FFFFFF>Total Cost ( in Rupiah )</label></strong></TD>
	</TR>    
	<% 
		dim no_  
	'	no_ = 1 
		no_ = 1 + ((PageIndex*intPageSize)-intPageSize)
		do while not DataRS.eof and Count<intPageSize
	   	if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
	%>
	   	<TR bgcolor="<%=bg%>">
			<td align="right"><%=No_%>&nbsp;</td>
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("Extension")%></font></td> 
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("EmpName")%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2>&nbsp;<%=DataRS("TotalCall")%></font>&nbsp;</td> 
	        	<td align="right"><FONT color=#330099 size=2>&nbsp;<%=DataRS("TotalDuration")%></font>&nbsp;</td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("TotalCost"),-1) %>&nbsp;</font>&nbsp;</td> 
	  	 </TR>
	<%   
			Count=Count +1
	 		DataRS.movenext
	   		no_ = no_ + 1
		loop
	%>
	</table>
	<table align="center" cellpadding="1" cellspacing="0" width="100%">
		<tr>
			<td align="right">
	<%
			Do while PageNo<=TotalPages 
				if trim(pageNo) = trim(PageIndex) Then
	%>		
					<label class="ActivePage"><%=PageNo%></label>&nbsp;
				<%Else%>
					<a href="LineUsagesReport.asp?PageIndex=<%=PageNo%>&Post=<%=Post_%>&StartDate=<%=StartDate_ %>&EndDate=<%=EndDate_%>&SortBy=<%=SortBy_%>&Order=<%=Order_%>"><%=PageNo%></a>&nbsp;
	<%	
				End If						
				PageNo=PageNo+1
			Loop
	%>
			</td>
		</tr>
	</table>

		<script language="JavaScript">
	    	    var cal0 = new calendar1(document.forms['frmSearch'].elements['txtStartDate']);
			cal0.year_scroll = true;
			cal0.time_comp = false;
	    	    var cal1 = new calendar1(document.forms['frmSearch'].elements['txtEndDate']);
			cal1.year_scroll = true;
			cal1.time_comp = false;
		</script>

	</form>
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


