<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>

<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!--#include file="connect.inc" -->

<script language="JavaScript" src="calendar.js"></script>
<%
'if (session("Month") = "") or (session("Year") = "") then
'	strsql = "Select MonthP, YearP From Period"
	'response.write strsql & "<br>"
'	set rsData = server.createobject("adodb.recordset") 
'	set rsData = BillingCon.execute(strsql)
'	if not rsData.eof then
'		session("Month") = rsData("MonthP")
'		session("Year") = rsData("YearP")
'	end if
'end if

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

Post_ = Request.Form("PostList")
if Post_ ="" then
	If Request("Post")<> "" Then
		Post_ = Request("Post")	
	Else
		Post_ = "A"
	End If
end if

Agency_ = Request.Form("AgencyList")
if Agency_ ="" then
	If Request("Agency")<> "" Then
		Agency_ = Request("Agency")
	Else
		Agency_ = "A"
	End If
end if

Office_ = Request.Form("OfficeList")
if Office_ ="" then
	If Request("Office") <> "" Then
		Office_ = Request("Office")
	Else
		Office_ = "A"
	End If
end if

EmpId_ = Request.Form("EmpList")
if EmpId_ ="" then
	If Request("EmpId") <> "" Then
		EmpId_ = Request("EmpId")
	Else
		EmpId_ = "A"
	End If
end if

PhoneType_ = Request.Form("PhoneTypeList")
'response.write "PhoneType_  :" & PhoneType_ 
if PhoneType_ ="" then
	If Request("PhoneType")<> "" Then
		PhoneType_ = Request("PhoneType")
	Else
		PhoneType_ = "A"
	End If
end if

CallType_ = Request.Form("CallTypeList")
'response.write "CallType :" & CallType_ 
if CallType_ ="" then
	If Request("CallType") <> "" Then
		CallType_ = Request("CallType")
	Else
		CallType_ = "A"
	End If
end if

SortBy_ = Request.Form("SortList")

if (SortBy_ ="") then
	if Request("SortBy")<>"" then
		SortBy_ = Request("SortBy")	
	Else
		SortBy_ = "EmpName"
	end if
end if
'response.write "SortBy" & SortBy_

Order_ = Request.Form("OrderList")
if (Order_ ="") then
	if Request("Order")<>"" then
		Order_ = Request("Order")	
	Else
		Order_ = "Asc"
	end if
end if
%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">LONG DISTANCE CALLS REPORT</TD>
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
intpageSize=50 
PageIndex=request("PageIndex")

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

If (UserRole_ <> "") Then
	strsql = "spLongDistanceCallsReport '1','" & StartDate_ & "','" & EndDate_ & "','" & Post_ & "','" & Agency_ & "','" & Office_ & "','" & EmpId_ & "','" & PhoneType_ & "','" & CallType_ & "','" & SortBy_ & "','" & Order_ & "'"
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
	<table cellspadding="1" cellspacing="0" width="80%" border="1" align="center">
	<tr align="Center">
		<td colspan="2" align="center">
			<form method="post" name="frmSearch" Action="LongDistanceCallsReport.asp">
			<table  width="100%">
			<tr bgcolor="#000099">
				<td height="25" colspan="7"><strong>&nbsp;<span class="style5">Search &amp; Sort By </span></strong></td>
			</tr>
			<tr>
				<td>&nbsp;Period&nbsp;</td>				
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
				<td width="15%">&nbsp;Post&nbsp;</td>
				<td>:</td>
		        	<td>
				<Select name="PostList">
					<Option value="A">-- All --</Option>
					<Option value="ZAGREB" <%if Post_ ="ZAGREB" then %>Selected<%End If%> >ZAGREB</Option>
					<Option value="PODGORICA" <%if Post_ ="PODGORICA" then %>Selected<%End If%> >PODGORICA</Option>
				</Select>
				</td>  					
				<td width="15%">&nbsp;Agency&nbsp;</td>
				<td>:</td>
		        	<td>
				<%
					strsql = "select Distinct Agency From vwPhoneCustomerList Where Len(rtrim(Agency))>0 order by Agency"
					set AgencyRS = server.createobject("adodb.recordset")
					set AgencyRS = BillingCon.execute(strsql)
				%>
					<select name="AgencyList"> 
						<Option value="A">-- All --</option>
				<%	do while not AgencyRS.eof %>
						<OPTION value=<%=AgencyRS("Agency")%> <% if trim(Agency_) = trim(AgencyRS("Agency")) then%> Selected <%end if%> /><%=AgencyRS("Agency")%>
				<%		AgencyRS.movenext	
					loop%>
					</Select>
				</td>
			</tr>
			<tr>
				<td width="15%">&nbsp;Office&nbsp;</td>
				<td>:</td>
		        	<td>
				<Select name="OfficeList">
				<%
					strsql = "select Distinct Office From vwPhoneCustomerList Where Len(rtrim(Office))>0 order by Office"
					set OfficeRS = server.createobject("adodb.recordset")
					set OfficeRS = BillingCon.execute(strsql)
				%>
						<Option value="A">-- All --</option>
				<%	do while not OfficeRS.eof %>
						<OPTION value=<%=OfficeRS("Office")%> <% if trim(Office_) = trim(OfficeRS("Office")) then%> Selected <%end if%> /><%=OfficeRS("Office")%>
				<%		OfficeRS.movenext	
					loop%>
				</Select>
				</td>  	
				<td width="15%">&nbsp;Employee&nbsp;</td>
				<td>:</td>
		        	<td>
				<%
					strsql = "Select EmpID, EmpName from vwPhoneCustomerList Order by EmpName"
					set EmpRS = server.createobject("adodb.recordset")
					set EmpRS = BillingCon.execute(strsql)
				%>
					<select name="EmpList"> 
						<Option value="A">-- All --</option>
				<%	do while not EmpRS.eof %>
						<OPTION value=<%=EmpRS("EmpID")%> <% if trim(EmpID_) = trim(EmpRS("EmpID")) then%> Selected <%end if%> /><%=EmpRS("EmpName")%>
				<%		EmpRS.movenext	
					loop%>
					</Select>
				</td>
			</tr>
			<tr>
				<td>&nbsp;Phone Type&nbsp;</td>
				<td>:</td>
				<td>
					<Select name="PhoneTypeList">
						<Option value="A">-- All --</Option>
						<Option value="Office Phone" <%if trim(PhoneType_) = "Office Phone" then %>Selected<%End If%>>Office Phone</Option>
<!--						<Option value="Cell Phone" <%if trim(PhoneType_) = "Cell Phone" then %>Selected<%End If%>>Cell Phone</Option>-->
						<Option value="Home Phone" <%if trim(PhoneType_) = "Home Phone" then %>Selected<%End If%>>Home Phone</Option>
					</Select>&nbsp;
				</td>
				<td>&nbsp;Call Type&nbsp;</td>				
				<td>:</td>
				<td>
					<Select name="CallTypeList">
						<Option value="A">-- All --</Option>
						<Option value="Y" <%if CallType_ ="Y" then %>Selected<%End If%> >Personal</Option>
						<Option value="N" <%if CallType_ ="N" then %>Selected<%End If%> >Official</Option>
					</Select>					
				</td>
			</tr>			
			<tr>
				<td>&nbsp;Sort By&nbsp;</td>
				<td>:</td>
				<td>
					<Select name="SortList">
						<Option value="DialedNumber" <%if SortBy_ ="DialedNumber" then %>Selected<%End If%> >Dialed Number</Option>
						<Option value="EmpName" <%if SortBy_ ="EmpName" then %>Selected<%End If%> >Employee Name</Option>
						<Option value="Cost" <%if SortBy_ ="Cost" then %>Selected<%End If%> >Cost</Option>
					</Select>&nbsp;
					<Select name="OrderList">
						<Option value="Asc" <%if Order_ ="Asc" then %>Selected<%End If%> >Asc</Option>
						<Option value="Desc" <%if Order_ ="Desc" then %>Selected<%End If%> >Desc</Option>
					</Select>

				</td>	
		               <td height="30" colspan="6" align="center">
					<input type="submit" name="Submit" value="Search  / Show All">
				</td>
        		</tr>
			</table>
			</form>
		</td>
	</tr>	
	</table>
	<form method="post" name="frmLongDistanceCallsList" action="LongDistanceCallsPrint.asp?StartDate=<%=StartDate_ %>&EndDate=<%=EndDate_%>&Post=<%=Post_%>&Agency=<%=Agency_%>&Office=<%=Office_%>&EmpId=<%=EmpId_%>&PhoneType=<%=PhoneType_%>&CallType=<%=CallType_%>&SortBy=<%=SortBy_%>&Order=<%=Order_%>">
	<%
		strsql = "spLongDistanceCallsReport '2','" & StartDate_ & "','" & EndDate_ & "','" & Post_ & "','" & Agency_ & "','" & Office_ & "','" & EmpId_ & "','" & PhoneType_ & "','" & CallType_ & "','" & SortBy_ & "','" & Order_ & "'"
		set sumRS = server.createobject("adodb.recordset")
		'response.write strsql & "<br>"
		set sumRS = BillingCon.execute(strsql)
	%>
	<table cellpadding="1" cellspacing="0" width="100%" class="FontText">
	<tr>
		<td>&nbsp;</td>
		<td><b>Total Call Duration</b></td>
		<td width="1%">:</td>
		<td><b><%=formatnumber(sumRS("TotalDuration"),-1)%></b>&nbsp;second(s)&nbsp;</td>
	</tr>
	<tr>
		<td width="50%">&nbsp;<input type="submit" value="Export to Excel" /></td>
		<td><b>Total Cost</b></td>
		<td width="1%">:</td>
		<td><b>Kn. <%=formatnumber(sumRS("TotalCost"),-1)%></b>&nbsp;</td>
	</tr>
	</table>
	<table align="center" cellpadding="1" cellspacing="0" width="100%" border="1" bordercolor="black"  class="FontText"> 
	<TR BGCOLOR="#000099" align="center" cellpadding="0" cellspacing="0" >
		<TD width="4%"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
		<TD><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
	       	<TD width="10%"><strong><label STYLE=color:#FFFFFF>Phone Type</label></strong></TD>
	       	<TD width="14%"><strong><label STYLE=color:#FFFFFF>Dialed Number</label></strong></TD>
	       	<TD width="14%"><strong><label STYLE=color:#FFFFFF>Call date & time</label></strong></TD>
		<TD width="10%"><strong><label STYLE=color:#FFFFFF>Duration (Second)</label></strong></TD>
		<TD width="10%"><strong><label STYLE=color:#FFFFFF>Cost</label></strong></TD>
		<TD width="10%"><strong><label STYLE=color:#FFFFFF>Call Type</label></strong></TD>
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
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("EmpName")%></font></td> 
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("PhoneType")%></font></td> 
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("DialedNumber")%></font></td> 
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("DialedDateTime")%></font></td> 
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("CallDurationSecond")%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("Cost"),-1)%>&nbsp;</font></td> 
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("CallType")%></font></td> 
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
					<a href="LongDistanceCallsReport.asp?PageIndex=<%=PageNo%>&StartDate=<%=StartDate_ %>&EndDate=<%=EndDate_%>&Post=<%=Post_%>&Agency=<%=Agency_%>&Office=<%=Office_%>&EmpId=<%=EmpId_%>&PhoneType=<%=PhoneType_%>&CallType=<%=CallType_%>&SortBy=<%=SortBy_%>&Order=<%=Order_%>"><%=PageNo%></a>&nbsp;
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


