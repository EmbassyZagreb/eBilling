<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>

<html>
<head>
<%
Response.ContentType ="application/vnd.ms-excel" 
Response.Buffer  =  True 
Response.Clear() 
%> 
<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!--#include file="connect.inc" -->

<script language="JavaScript" src="calendar.js"></script>
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

StartDate_ = Request("StartDate")
EndDate_ = Request("EndDate")
Post_ = Request("Post")	
Agency_ = Request("Agency")
Office_ = Request("Office")
EmpId_ = Request("EmpId")
PhoneType_ = Request("PhoneType")
CallType_ = Request("CallType")
SortBy_ = Request("SortBy")	
Order_ = Request("Order")	
%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
<table align="center" cellspadding="1" cellspacing="0" width="100%">  
<tr>
	<td colspan="8" align="center"><h3>Long Distance Call</h3></td>
</tr>
<tr>
	<td colspan="8" align="center">Billing Period : <Label style="color:blue"><%= StartDate_ %> - <%= EndDate_ %></lable></td>
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
	'DataRS.CursorLocation = 3
	'DataRS.open strsql,BillingCon
	set DataRS = BillingCon.execute(strsql)

%>
	<form method="post" name="frmLongDistanceCallsList" action="LongDistanceCallsPrint.asp?Post=<%=Post_%>&Agency=<%=Agency_%>&StartDate=<%=StartDate_ %>&EndDate=<%=EndDate_%>&CallType=<%=CallType_%>&SortBy=<%=SortBy_%>&Order=<%=Order_%>">
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
		<td><b><%=formatnumber(sumRS("TotalDuration"),-1)%></b>second(s)</td>
	</tr>
	<tr>
		<td width="50%">&nbsp;</td>
		<td><b>Total Cost</b></td>
		<td width="1%">:</td>
		<td><b>Kn. <%=formatnumber(sumRS("TotalCost"),-1)%></b></td>
	</tr>
	</table>
	<table align="center" cellpadding="1" cellspacing="0" width="100%" border="1" bordercolor="black"  class="FontText"> 
	<TR BGCOLOR="#000099" align="center" cellpadding="0" cellspacing="0" >
		<TD width="4%" class="style5"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
		<TD class="style5"><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
	       	<TD width="10%" class="style5"><strong><label STYLE=color:#FFFFFF>Phone Type</label></strong></TD>
	       	<TD width="14%" class="style5"><strong><label STYLE=color:#FFFFFF>Dialed Number</label></strong></TD>
	       	<TD width="14%" class="style5"><strong><label STYLE=color:#FFFFFF>Call date & time</label></strong></TD>
		<TD width="10%" class="style5"><strong><label STYLE=color:#FFFFFF>Duration (Second)</label></strong></TD>
		<TD width="10%" class="style5"><strong><label STYLE=color:#FFFFFF>Cost</label></strong></TD>
		<TD width="10%" class="style5"><strong><label STYLE=color:#FFFFFF>Call Type</label></strong></TD>
	</TR>    
	<% 
		dim no_  
		no_ = 1 
		do while not DataRS.eof
	   	if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
	%>
	   	<TR bgcolor="<%=bg%>">
			<td align="right"><%=No_%></td>
	        	<td><FONT color=#330099 size=2><%=DataRS("EmpName")%></font></td> 
	        	<td><FONT color=#330099 size=2><%=DataRS("PhoneType")%></font></td> 
	        	<td><FONT color=#330099 size=2><%=DataRS("DialedNumber")%></font></td> 
	        	<td><FONT color=#330099 size=2><%=DataRS("DialedDateTime")%></font></td> 
	        	<td><FONT color=#330099 size=2><%=DataRS("CallDurationSecond")%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("Cost"),-1)%></font></td> 
	        	<td><FONT color=#330099 size=2><%=DataRS("CallType")%></font></td> 
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
					<label class="ActivePage"><%=PageNo%></label>
				<%Else%>
					<a href="LongDistanceCallsReport.asp?PageIndex=<%=PageNo%>&StartDate=<%=StartDate_ %>&EndDate=<%=EndDate_%>&Post=<%=Post_%>&Agency=<%=Agency_%>&Office=<%=Office_%>&EmpId=<%=EmpId_%>&PhoneType=<%=PhoneType_%>&CallType=<%=CallType_%>&SortBy=<%=SortBy_%>&Order=<%=Order_%>"><%=PageNo%></a>
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


