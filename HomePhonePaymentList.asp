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
	document.forms['frmSearch'].elements['txtEmptName'].value ="";
	document.forms['frmSearch'].elements['cmbOfficeSection'].value ="";
	document.forms['frmSearch'].elements['cmbStatus'].value ="X";

}

</script>

<%
curMonth_ = month(date())
curYear_ = year(date())
if len(curMonth_)= 1 then
	curMonth_ = "0" & curMonth_
end if

MonthP_ = Request.Form("MonthList")
if MonthP_ ="" then
	if request("MonthP") <> "" then
		MonthP_ = request("MonthP")
	else
		MonthP_ = curMonth_ 
	end if
end if

YearP_ = Request.Form("YearList")
if YearP_ ="" then
	if request("YearP") <> "" then
		YearP_ = request("YearP")
	else
		YearP_ = curYear_
	end if
end if


dim intPageSize,PageIndex,TotalPages 
dim RecordCount,RecordNumber,Count 
intpageSize=20 
PageIndex=request("PageIndex")

EmpName_ = Trim(Request.Form("txtEmpName"))
if EmpName_ ="" then
	EmpName_ = Trim(request("EmpName"))
End If
'response.write EmpName_
OfficeSection_ = Trim(Request.Form("cmbOfficeSection"))
if OfficeSection_ ="" then
	OfficeSection_ = Trim(request("OfficeSection"))
End If
'response.write OfficeSection_ 
Outstanding_ = Trim(Request.Form("txtOutstanding"))
if Outstanding_ ="" then
	Outstanding_ = Trim(request("Outstanding"))
End If

'response.write Outstanding_
if Outstanding_ = "" then
	Outstanding_ = 0
end if

Status_ = Trim(Request.Form("cmbStatus"))
if Status_ ="" then
	Status_ = Trim(request("Status"))
End If
'response.write Status_

if Status_ = "" then
	Status_ ="X"
end if
%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">HOME PHONE PAYMENT LIST</TD>
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
dim tombol
dim hlm
%>

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

If (UserRole_ <> "") Then

	strsql = "spGetPaymentList '2','" & MonthP_ & "','" & YearP_ & "','','" & EmpName_ & "','" & OfficeSection_ & "'," & Outstanding_ & ",'" & Status_ & "'"
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
	<form method="post" name="frmSearch" id="frmSearch">
	<table align="center" cellpadding="1" cellspacing="0" width="70%">
	<tr bgcolor="#000099">
		<td height="25" colspan="4"><strong>&nbsp;<span class="style5">Search</span></strong></td>
	</tr>
	<tr>
		<td width="20%">Period :</td>				
		<td colspan="3">
			<Select name="MonthList">
				<Option value="01" <%if MonthP_ ="01" then %>Selected<%End If%> >January</Option>
				<Option value="02" <%if MonthP_ ="02" then %>Selected<%End If%> >February</Option>
				<Option value="03" <%if MonthP_ ="03" then %>Selected<%End If%> >March</Option>
				<Option value="04" <%if MonthP_ ="04" then %>Selected<%End If%> >April</Option>
				<Option value="05" <%if MonthP_ ="05" then %>Selected<%End If%> >May</Option>
				<Option value="06" <%if MonthP_ ="06" then %>Selected<%End If%> >June</Option>
				<Option value="07" <%if MonthP_ ="07" then %>Selected<%End If%> >July</Option>
				<Option value="08" <%if MonthP_ ="08" then %>Selected<%End If%> >August</Option>
				<Option value="09" <%if MonthP_ ="09" then %>Selected<%End If%> >September</Option>
				<Option value="10" <%if MonthP_ ="10" then %>Selected<%End If%> >October</Option>
				<Option value="11" <%if MonthP_ ="11" then %>Selected<%End If%> >November</Option>
				<Option value="12" <%if MonthP_ ="12" then %>Selected<%End If%> >December</Option>
			</Select>&nbsp;
<%
			Year_ = Year(Date()) - 1
'					response.write YearP_
%>

			<Select name="YearList">
<% 				Do While Year_ <= Year(Date()) %>
			<Option value='<%=Year_%>' <%if trim(Year_) = trim(YearP_) then %>Selected<%End If%> ><%=Year_%></Option>		
<% 
		Year_ = Year_ + 1
		Loop %>	
			</Select>										
		</td>

	</tr>
	<tr>
		<td>Employee Name :</td>
		<td colspan="3"><input name="txtEmpName" type="Input" size="30" Value='<%=EmpName_%>'></td>
	</tr>
	<tr>
		<td>Office Section :</td>
		<td>
			<select name="cmbOfficeSection">
				<option value="">--All--</option>
				<% Dim OfficeRS
				    strsql = "select distinct OfficeLocation from employees Where Status='C' and len(OfficeLocation)>1 Order by OfficeLocation"
				    set OfficeRS = server.createobject("adodb.recordset")    
	   			    set OfficeRS = JakEmpCon.execute(strsql) 
				    
				    Do while not OfficeRS.eof
%>
					<option value="<%=OfficeRS("OfficeLocation")%>" <%if OfficeSection_ = OfficeRS("OfficeLocation") then%>Selected<%End If%>><%=OfficeRS("OfficeLocation")%></option>
<%
					OfficeRS.MoveNext
				    loop
				%>
			</select>
		</td>
		<td>Status :</td>
		<td>
			<select name="cmbStatus">
				<option value="X" <%if Status_ ="X" then%>Selected<%End If%>>--All--</option>
				<option value="N" <%if Status_ ="N" then%>Selected<%End If%>>Pending</option>
				<option value="P" <%if Status_ ="P" then%>Selected<%End If%>>Paid</option>
			</select>			
		</td>
	</tr>
	<tr>
		<td>Outstanding Payment :</td>
		<td colspan="3"><input name="txtOutstanding" type="Input" size="3" Value='<%=Outstanding_%>'>&nbsp;day(s)</td>
	</tr>
	<tr>
		<td colspan="2"></td>
		<td colspan="2">
			<input type="submit" name="btnSearch" value="Search">
			&nbsp;&nbsp;<input type="button" name="btnClear" value="Reset filter" onclick="javascript:ClearFilter();">
		</td>
	</tr>
	<tr>
		<td colspan="4"><hr></td>
	</tr>
</table>
</form>
<form method="post" name="frmPaymentList" action="" onSubmit="return ValidateCheckBox();">
<table align="center" cellpadding="1" cellspacing="0" width="100%" border="1" bordercolor="black"> 
<TR BGCOLOR="#000099" align="center" cellpadding="0" cellspacing="0" >
	<TD width="4%"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
	<TD><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
	<TD width="15%"><strong><label STYLE=color:#FFFFFF>Agency / Office</label></strong></TD>
       	<TD width="12%"><strong><label STYLE=color:#FFFFFF>HomePhone</label></strong></TD>
	<TD width="10%"><strong><label STYLE=color:#FFFFFF>Total Cost (Kn)</label></strong></TD>
	<TD width="10%"><strong><label STYLE=color:#FFFFFF>Paid Date</label></strong></TD>
	<TD width="10%"><strong><label STYLE=color:#FFFFFF>Outstanding (Day)</label></strong></TD>
	<TD width="8%"><strong><label STYLE=color:#FFFFFF>Paid Amount</label></strong></TD>
	<TD width="8%"><strong><label STYLE=color:#FFFFFF>Action</label></strong></TD>
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
        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("OfficeLocation")%></font></td> 
        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("PhoneNumber")%></font></td> 
        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("HomePhonePrsBillRp"),-1)%>&nbsp;</font></td> 
        	<td align="right"><FONT color=#330099 size=2><%=DataRS("PaidDate")%>&nbsp;</font></td> 
		<td align="right"><FONT color=#330099 size=2><%=DataRS("Outstanding")%>&nbsp;</font></td> 
        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("PaidAmount"),-1)%>&nbsp;</font></td> 
        	<td>&nbsp;
		<%If (DataRS("ProgressID")<=4) and (trim(DataRS("Status")) <>"Paid") Then %>			
			<A HREF="HomePhonePaymentApproval.asp?HomePhone=<%=DataRS("PhoneNumber")%>&PageIndex=<%=PageIndex%>&MonthP=<%=MonthP_%>&YearP=<%=YearP_%>" >Payment</A>
		<%Else%>
			&nbsp;<%=DataRS("Status")%>
		<%End If%>
		</td> 
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
				<a href="HomePhonePaymentList.asp?PageIndex=<%=PageNo%>&EmpName=<%=EmpName_%>&OfficeSection=<%=OfficeSection_%>&Outstanding=<%=Outstanding_%>&Status=<%=Status_%>&MonthP=<%=MonthP_%>&YearP=<%=YearP_%>"><%=PageNo%></a>&nbsp;
<%	
			End If						
			PageNo=PageNo+1
		Loop
%>
		</td>
	</tr>
</table>
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


