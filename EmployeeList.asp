 <%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>
 

<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<!--#include file="connect.inc" -->

<script language="JavaScript">
function ClearFilter()
{
	document.forms['frmSearch'].elements['PostList'].value ='A';
	document.forms['frmSearch'].elements['StatusList'].value ='A';
	document.forms['frmSearch'].elements['SectionList'].value ='A';
	document.forms['frmSearch'].elements['cmbSectionGroup'].value ='A';
	document.forms['frmSearch'].elements['txtEmpName'].value ='';
	document.forms['frmSearch'].elements['SortList'].value ='EmpName';
}

function ValidateForm()
{
	valid = true;
	nRec = 0;
	for (var x=0; x<frmHomePhoneList.elements.length; x++)
	{	
		cbElement = frmHomePhoneList.elements[x]
		if ((cbElement.checked) && (cbElement.name=="cbApproval"))
		{
			nRec++;
		}
	}
	if (nRec == 0)
	{
		alert("Please select data that you want to approve !!!");
		valid = false;
	}
	return valid;
}

function checkall(obj)
{
	for (var x=0; x<frmHomePhoneList.elements.length; x++)
	{
		cbElement = frmHomePhoneList.elements[x]
		if (cbElement.type == "checkbox")
		{
			cbElement.checked= obj.checked?true:false
		}
	}
}

</script>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">

<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate" />
<meta http-equiv="Pragma" content="no-cache" />
<meta http-equiv="Expires" content="0" />

<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">EMPLOYEE LIST</TD>
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

dim x , ticket_, user_ , user1_
dim intPageSize,PageIndex,TotalPages 
dim RecordCount,RecordNumber,Count 
intpageSize=50 
PageIndex=request("PageIndex")


user_ = request.servervariables("remote_user")
user1_ = user_  'user1_ = right(user_,len(user_)-4)

Post_ = Request.Form("PostList")
'response.write "Post " & Post_ 
if (Post_  ="") then
	if Request("Post") <>"" then
		Post_ = Request("Post")	
	Else
		Post_ = "A"
	end if
end if

Section_ = Request.Form("SectionList")
'response.write "Section " & Section_ 
if (Section_ ="") then
	if Request("Section")<>"" then
		Section_ = Request("Section")	
	Else
		Section_ = "A"
	end if
end if

SectionGroup_ = Request.Form("cmbSectionGroup")
'response.write "Section " & SectionGroup_ 
if (SectionGroup_ ="") then
	if Request("SectionGroup") <>"" then
		SectionGroup_ =  Request("SectionGroup")	
	Else
		SectionGroup_ = "A"
	end if
end if

EmpName_ = trim(request.form("txtEmpName"))
if EmpName_ ="" then
	EmpName_ = Request("EmpName")
	if EmpName_ ="" then EmpName_ =""
end if

SortBy_ = Request.Form("SortList")
'response.write "SortBy" & SortBy_
if (SortBy_ ="") then
	if Request("SortBy")<>"" then
		SortBy_ = Request("SortBy")
	Else
		SortBy_ = "EmpName"
	end if
end if

AgencyFundingCode_ = Request.Form("cmbAgencyFundingCode")
'response.write AgencyFundingCode_
if AgencyFundingCode_ = "" Then AgencyFundingCode_ = request("AgencyFundingCode")
if AgencyFundingCode_ = "" then
	AgencyFundingCode_ = 0
end if
'Response.write AgencyFundingCode_

Order_ = Request("OrderList")	
if (Order_ ="") then
	if Request.Form("OrderList")<>"" then
		Order_ = Request.Form("OrderList")
	Else
		Order_ = "Asc"
	end if
end if


Status_ = Request.Form("StatusList")
'response.write "Status" & Status_
if (Status_ ="") then
	if Request("Status") <>"" then
		Status_ = Request("Status")	
	Else
		Status_ = "A"
	end if
end if

strsql = "select * from Users where loginId='" & user1_ & "'"
set RS_Query = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set RS_Query = BillingCon.execute(strsql)

  if not RS_Query.eof then 
     if (trim(RS_Query("RoleID")) = "Admin") or (trim(RS_Query("RoleID")) = "IM") or (trim(RS_Query("RoleID")) = "FMC" or trim(user1_)= "PribanicM") then     
	strsql = "Select EmpID, EmpName, PostName, StatusName, Agency, Office, Type, ReportToID, ReportToName, EmailAddress, AgencyFunding, Remark, ExistInMonthlyBilling "_
		  &"From vwDirectReport Where (PostName='" & Post_ & "' or '" & Post_ & "'='A') "_
		  &"And (Office='" & Section_ & "' or '" & Section_ & "'='A') "_
		  &"And (SectionGroup='" & SectionGroup_ & "' or '" & SectionGroup_ & "'='A') "_
		  &"And (EmpName like '%" & EmpName_  & "%' or '" & EmpName_ & "'='') "_
		  &"And (AgencyId = " & AgencyFundingCode_ & " or " & AgencyFundingCode_ & "=0) "_
		  &"And (Status='" & Status_ & "' or '" & Status_ & "'='A') "_
		  &"Order by " & SortBy_ & " " & Order_ 

'		  &"And EmpID in (Select EmpID From MsCellPhoneNumber Where BillFlag='Y' ) "_
		
	'response.write strsql & "<br>"
	set rs = server.createobject("adodb.recordset") 
	rs.CursorLocation = 3
	rs.open strsql,BillingCon

	if ((PageIndex ="") or (request.form("btnSearch")="Search")) then PageIndex=1 
	if not rs.eof then
		RecordCount = rs.RecordCount   
		'response.write RecordCount & "<br>"
		RecordNumber=(intPageSize * PageIndex) - intPageSize 
		'response.write RecordNumber
		rs.PageSize =intPageSize 
		rs.AbsolutePage = PageIndex
		TotalPages=rs.PageCount 
		'response.write TotalPages & "<br>"
	End If
'	set rs = BillingCon.execute(strsql) 



%>
	
	<table align="center" border="1" cellpadding="1" cellspacing="0" width="70%" bgcolor="white">
	<tr>
		<td>
		<form method="post" name="frmSearch">
		<table align="center" cellpadding="1" cellspacing="0" width="100%">		
		<tr bgcolor="#000099">
			<td height="25" colspan="6"><strong>&nbsp;<span class="style5">Search</span></strong></td>
		</tr>
		<tr>
			<td>&nbsp;Post</td>
			<td>:</td>
			<td>
				<Select name="PostList">
					<Option value="A">-All-</Option>
					<Option value="ZAGREB" <%if Post_ ="ZAGREB" then %>Selected<%End If%> >ZAGREB</Option>
					<Option value="SARAJEVO" <%if Post_ ="SARAJEVO" then %>Selected<%End If%> >SARAJEVO</Option>
				</Select>&nbsp;
			</td>
			<td align="right">Office&nbsp;</td>
			<td>:</td>
			<td>
<%
 				'strsql ="select distinct Office from vwHomePhoneNumberList Where Office<>'' order by Office"
				strsql ="select distinct Office from vwPhoneCustomerList Where Office<>'' order by Office"
				set SectionRS = server.createobject("adodb.recordset")
				set SectionRS = BillingCon.execute(strsql)
'				response.write strStr 
%>	
				<Select name="SectionList">
					<Option value='A'>--All--</Option>
<%				Do While not SectionRS.eof %>
					<Option value='<%=SectionRS("Office")%>' <%if trim(Section_) = trim(SectionRS("Office")) then %>Selected<%End If%> ><%=SectionRS("Office")%></Option>
					
<%					SectionRS.MoveNext
				Loop%>
				</select>

			</td>			
		</tr>
		<tr>
			<td width="20%">&nbsp;Employee Name</td>
			<td width="1%">:</td>
			<td><input name="txtEmpName" type="Input" size="30" Value='<%=EmpName_%>'></td>
			<td align="right">Section by Group&nbsp;</td>
			<td>:</td>
			<td>
<%
 				strsql ="select distinct SectionGroup from vwPhoneCustomerList Where SectionGroup<>'' order by SectionGroup"
				set SectionGroupRS = server.createobject("adodb.recordset")
				set SectionGroupRS = BillingCon.execute(strsql)
'				response.write strStr 
%>	
				<Select name="cmbSectionGroup">
					<Option value='A'>--All--</Option>
<%				Do While not SectionGroupRS.eof %>
					<Option value='<%=SectionGroupRS("SectionGroup")%>' <%if trim(SectionGroup_) = trim(SectionGroupRS("SectionGroup")) then %>Selected<%End If%> ><%=SectionGroupRS("SectionGroup")%></Option>
					
<%					SectionGroupRS.MoveNext
				Loop%>
				</select>
			</td>
		</tr>
		<tr>
			<td>&nbsp;Funding Agency&nbsp;</td>
			<td>:</td>
			<td>
<%
 				strsql ="select AgencyId, AgencyFundingCode, AgencyDesc from AgencyFunding order by AgencyDesc"
				set AgencyRS = server.createobject("adodb.recordset")
				set AgencyRS = BillingCon.execute(strsql)
'				response.write strStr 
%>	
				<Select name="cmbAgencyFundingCode">
					<Option value='0'>--All--</Option>
<%				Do While not AgencyRS.eof %>
					<Option value='<%=AgencyRS("AgencyId")%>' <%if trim(AgencyFundingCode_) = trim(AgencyRS("AgencyId")) then %>Selected<%End If%> ><%=AgencyRS("AgencyDesc")%></Option>
					
<%					AgencyRS.MoveNext
				Loop%>
				</select>

			</td>	
			<td align="right">Sort By&nbsp;</td>
			<td>:</td>
			<td>
				<Select name="SortList">
					<Option value="EmpName" <%if SortBy_ ="EmpName" then %>Selected<%End If%> >Employee Name</Option>
					<Option value="EmailAddress" <%if SortBy_ ="EmailAddress" then %>Selected<%End If%> >Email Address</Option>
					<Option value="AgencyFunding" <%if SortBy_ ="AgencyFunding" then %>Selected<%End If%> >Funding Agency</Option>
					<Option value="PostName" <%if SortBy_ ="PostName" then %>Selected<%End If%> >Post</Option>
					<Option value="ReportToName" <%if SortBy_ ="ReportToName" then %>Selected<%End If%> >Supervisor</Option>
					<Option value="Type" <%if SortBy_ ="Type" then %>Selected<%End If%> >Type</Option>
				</Select>&nbsp;
				<Select name="OrderList">
					<Option value="Asc" <%if Order_ ="Asc" then %>Selected<%End If%> >Asc</Option>
					<Option value="Desc" <%if Order_ ="Desc" then %>Selected<%End If%> >Desc</Option>
				</Select>
			</td
		</tr>
		<tr>
			<td>&nbsp;Status</td>
			<td>:</td>
			<td>
				<Select name="StatusList">
					<Option value="A">-All-</Option>
					<Option value="C" <%if Status_ ="C" then %>Selected<%End If%> >Current</Option>
					<Option value="D" <%if Status_ ="D" then %>Selected<%End If%> >Departed</Option>
				</Select>&nbsp;
			</td>
			<td colspan="2"><input type="submit" name="btnSearch" value="Search">&nbsp;<input type="button" name="btnClear" value="Reset filter" onclick="javascript:ClearFilter();">
			</td>
			<td><input type="button" name="btnExport" value="Export to Excel" onClick="javascript:document.location.href('EmployeeListPrint.asp?Post=<%=Post_%>&Status=<%=Status_%>&Section=<%=Section_%>&SectionGroup=<%=SectionGroup_%>&EmpName=<%=EmpName_%>&AgencyFundingCode=<%=AgencyFundingCode_%>&SortBy=<%=SortBy_%>&Order=<%=Order_%>');"/>
			</td>
		</tr>
		</table></form>
		</td>
	</tr>
	</table>
	<table width="70%">
			<tr>
				<td class="Hint" align="left">*Employee assigned in historical data cannot be Deleted, only set as Departed.</td>
			</tr>
	</table>	
     <form name="frmHomePhoneList" Action="HomePhoneListAll.asp" onSubmit="return ValidateForm()">
	<table width="90%">
		<tr>
			<td><a href="EmployeeEdit.asp?State=I">Add New employee</a></td>
		</tr>
	</table>
     <table border="1" bordercolor="#EEEEEE" cellpadding="2" cellspacing="0" width="100%">
     <TR BGCOLOR="#330099" align="center">
         <TD width="3%" align="center"><strong><label STYLE=color:#FFFFFF>NO</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>&nbsp;Employee Name</label></strong></TD>
         <TD width="12%"><strong><label STYLE=color:#FFFFFF>&nbsp;Post</label></strong></TD>

	 <TD width="12%"><strong><label STYLE=color:#FFFFFF>&nbsp;Office</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Type</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>&nbsp;Email Address</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>&nbsp;Supervisor</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>&nbsp;Agency Funding</label></strong></TD>
	 <TD><strong><label STYLE=color:#FFFFFF>&nbsp;Remark</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>&nbsp;Status</label></strong></TD>	 
	 <TD><strong><label STYLE=color:#FFFFFF>&nbsp;Action</label></strong></TD>
     </TR>    
<% 
	   dim no_  
	   'no_ = 1 
	   no_ = 1 + ((PageIndex*intPageSize)-intPageSize)
	   do while not rs.eof and Count<intPageSize
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 

%>
      
     <TR bgcolor="<%=bg%>">
        <TD align=center > <%= no_ %> </font>   </TD>
        <TD><FONT color=#330099><A HREF="EmployeeEdit.asp?EmpID=<%= rs("EmpID")%>&Type=<%= rs("Type")%>&State=E"> <%= rs("EmpName") %></A></font>   </TD>
        <TD>&nbsp;<%= rs("PostName") %></font>   </TD>
        <TD>&nbsp;<%= rs("Office") %></font>   </TD>
        <TD>&nbsp;<%if rs("Type") ="AMER" then %>Supervisor<%Else%>Regular<%End If%></font>   </TD>
        <TD>&nbsp;<%= rs("EmailAddress") %></font>   </TD>
        <TD>&nbsp;<%= rs("ReportToName") %></font>   </TD>
        <TD>&nbsp;<%= rs("AgencyFunding") %></font>   </TD>
		<TD>&nbsp;<%= rs("Remark") %></font>   </TD>
	    <TD>&nbsp;<%= rs("StatusName") %></font>   </TD>
			<TD>
<%			If rs("ExistInMonthlyBilling")="N" Then %>				
				<A HREF="EmployeeDelete.asp?ID=<%= rs("EmpID")%>&Name='<%= rs("EmpName")%>'">Delete</A></font>
<%			else %>	
				&nbsp;
<%			end if %>
		</TD>
      </TR>

<%   
	   Count=Count +1
	   rs.movenext
	   no_ = no_ + 1 
	   loop
%>
     </TABLE>
     <table align="center" cellpadding="1" cellspacing="0" width="100%">
		<tr>
			<td align="right">
	<%
			Do while PageNo<=TotalPages 
				if trim(pageNo) = trim(PageIndex) Then
	%>		
					<label class="ActivePage"><%=PageNo%></label>&nbsp;
				<%Else%>
					<a href="EmployeeList.asp?PageIndex=<%=PageNo%>&Post=<%=Post_%>&Section=<%=Section_%>&SectionGroup=<%=SectionGroup_%>&EmpName=<%=EmpName_%>&AgencyFundingCode=<%=AgencyFundingCode_%>&SortBy=<%=SortBy_%>&Order=<%=Order_%>"><%=PageNo%></a>&nbsp;
	<%	
				End If						
				PageNo=PageNo+1
			Loop
	%>
			</td>
		</tr>
	</table>
     <table width="100%">
	<tr ><td align="Center">
	<a href="default.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a>
	</td></tr>
     </table>
<%
   else 
%>
      <table>
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
		</tr>
	</table>
<%   end if 
else %>
	<table>
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
		</tr>
	</table>

<%
end if 
%>
	</form>
</body> 

</html>


