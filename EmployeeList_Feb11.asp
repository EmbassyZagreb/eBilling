<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>
 
 
<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<!--#include file="connect.inc" -->
<style type="text/css">
<!--
.style5 {color: #FFFFFF;}
-->
</style>
<script language="JavaScript">
function ClearFilter()
{
	document.forms['frmSearch'].elements['PostList'].value ='A';
	document.forms['frmSearch'].elements['SectionList'].value ='A';
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
user1_ = right(user_,len(user_)-4)

Post_ = Request.Form("PostList")
'response.write "Post " & Post_ 
if (Post_  ="") then
	if Request("Post")<>"" then
		Post_ = Request("Post")	
	Else
		Post_ = "A"
	end if
end if

Section_ = Request.Form("SectionList")
'response.write "Section " & Section_ 
if (Section_  ="") then
	if Request("Section")<>"" then
		Section_ = Request("Section")	
	Else
		Section_ = "A"
	end if
end if

EmpName_ = trim(request.form("txtEmpName"))
if EmpName_ ="" then
	EmpName_ = Request("EmpName")	
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

Order_ = Request.Form("OrderList")
if (Order_ ="") then
	if Request("Order")<>"" then
		Order_ = Request("OrderList")	
	Else
		Order_ = "Asc"
	end if
end if

strsql = "select * from Users where loginId='" & user1_ & "'"
set RS_Query = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set RS_Query = BillingCon.execute(strsql)

  if not RS_Query.eof then 
     if (trim(RS_Query("RoleID")) = "Admin") or (trim(RS_Query("RoleID")) = "IM") or (trim(RS_Query("RoleID")) = "FMC") then     
	strsql = "Select EmployeesID As EmpID, EmpName, PostName, Agency, Office, Type, ReportToID, ReportToName, EmailAddress "_
		  &"From vwDirectReport Where (PostName='" & Post_ & "' or '" & Post_ & "'='A') "_
		  &"And (Office='" & Section_ & "' or '" & Section_ & "'='A') "_
		  &"And (EmpName like '%" & EmpName_  & "%' or '" & EmpName_ & "'='') "_
		  &"Order by " & SortBy_ & " " & Order_ 
		
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
					<Option value="(AID) JAKARTA" <%if Post_ ="(AID) JAKARTA" then %>Selected<%End If%> >(AID) JAKARTA</Option>
					<Option value="DENPASAR" <%if Post_ ="DENPASAR" then %>Selected<%End If%> >DENPASAR</Option>
					<Option value="JAKARTA" <%if Post_ ="JAKARTA" then %>Selected<%End If%> >JAKARTA</Option>
					<Option value="MEDAN" <%if Post_ ="MEDAN" then %>Selected<%End If%> >MEDAN</Option>
					<Option value="SURABAYA" <%if Post_ ="SURABAYA" then %>Selected<%End If%> >SURABAYA</Option>
				</Select>&nbsp;
			</td>
			<td align="right">Section&nbsp;</td>
			<td>:</td>
			<td>
<%
 				strsql ="select distinct Office from vwHomePhoneNumberList Where Office<>'' order by Office"
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
			<td align="right">Sort By&nbsp;</td>
			<td>:</td>
			<td>
				<Select name="SortList">
					<Option value="EmpName" <%if SortBy_ ="EmpName" then %>Selected<%End If%> >Employee Name</Option>
					<Option value="EmailAddress" <%if SortBy_ ="EmailAddress" then %>Selected<%End If%> >Email Address</Option>
					<Option value="PostName" <%if SortBy_ ="PostName" then %>Selected<%End If%> >Post</Option>
					<Option value="ReportToName" <%if SortBy_ ="ReportToName" then %>Selected<%End If%> >Supervisor</Option>
				</Select>&nbsp;
				<Select name="OrderList">
					<Option value="Asc" <%if Order_ ="Asc" then %>Selected<%End If%> >Asc</Option>
					<Option value="Desc" <%if Order_ ="Desc" then %>Selected<%End If%> >Desc</Option>
				</Select>
			</td
		</tr>
		<tr>
			<td colspan="3"></td>
			<td><input type="submit" name="btnSearch" value="Search"></td>
			<td>&nbsp;</td>
			<td>&nbsp;&nbsp;<input type="button" name="btnClear" value="Reset filter" onclick="javascript:ClearFilter();">
<!--			    &nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="btnExport" value="Export to Excel" onClick="javascript:document.location.href('HomePhoneNumberListPrint.asp?Post=<%=Post_%>&Section=<%=Section_%>&EmpName=<%=EmpName_%>&SortBy=<%=SortBy_%>&Order=<%=Order_%>');"/> -->
			</td>
		</tr>
		</table></form>
		</td>
	</tr>
	</table>
     <form name="frmHomePhoneList" Action="HomePhoneListAll.asp" onSubmit="return ValidateForm()">
     <table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="90%">
     <TR BGCOLOR="#330099" align="center">
         <TD width="3%" align="center"><strong><label STYLE=color:#FFFFFF>NO</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>&nbsp;Employee Name</label></strong></TD>
         <TD width="12%"><strong><label STYLE=color:#FFFFFF>&nbsp;Post</label></strong></TD>
	 <TD width="12%"><strong><label STYLE=color:#FFFFFF>&nbsp;Office</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Type</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>&nbsp;Email Address</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>&nbsp;Supervisor</label></strong></TD>
     </TR>    
<% 
	   dim no_  
	   'no_ = 1 
	   no_ = 1 + ((PageIndex*intPageSize)-intPageSize)
	   do while not rs.eof and Count<intPageSize
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
     <TR bgcolor="<%=bg%>">
        <TD align=center ><FONT color=#330099 size=2> <%= no_ %> </font>   </TD>
        <TD><FONT color=#330099 size=2><A HREF="EmployeeEdit.asp?EmpID=<%= rs("EmpID")%>&Type=<%= rs("Type")%>"> <%= rs("EmpName") %></A></font>   </TD>
        <TD><FONT color=#330099 size=2>&nbsp;<%= rs("PostName") %></font>   </TD>
        <TD><FONT color=#330099 size=2>&nbsp;<%= rs("Office") %></font>   </TD>
        <TD><FONT color=#330099 size=2>&nbsp;<%= rs("Type") %></font>   </TD>
        <TD><FONT color=#330099 size=2>&nbsp;<%= rs("EmailAddress") %></font>   </TD>
        <TD><FONT color=#330099 size=2>&nbsp;<%= rs("ReportToName") %></font>   </TD>
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
					<a href="EmployeeList.asp?PageIndex=<%=PageNo%>&Post=<%=Post_%>&Section=<%=Section_%>&EmpName=<%=EmpName_%>&SortBy=<%=SortBy_%>&Order=<%=Order_%>"><%=PageNo%></a>&nbsp;
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
			<td>Please <a href="/CSC">Submit Request </a> or contact Jakarta CSC Helpdesk at ext.9111.</td>
		</tr>
	</table>
<%   end if 
else %>
	<table>
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="/CSC">Submit Request </a> or contact Jakarta CSC Helpdesk at ext.9111.</td>
		</tr>
	</table>

<%
end if 
%>
	</form>
</body> 

</html>


