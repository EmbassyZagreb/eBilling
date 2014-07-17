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
	document.forms['frmSearch'].elements['txtPhoneNumber'].value ='';
	document.forms['frmSearch'].elements['SortList'].value ='PhoneNumber';
}
</script>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
   <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">OFFICE PHONE NUMBER LIST</TD>
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

PhoneNumber_ = trim(request.form("txtPhoneNumber"))
if PhoneNumber_ ="" then
	PhoneNumber_ = Request("PhoneNumber")	
end if

Charge_ = trim(request.form("ChargeList"))
if Charge_ ="" then
	if Request("Charge")<>"" then
		Charge_ = Request("Charge")	
	Else
		Charge_ = "A"
	end if
end if

SortBy_ = Request.Form("SortList")
'response.write "SortBy" & SortBy_
if (SortBy_ ="") then
	if Request("SortBy")<>"" then
		SortBy_ = Request("SortBy")	
	Else
		SortBy_ = "PhoneNumber"
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
     if (trim(RS_Query("RoleID")) = "Admin") or (trim(RS_Query("RoleID")) = "IM") or (mid(RS_Query("RoleID"),1,3) = "FMC")  then     
	strsql = "Select ID, PhoneNumber, PhoneTypeName, EmpName, Post, Office, Case When Len(isNull(EmailAddress,''))<4 Then AlternateEmail Else isNull(EmailAddress,'') End As EmailAddress, BillFlag "_
		  &"From vwOfficePhoneNumberList Where (PhoneNumber like '" & PhoneNumber_  & "%' or '" & PhoneNumber_ & "'='') "_
		  &"And (Post='" & Post_ & "' or '" & Post_ & "'='A') "_
		  &"And (Office='" & Section_ & "' or '" & Section_ & "'='A') "_
		  &"And (BillFlag = '" & Charge_ & "' or '" & Charge_ & "'='A') "_
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
					<Option value="ZAGREB" <%if Post_ ="ZAGREB" then %>Selected<%End If%> >ZAGREB</Option>
					<Option value="PODGORICA" <%if Post_ ="PODGORICA" then %>Selected<%End If%> >PODGORICA</Option>
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
			<td colspan="4"><input name="txtEmpName" type="Input" size="50" Value='<%=EmpName_%>'></td>
		<tr>
		<tr>
			<td width="20%">&nbsp;Phone Number / Ext.</td>
			<td width="1%">:</td>
			<td><input name="txtPhoneNumber" type="Input" size="20" Value='<%=PhoneNumber_%>'></td>
			<td>Sort By&nbsp;</td>
			<td>:</td>
			<td>
				<Select name="SortList">
					<Option value="PhoneNumber" <%if SortBy_ ="PhoneNumber" then %>Selected<%End If%> >PhoneNumber</Option>
					<Option value="EmpName" <%if SortBy_ ="Location" then %>Selected<%End If%> >Employee Name / Location</Option>
				</Select>&nbsp;
				<Select name="OrderList">
					<Option value="Asc" <%if Order_ ="Asc" then %>Selected<%End If%> >Asc</Option>
					<Option value="Desc" <%if Order_ ="Desc" then %>Selected<%End If%> >Desc</Option>
				</Select>
			</td
		</tr>
		<tr>
			<td>&nbsp;Bill Charged</td>
			<td>:</td>
			<td colspan="4"><select name="ChargeList">
				<option value="A">-All-</option>
				<option value="Y" <%if Charge_ ="Y" then%>Selected<%End If%> >Yes</option>
				<option value="N" <%if Charge_ ="N" then%>Selected<%End If%> >No</option>
			  </select>
			</td>
		</tr>
		<tr>
			<td align="left">
				&nbsp;&nbsp;<input type="Button" name="btnBack" value="Back" onClick="Javascript:document.location.href('Default.asp');">
			</td>
			<td colspan="5" align="center">&nbsp;&nbsp;
				
				&nbsp;&nbsp;<input type="submit" name="btnSearch" value="Search">
				&nbsp;&nbsp;<input type="button" name="btnClear" value="Reset filter" onclick="javascript:ClearFilter();">


			</td>
		</tr>
		</table></form>
		</td>
	</tr>
	</table>
	
     <table width="100%">
	<tr>
		<td width="50%"><a href="OfficePhoneNumberEdit.asp?State=I">Add New Office Phone Number</a></td>
		<td align="right"><input type="button" name="btnExport" value="Export to Excel" onClick="javascript:document.location.href('OfficePhoneNumberListPrint.asp?Post=<%=Post_%>&Section=<%=Section_%>&EmpName=<%=EmpName_%>&PhoneNumber=<%=PhoneNumber_%>&Charge=<%=Charge_ %>&SortBy=<%=SortBy_%>&Order=<%=Order_%>');"/></td>
	</tr>
     </table>
     <table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width=100%>
     <TR BGCOLOR=#330099>
         <TD width="5%" align="center"><strong><label STYLE=color:#FFFFFF>NO</label></strong></TD>
         <TD width="10%"><strong><label STYLE=color:#FFFFFF>&nbsp;Phone Number</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>&nbsp;Employee Name / Location</label></strong></TD>
<!--         <TD><strong><label STYLE=color:#FFFFFF>Phone Type</label></strong></TD> -->
         <TD width="12%"><strong><label STYLE=color:#FFFFFF>&nbsp;Post</label></strong></TD>
	 <TD width="12%"><strong><label STYLE=color:#FFFFFF>&nbsp;Office</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Email Address / Alternate</label></strong></TD>
         <TD width="6%"><strong><label STYLE=color:#FFFFFF>&nbsp;Action</label></strong></TD>
	 <TD width="6%"><strong><label STYLE=color:#FFFFFF>&nbsp;Charged</label></strong>
<!--		<input type="checkbox" name="cbAll" value="true" onclick="checkall(this)" /> -->
	</TD>
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
        <TD><FONT color=#330099 size=2><A HREF="OfficePhoneNumberEdit.asp?ID=<%= rs("ID")%>&State=E"> <%= rs("PhoneNumber") %></A></font>   </TD>
        <TD><FONT color=#330099 size=2>&nbsp;<%= rs("EmpName") %></font>   </TD>
<!--        <TD><FONT color=#330099 size=2>&nbsp;<%= rs("PhoneTypeName") %></font>   </TD> -->
        <TD><FONT color=#330099 size=2>&nbsp;<%= rs("Post") %></font>   </TD>
        <TD><FONT color=#330099 size=2>&nbsp;<%= rs("Office") %></font>   </TD>
        <TD><FONT color=#330099 size=2>&nbsp;<%= rs("EmailAddress") %></font>   </TD>
	<TD><FONT color=#330099 size=2><A HREF="OfficePhoneNumberDeleteConfirm.asp?ID=<%= rs("ID")%>&State=D" >Delete</A></font>   </TD>
	<td align="center">
		<Input type="Checkbox" name="cbBillFlag" Value='<%=rs("ID")%>' <%if rs("BillFlag")="Y" then%> Checked <%end if%> disabled>
	</td>
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
					<a href="OfficePhoneNumberList.asp?PageIndex=<%=PageNo%>&Post=<%=Post_%>&Section=<%=Section_%>&EmpName=<%=EmpName_%>&PhoneNumber=<%=PhoneNumber_%>&Charge=<%=Charge_ %>&SortBy=<%=SortBy_%>&Order=<%=Order_%>"><%=PageNo%></a>&nbsp;
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
</body> 

</html>


