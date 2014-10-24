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
	document.forms['frmSearch'].elements['txtEmpName'].value ='';
	document.forms['frmSearch'].elements['PostList'].value ='A';
	document.forms['frmSearch'].elements['SectionList'].value ='A';
	document.forms['frmSearch'].elements['cmbSectionGroup'].value ='A';
	document.forms['frmSearch'].elements['ChargeList'].value ='Y';
	document.forms['frmSearch'].elements['SortList'].value ='PhoneNumber';
}
</script>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">CELL PHONE NUMBER LIST</TD>
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
		Charge_ = "Y"
	end if
end if

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

SectionGroup_ = Request.Form("cmbSectionGroup")
'response.write "Section " & SectionGroup_ 
if (SectionGroup_ ="") then
	if Request("SectionGroup")<>"" then
		SectionGroup_ = Request("SectionGroup")	
	Else
		SectionGroup_ = "A"
	end if
end if

Discontinued_ = Request.Form("cmbDiscontinued")
'response.write "Section " & Discontinued_ 
if (Discontinued_  ="") then
	if Request("Discontinued")<>"" then
		Discontinued_ = Request("Discontinued")	
	Else
		Discontinued_ = "N"
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
     if (trim(RS_Query("RoleID")) = "Admin") or (trim(RS_Query("RoleID")) = "IM") or (mid(RS_Query("RoleID"),1,3) = "FMC") then     
	strsql = "Select ID, PhoneNumber, PhoneTypeName, EmpName, Post, Office, Case When Len(isNull(EmailAddress,''))<4 Then AlternateEmail Else isNull(EmailAddress,'') End As EmailAddress, OwnerName, BillFlag, Discontinued, DiscontinuedDesc, DiscontinuedDate, ExistInMonthlyBilling "_
		  &"From vwCellPhoneNumberList Where (PhoneNumber like '%" & PhoneNumber_  & "%' or '" & PhoneNumber_ & "'='') "_
		  &"And (EmpName like '%" & EmpName_  & "%' or '" & EmpName_ & "'='') "_
		  &"And (Post='" & Post_ & "' or '" & Post_ & "'='A') "_
		  &"And (Office='" & Section_ & "' or '" & Section_ & "'='A') "_
		  &"And (SectionGroup='" & SectionGroup_ & "' or '" & SectionGroup_ & "'='A') "_
		  &"And (BillFlag = '" & Charge_ & "' or '" & Charge_ & "'='A') "_
		  &"And (Discontinued='" & Discontinued_ & "' or '" & Discontinued_ & "'='A') "_
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
					<Option value="">--All--</Option>
<%
				strsql ="select distinct Post from vwPhoneCustomerList Where Post<>'' order by Post"
				set SectionRS = server.createobject("adodb.recordset")
				set SectionRS = BillingCon.execute(strsql)
				Do While not SectionRS.eof 
%>
					<Option value='<%=SectionRS("Post")%>' <%if trim(Post_) = trim(SectionRS("Post")) then %>Selected<%End If%> ><%=SectionRS("Post")%></Option>
<%					
				SectionRS.MoveNext
				Loop%>
			</select>
			</td>

			<td align="right">Section&nbsp;</td>
			<td>:</td>
			<td>
<%
 				strsql ="select distinct Office from vwCellPhoneNumberList Where Office<>'' order by Office"
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
			<td><input name="txtEmpName" type="Input" size="40" Value='<%=EmpName_%>'></td>
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
		<tr>
		<tr>
			<td width="20%">&nbsp;Phone Number</td>
			<td width="1%">:</td>
			<td><input name="txtPhoneNumber" type="Input" size="20" Value='<%=PhoneNumber_%>'></td>
			<td align="right">Discontinued</td>
			<td>:</td>
			<td>
				<Select name="cmbDiscontinued"">
					<Option value="A">-All-</Option>
					<Option value="Y" <%if Discontinued_ ="Y" then %>Selected<%End If%> >Yes</Option>
					<Option value="N" <%if Discontinued_ ="N" then %>Selected<%End If%> >No</Option>
				</Select>&nbsp;
			</td>
		</tr>
		<tr>
			<td>&nbsp;Bill Charged</td>
			<td>:</td>
			<td><select name="ChargeList">
				<option value="A">-All-</option>


				<option value="Y" <%if Charge_ ="Y" Then %>Selected<%End If%>>Yes - Approved by supervisor</option>
				<option value="P" <%if Charge_ ="P" Then %>Selected<%End If%>>Personal phone - Full payment</option>
				<option value="N" <%if Charge_ ="N" Then %>Selected<%End If%>>No</option>

			  </select>
			</td>
			<td align="right">Sort By&nbsp;</td>
			<td>:</td>
			<td>
				<Select name="SortList">
					<Option value="EmpName" <%if SortBy_ ="EmpName" then %>Selected<%End If%> >Employee Name</Option>
					<Option value="EmailAddress" <%if SortBy_ ="EmailAddress" then %>Selected<%End If%> >Email Address</Option>
					<Option value="PhoneNumber" <%if SortBy_ ="PhoneNumber" then %>Selected<%End If%> >PhoneNumber</Option>
					<Option value="Post" <%if SortBy_ ="Post" then %>Selected<%End If%> >Post</Option>
					<Option value="Office" <%if SortBy_ ="Office" then %>Selected<%End If%> >Office</Option>
				</Select>&nbsp;
				<Select name="OrderList">
					<Option value="Asc" <%if Order_ ="Asc" then %>Selected<%End If%> >Asc</Option>
					<Option value="Desc" <%if Order_ ="Desc" then %>Selected<%End If%> >Desc</Option>
				</Select>
			</td
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
		<table width="70%">
			<tr>
				<td class="Hint" align="left">*Cellphone numbers found in historical data cannot be Deleted, only set as Discontinued.</td>
			</tr>
		</table>
	</table>
	
     <table width="100%">
	<tr>
		<td width="50%"><a href="CellPhoneNumberEdit.asp?State=I">Add New CellPhone Number</a></td>
		<td align="right"><input type="button" name="btnExport" value="Export to Excel" onClick="javascript:document.location.href('CellPhoneNumberListPrint.asp?Post=<%=Post_%>&Section=<%=Section_%>&SectionGroup=<%=SectionGroup_%>&EmpName=<%=EmpName_%>&PhoneNumber=<%=PhoneNumber_%>&Charge=<%=Charge_ %>&Discontinued=<%=Discontinued_%>&SortBy=<%=SortBy_%>&Order=<%=Order_%>');"/></td>
	</tr>
     </table>
     <table border="1" bordercolor="#EEEEEE" cellpadding="2" cellspacing="0" width="100%">
     <TR BGCOLOR="#330099" align="center">
         <TD width="4%" align="center"><strong><label STYLE=color:#FFFFFF>NO</label></strong></TD>
         <TD width="8%"><strong><label STYLE=color:#FFFFFF>&nbsp;Phone Number</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>&nbsp;Employee Name</label></strong></TD>
<!--         <TD><strong><label STYLE=color:#FFFFFF>Phone Type</label></strong></TD>-->
         <TD width="8%"><strong><label STYLE=color:#FFFFFF>&nbsp;Post</label></strong></TD>
         <TD width="12%"><strong><label STYLE=color:#FFFFFF>&nbsp;Office</label></strong></TD>
         <TD width="15%"><strong><label STYLE=color:#FFFFFF>Email Address / Alternate</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Owner</label></strong></TD>
         <TD width="4%"><strong><label STYLE=color:#FFFFFF>&nbsp;Action</label></strong></TD>
	 <TD width="6%" align="center"><strong><label STYLE=color:#FFFFFF>&nbsp;Charged</label></strong>
<!--		<input type="checkbox" name="cbAll" value="true" onclick="checkall(this)" /> -->
	</TD>
         <TD><strong><label STYLE=color:#FFFFFF>Discontinued</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Discontinued Date</label></strong></TD>
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
        <TD><A HREF="CellPhoneNumberEdit.asp?ID=<%= rs("ID")%>&State=E"> <%= rs("PhoneNumber") %></A></font>   </TD>
        <TD>&nbsp;<%= rs("EmpName") %></font>   </TD>
        <TD>&nbsp;<%= rs("Post") %></font>   </TD>
        <TD>&nbsp;<%= rs("Office") %></font>   </TD>
        <TD>&nbsp;<%= rs("EmailAddress") %></font>   </TD>
        <TD>&nbsp;<%= rs("OwnerName") %></font>   </TD>
		
		<TD>
<%			If rs("ExistInMonthlyBilling")="N" Then %>				
				<A HREF="CellPhoneNumberDeleteConfirm.asp?ID=<%= rs("ID")%>&State=D" >Delete</A>
<%			else %>	
				&nbsp;
<%			end if %>
		</TD>
	<td align="center">
<!--		<Input type="Checkbox" name="cbBillFlag" Value='<%=rs("ID")%>' <%if rs("BillFlag")="Y" then%> Checked <%end if%> disabled> -->
		<%if rs("BillFlag")="Y" then%> Yes <%end if%>
		<%if rs("BillFlag")="P" then%> Personal <%end if%>
		<%if rs("BillFlag")="N" then%> No <%end if%>
	</td>
	<td align="center">
		<%if rs("Discontinued")="Y" then%> Yes <%else%> No <%end if%>
	</td>
<!--    <TD>&nbsp;<%= rs("Discontinued") %></font>   </TD> -->
        <TD>&nbsp;<%= rs("DiscontinuedDate") %></font>   </TD>
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
					<a href="CellPhoneNumberList.asp?PageIndex=<%=PageNo%>&Post=<%=Post_%>&Section=<%=Section_%>&SectionGroup=<%=SectionGroup_%>&EmpName=<%=EmpName_%>&PhoneNumber=<%=PhoneNumber_%>&Charge=<%=Charge_ %>&SortBy=<%=SortBy_%>&Order=<%=Order_%>"><%=PageNo%></a>&nbsp;
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


