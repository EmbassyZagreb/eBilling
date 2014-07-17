<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>
 
 
<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<!--#include file="connect.inc" -->
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
%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function ClearFilter()
{
	document.forms['frmSearch'].elements['txtPhoneNumber'].value ='';
	document.forms['frmSearch'].elements['PhoneTypeList'].value ='';
	document.forms['frmSearch'].elements['SortList'].value ='PhoneNumber';
}
</script>
</head>

<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">PHONE LIST</TD>
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
intpageSize=20 
PageIndex=request("PageIndex")


user_ = request.servervariables("remote_user")
user1_ = right(user_,len(user_)-4)

PhoneNumber_ = trim(request.form("txtPhoneNumber"))
if PhoneNumber_ ="" then
	PhoneNumber_ = Request("PhoneNumber")	
end if

PhoneType_ = trim(request.form("PhoneTypeList"))
if PhoneType_ ="" then
	if Request("PhoneType")<>"" then
		PhoneType_ = Request("PhoneType")	
	Else
		PhoneType_ = "A"
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
     if trim(RS_Query("RoleID")) = "Admin" then     
	strsql = "Select ID, PhoneNumber, Location, PhoneType, Post, Case When PhoneType='O' Then 'Office Phone' When PhoneType='H' Then 'Home Phone' Else '' End As PhoneTypeName "_
		  &"From vwPhoneList Where (PhoneType = '" & PhoneType_ & "' or '" & PhoneType_ & "'='A') "_
		  &"And (PhoneNumber like '" & PhoneNumber_  & "%' or '" & PhoneNumber_ & "'='') "_
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
	
	<table align="center" border="1" cellpadding="1" cellspacing="0" width="50%" bgcolor="white">
	<tr>
		<td>
		<form method="post" name="frmSearch">
		<table align="center" cellpadding="1" cellspacing="0" width="100%">		
		<tr bgcolor="#000099">
			<td height="25" colspan="6"><strong>&nbsp;<span class="style5">Search</span></strong></td>
		</tr>
		<tr>
			<td width="40%">&nbsp;&nbsp;Phone Number / Extension</td>
			<td width="2%">:</td>
			<td><input name="txtPhoneNumber" type="Input" size="20" Value='<%=PhoneNumber_%>'></td>
		</tr>
		<tr>
			<td>&nbsp;&nbsp;Filter by Phone Type</td>
			<td>:</td>
			<td><select name="PhoneTypeList">
				<option value="A">-All-</option>
				<option value="C" <%if PhoneType_ ="C" then%>Selected<%End If%> >Cell Phone</option>
				<option value="O" <%if PhoneType_ ="O" then%>Selected<%End If%> >Office Phone</option>
				<option value="H" <%if PhoneType_ ="H" then%>Selected<%End If%> >Home Phone</option>
			  </select>
			</td>
		</tr
		<tr>
			<td>&nbsp;&nbsp;Sort By&nbsp;</td>
			<td>:</td>
			<td>
				<Select name="SortList">
					<Option value="PhoneNumber" <%if SortBy_ ="PhoneNumber" then %>Selected<%End If%> >PhoneNumber</Option>
					<Option value="Location" <%if SortBy_ ="Location" then %>Selected<%End If%> >Employee Name / Location</Option>
					<Option value="PhoneType" <%if SortBy_ ="PhoneType" then %>Selected<%End If%> >Phone Type</Option>
				</Select>&nbsp;
				<Select name="OrderList">
					<Option value="Asc" <%if Order_ ="Asc" then %>Selected<%End If%> >Asc</Option>
					<Option value="Desc" <%if Order_ ="Desc" then %>Selected<%End If%> >Desc</Option>
				</Select>
			</td
		</tr>
		<tr>
			<td colspan="3" align="center">&nbsp;&nbsp;</td>
		<tr>
			<td colspan="3" align="center">&nbsp;&nbsp;
				<input type="submit" name="btnSearch" value="Search">
				&nbsp;&nbsp;<input type="button" name="btnClear" value="Reset filter" onclick="javascript:ClearFilter();">
			</td>
		</tr>
		</table></form>
		</td>
	</tr>
	</table>
	
     <table width="100%">
	<tr>
		<td width="50%"><a href="PhoneEdit.asp?State=I">Add New Phone Number</a></td>
	</tr>
     </table>
     <table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width=100%>
     <TR BGCOLOR=#330099>
         <TD width="5%" align="center"><strong><label STYLE=color:#FFFFFF>NO</label></strong></TD>
         <TD width="15%"><strong><label STYLE=color:#FFFFFF>&nbsp;Phone Number / Extension</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>&nbsp;Employee Name / Location</label></strong></TD>
         <TD width="12%"><strong><label STYLE=color:#FFFFFF>&nbsp;Post</label></strong></TD>
         <TD width="12%"><strong><label STYLE=color:#FFFFFF>&nbsp;Phone Type</label></strong></TD>
         <TD width="8%"><strong><label STYLE=color:#FFFFFF>&nbsp;Action</label></strong></TD>
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
        <TD><FONT color=#330099 size=2><A HREF="PhoneEdit.asp?ID=<%= rs("ID")%>&State=E"> <%= rs("PhoneNumber") %></A></font>   </TD>
        <TD><FONT color=#330099 size=2>&nbsp;<%= rs("Location") %></font>   </TD>
        <TD><FONT color=#330099 size=2>&nbsp;<%= rs("Post") %></font>   </TD>
        <TD><FONT color=#330099 size=2>&nbsp;<%= rs("PhoneTypeName") %></font>   </TD>
	<TD><FONT color=#330099 size=2><A HREF="PhoneDeleteConfirm.asp?ID=<%= rs("ID")%>&State=D" >Delete</A></font>   </TD>
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
					<a href="PhoneList.asp?PageIndex=<%=PageNo%>&PhoneNumber=<%=PhoneNumber_%>&PhoneType=<%=PhoneType_ %>&SortBy=<%=SortBy_%>&Order=<%=Order_%>"><%=PageNo%></a>&nbsp;
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


