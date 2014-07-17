<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>

<%
'Response.ContentType ="application/vnd.ms-excel" 
'Response.Buffer  =  True 
'Response.Clear() 
%> 
<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!--#include file="connect.inc" -->

<script language="JavaScript">
function ClearFilter()
{
	document.forms['frmSearch'].elements['PostList'].value ='';
	document.forms['frmSearch'].elements['SortList'].value ='PhoneNumber';
}
</script>


<%

Post_ = Request.Form("PostList")
if Post_ ="" then
	Post_ = Request("Post")	
end if

Extension_ = Request.Form("txtExtension")
'response.write "Extension" & Extension_
if (Extension_ ="") then
	if Request("Extension")<>"" then
		Extension_ = Request("Extension")	
	Else
		Extension_ = ""
	end if
end if

MonthP1_ = Request.Form("MonthList1")
'response.write "MonthP :" & MonthP_
if MonthP1_ ="" then
	if Request("Month1")<>"" then 
		MonthP1_ = Request("Month1")
	else
		MonthP1_ = session("Month")
	end if
end if

MonthP2_ = Request.Form("MonthList2")
'response.write "MonthP :" & MonthP_
if MonthP2_ ="" then
	if Request("Month2")<>"" then 
		MonthP2_ = Request("Month2")
	else
		MonthP2_ = session("Month")
	end if
end if

YearP1_ = Request.Form("YearList1")
'response.write "YearP :" & YearP1_
if YearP1_ ="" then
	if Request("Year")<> "" then
		YearP1_ = Request("Year")
	else
		YearP1_ = session("Year")
	end if
end if

YearP2_ = Request.Form("YearList2")
'response.write "YearP :" & YearP2_
if YearP2_ ="" then
	if Request("Year2")<> "" then
		YearP2_ = Request("Year2")
	else
		YearP2_ = session("Year")
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
%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">COMPARISON MONTHLY BILLING BY EXT. REPORT</TD>
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
	strsql = "spMonthlyBillingByExtReport '" & Post_ & "','" & Extension_ & "','" & MonthP1_ & "','" & YearP1_ & "','" & MonthP2_ & "','" & YearP2_ & "','" & SortBy_ & "','" & Order_ & "'"
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
	<table cellspadding="1" cellspacing="0" width="70%" border="1" align="center">
	<tr align="Center">
		<td colspan="2" align="center">
			<form method="post" name="frmSearch" Action="MonthlyBillingByExtReport.asp">
			<table  width="100%">
			<tr bgcolor="#000099">
				<td height="25" colspan="7"><strong>&nbsp;<span class="style5">Search &amp; Sort By </span></strong></td>
			</tr>
			<tr>
				<td width="25%">&nbsp;Post&nbsp;</td>
				<td>:</td>
		        	<td>
				<Select name="PostList">
					<Option value="">-- All --</Option>
					<Option value="ZAGREB" <%if Post_ ="ZAGREB" then %>Selected<%End If%> >ZAGREB</Option>
					<Option value="PODGORICA" <%if Post_ ="PODGORICA" then %>Selected<%End If%> >PODGORICA</Option>
				</Select>
				</td>  	
				<td width="15%">&nbsp;Phone Ext.&nbsp;</td>
				<td>:</td>
		        	<td>
				<%
					strsql = "Select PhoneNumber From MsOfficePhoneNumber order by phonenumber"
					set PhoneRS = server.createobject("adodb.recordset")
					set PhoneRS = BillingCon.execute(strsql)
				%>
					<select name="ExtensionList"> 
						<Option value="A">-- All --</option>
				<%	do while not PhoneRS.eof 
					if trim(Extension_) = trim(PhoneRS("PhoneNumber")) then%>
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
				<td>&nbsp;Comparison Period 1 &nbsp;</td>
				<td>:</td>
				<td>
					<Select name="MonthList1">
						<Option value="01" <%if MonthP1_ ="01" then %>Selected<%End If%> >January</Option>
						<Option value="02" <%if MonthP1_ ="02" then %>Selected<%End If%> >February</Option>
						<Option value="03" <%if MonthP1_ ="03" then %>Selected<%End If%> >March</Option>
						<Option value="04" <%if MonthP1_ ="04" then %>Selected<%End If%> >April</Option>
						<Option value="05" <%if MonthP1_ ="05" then %>Selected<%End If%> >May</Option>
						<Option value="06" <%if MonthP1_ ="06" then %>Selected<%End If%> >June</Option>
						<Option value="07" <%if MonthP1_ ="07" then %>Selected<%End If%> >July</Option>
						<Option value="08" <%if MonthP1_ ="08" then %>Selected<%End If%> >August</Option>
						<Option value="09" <%if MonthP1_ ="09" then %>Selected<%End If%> >Sepetember</Option>
						<Option value="10" <%if MonthP1_ ="10" then %>Selected<%End If%> >October</Option>
						<Option value="11" <%if MonthP1_ ="11" then %>Selected<%End If%> >November</Option>
						<Option value="12" <%if MonthP1_ ="12" then %>Selected<%End If%> >December</Option>
					</Select>&nbsp;
<%
					Year_ = Year(Date()) - 1
'					response.write YearP1_
%>

					<Select name="YearList1">
<% 				Do While Year_ <= Year(Date()) %>
					<Option value='<%=Year_%>' <%if trim(Year_) = trim(YearP1_) then %>Selected<%End If%> ><%=Year_%></Option>		
<% 
				Year_ = Year_ + 1
				Loop %>	
					</Select>										
				</td>
				<td width="15%">&nbsp;Sort By&nbsp;</td>
				<td>:</td>
				<td>
					<Select name="SortList">
						<Option value="PhoneNumber" <%if SortBy_ ="PhoneNumber" then %>Selected<%End If%> >Phone Nember</Option>
						<Option value="Location" <%if SortBy_ ="Location" then %>Selected<%End If%> >Employee Name</Option>
						<Option value="TotalCost1" <%if SortBy_ ="TotalCost1" then %>Selected<%End If%> >Total Payment 1</Option>
					</Select>&nbsp;
					<Select name="OrderList">
						<Option value="Asc" <%if Order_ ="Asc" then %>Selected<%End If%> >Asc</Option>
						<Option value="Desc" <%if Order_ ="Desc" then %>Selected<%End If%> >Desc</Option>
					</Select>
				</td>
			</tr>
			<tr>
				<td>&nbsp;Comparison Period 2 &nbsp;</td>
				<td>:</td>
				<td>
					<Select name="MonthList2">
						<Option value="01" <%if MonthP2_ ="01" then %>Selected<%End If%> >January</Option>
						<Option value="02" <%if MonthP2_ ="02" then %>Selected<%End If%> >February</Option>
						<Option value="03" <%if MonthP2_ ="03" then %>Selected<%End If%> >March</Option>
						<Option value="04" <%if MonthP2_ ="04" then %>Selected<%End If%> >April</Option>
						<Option value="05" <%if MonthP2_ ="05" then %>Selected<%End If%> >May</Option>
						<Option value="06" <%if MonthP2_ ="06" then %>Selected<%End If%> >June</Option>
						<Option value="07" <%if MonthP2_ ="07" then %>Selected<%End If%> >July</Option>
						<Option value="08" <%if MonthP2_ ="08" then %>Selected<%End If%> >August</Option>
						<Option value="09" <%if MonthP2_ ="09" then %>Selected<%End If%> >Sepetember</Option>
						<Option value="10" <%if MonthP2_ ="10" then %>Selected<%End If%> >October</Option>
						<Option value="11" <%if MonthP2_ ="11" then %>Selected<%End If%> >November</Option>
						<Option value="12" <%if MonthP2_ ="12" then %>Selected<%End If%> >December</Option>
					</Select>&nbsp;
<%
					Year_ = Year(Date()) - 1
'					response.write YearP2_
%>

					<Select name="YearList2">
<% 				Do While Year_ <= Year(Date()) %>
					<Option value='<%=Year_%>' <%if trim(Year_) = trim(YearP2_) then %>Selected<%End If%> ><%=Year_%></Option>		
<% 
				Year_ = Year_ + 1
				Loop %>	
					</Select>										
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
	<form method="post" name="frmHomephoneList" action="MonthlyBillingByExtPrint.asp?Post=<%=Post_%>&Extension=<%=Extension_%>&Month1=<%=MonthP1_%>&Year1=<%=YearP1_%>&Month2=<%=MonthP2_%>&Year2=<%=YearP2_%>&SortBy=<%=SortBy_%>&Order=<%=Order_%>">
	<table cellpadding="1" cellspacing="0" width="100%">
	<tr align="right">
		<td align="left">&nbsp;<input type="submit" value="Export to Excel" /></td>
	</tr>
	</table>
	<table align="center" cellpadding="1" cellspacing="0" width="100%" border="1" bordercolor="black"> 
	<TR BGCOLOR="#000099" align="center" cellpadding="0" cellspacing="0" >
		<TD width="4%"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
	       	<TD width="14%"><strong><label STYLE=color:#FFFFFF>Phone Number</label></strong></TD>
		<TD><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
	       	<TD width="10%"><strong><label STYLE=color:#FFFFFF>Cost Comparison 1 (Kn)</label></strong></TD>
		<TD width="10%"><strong><label STYLE=color:#FFFFFF>Cost Comparison 2 (Kn)</label></strong></TD>
		<TD width="10%"><strong><label STYLE=color:#FFFFFF>Diff (%)</label></strong></TD>
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
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("PhoneNumber")%></font></td> 
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("Location")%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("TotalCost1"),-1)%>&nbsp;</font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("TotalCost2"),-1)%>&nbsp;</font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=DataRS("Percentage")%>&nbsp;</font></td> 
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
					<a href="MonthlyBillingByExtReport.asp?PageIndex=<%=PageNo%>&Post=<%=Post_%>&Extension=<%=Extension_%>&Month1=<%=MonthP1_%>&Year1=<%=YearP1_%>&Month2=<%=MonthP2_%>&Year2=<%=YearP2_%>&SortBy=<%=SortBy_%>&Order=<%=Order_%>"><%=PageNo%></a>&nbsp;
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


