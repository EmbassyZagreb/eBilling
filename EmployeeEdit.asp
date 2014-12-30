<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>


<script language="JavaScript">
function validate_form()
{
	valid = true;
	msg="";
	if (document.frmEmployee.txtName.value == "" )
	{
		msg = msg + "Employee Name cannot be blank !!!\n"
		valid = false;
	}

	if (document.frmEmployee.cmbPostList.value == "" )
	{
		msg = msg + "Please select Post !!!\n"
		valid = false;
	}

	if (document.frmEmployee.cmbAgencyList.value == "" )
	{
		msg = msg + "Please select Agency !!!\n"
		valid = false;
	}

	if (document.frmEmployee.cmbOfficeList.value == "" )
	{
		msg = msg + "Please select Office !!!\n"
		valid = false;
	}

	if (document.frmEmployee.cmbReportTo.value == "" )
	{
		msg = msg + "Please select Supervisor !!!\n"
		valid = false;
	}

	if (document.frmEmployee.cmbFundingAgency.value == "" )
	{
		msg = msg + "Please select Funding Agency !!!\n"
		valid = false;
	}

	if (document.frmEmployee.txtEmailAddress.value != "" )
	{
		var alnum="a-zA-Z0-9";
		exp="^[^@\\s]+@(["+alnum+"+\\-]+\\.)+["+alnum+"]["+alnum+"]["+alnum+"]?$";
		emailregexp = new RegExp(exp);

		result = document.frmEmployee.txtEmailAddress.value.match(emailregexp);
		if (result == null)
		{
			msg = msg + "Invalid data type for email address !!!\n"
			valid = false;
		}
	}

	if (document.frmEmployee.txtAlternateEmail.value != "" )
	{
		var alnum="a-zA-Z0-9";
		exp="^[^@\\s]+@(["+alnum+"+\\-]+\\.)+["+alnum+"]["+alnum+"]["+alnum+"]?$";
		emailregexp = new RegExp(exp);

		result = document.frmEmployee.txtAlternateEmail.value.match(emailregexp);
		if (result == null)
		{
			msg = msg + "Invalid data type for alternative email address !!!\n"
			valid = false;
		}
	}

	if (valid == false)
	{
		alert(msg)
	}
	return valid;
}
</script>
<% 
 dim user_ 
 dim user1_  

 user_ = request.servervariables("remote_user") 
 user1_ = user_  'user1_ = right(user_,len(user_)-4)
'user1_ = "pranataw"
'response.write user1_ & "<br>"

 EmpID_ = Request("EmpID")
 EmpType_ = Request("Type")
 State_ = request("State")

%> 
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

<form method="post" action="EmployeeSave.asp" name="frmEmployee" onsubmit="return validate_form()"> 
<%  
 dim rst 
 dim strsql
 dim SectionRS

 strsql = "Select EmpID, EmpName, PostName, Agency, Office, Type, ReportToID, ReportToName, EmailAddress, LoginID, AlternateEmail, Status, AgencyId, AgencyFundingCode, AgencyFunding From vwDirectReport where EmpID ='" & EmpID_ & "' And Type='" & EmpType_ & "'"


 'strsql = "Select EmpID, EmpName, PostName, Agency, Office, Type, ReportToID, ReportToName, EmailAddress, LoginID, AlternateEmail, AgencyId, AgencyFundingCode, AgencyFunding From vwDirectReport where EmpID ='" & EmpID_ & "' And Type='" & EmpType_ & "'"
 'response.write strsql 
  set rsData = server.createobject("adodb.recordset") 
  set rsData = BillingCon.execute(strsql)
  if not rsData.eof then 
	EmpID_ = rsData("EmpID")
	EmpName_ = rsData("EmpName")
	Post_ = rsData("PostName")
	Agency_ = rsData("Agency")
	Office_ = rsData("Office")
	Type_ = rsData("Type")
	ReportToID_ = rsData("ReportToID")
	EmailAddress_ = rsData("EmailAddress")
	LoginID_ = rsData("LoginID")
	AlternateEmail_ = rsData("AlternateEmail")
	AgencyFundingCodeEmp_ = rsData("AgencyId")
	Status_ = rsData("Status")
	'Remark_ = rsData("Remark") 
	'Remark = "No reMark"
  end if

 strsql = "Select Remark, WorkingTitle From MsEmployee where EmpID ='" & EmpID_ & "' And EmpType='" & EmpType_ & "'" 
  set rsRemark = server.createobject("adodb.recordset") 
  set rsRemark = BillingCon.execute(strsql)
  if not rsRemark.eof then 
	Remark_ = rsRemark("Remark")
	WorkingTitle_ = rsRemark("WorkingTitle") 
  end if

 'response.write ReportToID_ 
%>             
<table align="center">
<tr>
  <td>Employee Name :</td> 
  <td><input type="input" name="txtName" value='<%=EmpName_ %>' size="50" maxlength="50" /></td>
</tr>
<tr>
  <td>Working Title :</td>
  <td><input type="input" name="txtWorkingTitle" value='<%=WorkingTitle_ %>' size="50" maxlength="50" /></td>
</tr>
<tr>
  <td>Post :</td>
  <!--   <td><input type="input" name="txtPost" value='<%=Post_ %>' size="50" maxlength="50" /></td> -->


		<td>
			<Select id="cmbPostList" name="cmbPostList">
					<Option value="">-- Select --</Option>
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



</tr>
<tr>
  <td>Agency :</td>
  <!--  <td><input type="input" name="txtAgency" value='<%=Agency_ %>' size="50" maxlength="50" /></td> -->


		<td>
			<Select id="cmbAgencyList" name="cmbAgencyList">
					<Option value="">-- Select --</Option>
<%
				strsql ="select distinct Agency from vwPhoneCustomerList Where Agency<>'' order by Agency"
				set SectionRS = server.createobject("adodb.recordset")
				set SectionRS = BillingCon.execute(strsql)
				Do While not SectionRS.eof 
%>
					<Option value='<%=SectionRS("Agency")%>' <%if trim(Agency_) = trim(SectionRS("Agency")) then %>Selected<%End If%> ><%=SectionRS("Agency")%></Option>
<%					
				SectionRS.MoveNext
				Loop%>
			</select>
		</td>	


</tr>
<tr>
  <td>Office :</td>
 <!-- <td><input type="input" name="txtOffice" value='<%=Office_ %>' size="50" maxlength="50" /></td>   -->




		<td>
			<Select id="cmbOfficeList" name="cmbOfficeList">
					<Option value="">-- Select --</Option>
<%
				strsql ="select distinct Office from vwPhoneCustomerList Where Office<>'' order by Office"
				set SectionRS = server.createobject("adodb.recordset")
				set SectionRS = BillingCon.execute(strsql)
				Do While not SectionRS.eof 
%>
					<Option value='<%=SectionRS("Office")%>' <%if trim(Office_) = trim(SectionRS("Office")) then %>Selected<%End If%> ><%=SectionRS("Office")%></Option>
<%					
				SectionRS.MoveNext
				Loop%>
			</select>
		</td>	





</tr>
	<tr>
	  <td>Type :</td>
	  <td>
		  <select name="cmbType">
			<option value="AMER" <%if Type_ ="AMER" Then %>Selected<%End If%>>Supervisor</option>
			<option value="LES" <%if Type_ ="LES" Then %>Selected<%End If%>>Regular</option>
		  </select>&nbsp;&nbsp;* Only an American can be the supervisor
	  </td>
	</tr>
<tr>
  <td>Email Address :</td>
  <td><input type="input" name="txtEmailAddress" size="50" value="<%=EmailAddress_%>"/>
  </td>
</tr>
<tr>
  <td>Login ID :</td>
  <td><input type="input" name="txtLoginID" size="50" value="<%=LoginID_%>"/>
  </td>
</tr>
<tr>
  <td>Supervisor :</td>
  <td>
	<select id="cmbReportTo" name="cmbReportTo">
		<option value="">-- Select --</option>
<%
		Dim UserRS
		'strsql = "Select EmpID, ISNULL(EmpName,'')+' - '+ISNULL(Office,'') As EmpName from vwPhoneCustomerList Where LEN(ISNULL(EmpName,''))<>'' AND EmpType = 'AMER' Order by EmpName"
		strsql = "Select EmpID, ISNULL(EmpName,'')+' - '+ISNULL(Office,'') As EmpName from vwDirectReport Where LEN(ISNULL(EmpName,''))<>'' AND Type = 'AMER' Order by EmpName"
		response.write strsql & "<br>"
		set UserRS = server.createobject("adodb.recordset")
		set UserRS =BillingCon.execute(strsql)				        
		do while not UserRS.eof
			EmpIDX_ = UserRS("EmpID")
			Ename_ = UserRS("EmpName")
%>
	        <OPTION value='<%=EmpIDX_%>' <%if (trim(EmpIDX_) = trim(ReportToID_ )) then%>selected<%end if%>><%= EName_  %>
<%
                 UserRS.movenext
	        loop
%>  
	</select>	
  </td>
</tr>
<tr>
  <td>Alternate Email  :</td>
  <td><input id="txtAlternateEmail" name="txtAlternateEmail" value='<%=AlternateEmail_ %>' size="50" </input>
  </td>
</tr>

<tr>
  <td>Funding Agency :</td>
  <td>
	<select id="cmbFundingAgency" name="cmbFundingAgency">
		<option value="">-- Select --</option>
<%
		Dim AgencyRS
		strsql = "Select AgencyId, AgencyFundingCode, AgencyDesc from AgencyFunding Where Disabled='N' Order by AgencyDesc"
		'response.write strsql & "<br>"
		set AgencyRS = server.createobject("adodb.recordset")
		set AgencyRS =BillingCon.execute(strsql)				        
		do while not AgencyRS.eof
			AgencyFundingCode_ = AgencyRS("AgencyId")
			AgencyFunding_ = AgencyRS("AgencyDesc")
%>
	        <OPTION value='<%=AgencyFundingCode_%>' <%if (trim(AgencyFundingCodeEmp_) = trim(AgencyFundingCode_ )) then%>selected<%end if%>><%= AgencyFunding_   %>
<%
                 AgencyRS.movenext
	        loop
%>  
	</select>	
  </td>
</tr>
<tr>
  <td>Remark [Phone Number] :</td>
  <td><input type="input" name="txtRemark" size="50" value="<%=Remark_%>"/>
  </td>
</tr>
	<tr>
	  <td>Status :</td>
	  <td>
		  <select name="cmbStatus">
			<option value="C" <%if Status_ ="C" Then %>Selected<%End If%>>Current</option>
			<option value="D" <%if Status_ ="D" Then %>Selected<%End If%>>Departed</option>
		  </select>
	  </td>
	</tr>
<tr>
  <td colspan="2"><br></td>
</tr>
<tr>
  <td></td>
  <td><input type="submit" name="btnSubmit" value="Update">
		<%if State_= "E" then %>
		      <input type="hidden" name="txtEmpID" value='<%=EmpID_ %>'>
		<%End If%>
      <input type="hidden" name="txtEmpType" value=<%=Type_ %>>
<input type="hidden" name="State" value=<%=State_ %> >
      &nbsp;<input type="button" value="Cancel" name="btnCancel" onClick="javascript:location.href='EmployeeList.asp'">
 </td>
</tr>
</table>
<p></p>
<table>  
<tr><td>Note:</td><td>Users with multiple cell phones must have the number entered in the Remark [Phone Number] field!</td></tr>
<tr><td>&nbsp;</td><td>Users outside OpenNet system must use Alternate Email field. Leave Email Address blank.</td></tr>
</table>




<table border="0" bordercolor="#FFFFFF" cellpadding="2" cellspacing="0" width="65%"  class="FontText">
	<tr>
		<td><u><b>Historical assignment of Funding Agency:<b></u></td>
	</tr>
	<tr>
		<td class="Hint" align="left">*To alter historical data 'Generate Monthly Billing' procedure must be executed. Procedure sets bill to 'Pending' status.</td>
	</tr>
</table>	









<%

strsql = "Select Distinct YearP+MonthP, YearP, MonthP, AgencyFundingCode, AgencyFundingDesc, FiscalStripVAT, FiscalStripNonVAT From vwMonthlyBilling Where EmpID = '" & EmpID_ & "' Order by (YearP+MonthP) Desc"
set DataRS = server.createobject("adodb.recordset")
DataRS.CursorLocation = 3
DataRS.Open strsql,BillingCon
'set DataRS=BillingCon.execute(strsql)

dim intPageSize,PageIndex,TotalPages 
dim RecordCount,RecordNumber,Count 
intpageSize=10 
PageIndex=request("PageIndex")

if PageIndex ="" then PageIndex=1 

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
'response.write strsql

dim intPrev,intNext 	
intPrev=PageIndex - 1 
intNext=PageIndex +1 


if not DataRS.eof Then

%>
<!-- <div align="right"><input type="submit" name="btnApproval" value="Approve" /></div> -->
<!-- <div align="right"><input type="button" name="btnExport" value="Export to Excel" onClick="javascript:document.location.href('ARBillingReportAllPrint.asp?sMonthP=<%=sMonthP%>&sYearP=<%=sYearP%>&eMonthP=<%=eMonthP%>&eYearP=<%=eYearP%>&Agency=<%=Agency_%>&Section=<%=Section_%>&EmpID=<%=EmpID_%>&Status=<%=Status%>');"/></div> -->
<table border="1" bordercolor="#EEEEEE" cellpadding="2" cellspacing="0" width="65%"  class="FontText">
    <TR BGCOLOR="#330099" align="center">
         <TD width="7%"><strong><label STYLE=color:#FFFFFF>Billing<br>Period</label></strong></TD>
         <TD width="7%"><strong><label STYLE=color:#FFFFFF>Agency<br>Code</label></strong></TD>
         <TD width="20%"><strong><label STYLE=color:#FFFFFF>Agency Name</label></strong></TD>
         <TD width="33%"><strong><label STYLE=color:#FFFFFF>Fiscal Strip VAT</label></strong></TD>
         <TD width="33%"><strong><label STYLE=color:#FFFFFF>Fiscal Strip Non VAT</label></strong></TD>


    </TR>
<% 
   dim no_  
   no_ = 1 + ((PageIndex*intPageSize)-intPageSize)
   Count=1 
   do while not DataRS.eof   and Count<=intPageSize
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
	   <TR bgcolor="<%=bg%>">
	    <TD align="right">&nbsp;<%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font>&nbsp;</TD>
	    <TD>&nbsp;<%=DataRS("AgencyFundingCode") %></TD>
	    <TD>&nbsp;<%=DataRS("AgencyFundingDesc") %> </font></TD>
		<TD>&nbsp;<%= DataRS("FiscalStripVAT") %></font></TD>
		<TD>&nbsp;<%= DataRS("FiscalStripNonVAT") %></font></TD>
	   </TR>

<%   
		Count=Count +1
	   DataRS.movenext
	   no_ = no_ + 1 
   loop 
	PageNo=1
%>
</table>
<table width="65%">
	<tr>
		<td align="right">
<%
		Do while PageNo<=TotalPages 
			if trim(pageNo) = trim(PageIndex) Then
%>		
				<label class="ActivePage"><%=PageNo%></label>&nbsp;
			<%Else%>
				<a href="EmployeeEdit.asp?PageIndex=<%=PageNo%>&ID=<%=ID_%>&State=E"><%=PageNo%></a>&nbsp;
<%	
			End If						
			PageNo=PageNo+1
		Loop
%>
		</td>
	</tr>
</table>
<%
else 
%>
	<table cellspadding="1" cellspacing="0" width="100%">  
	<tr>
        	<td><br></TD>
	</tr>
	<tr>
		<td align="center">not data.</td>
	</tr>
	<tr>
        	<td><br></TD>
	</tr>
	<tr>
		<td align="center"><a href="Default.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
	</tr>	
	</table>
<% end if %>









</form>
</BODY>
</html>