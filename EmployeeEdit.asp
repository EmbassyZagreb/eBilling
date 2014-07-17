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
 user1_ = right(user_,len(user_)-4)
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

 strsql = "Select EmpID, EmpName, PostName, Agency, Office, Type, ReportToID, ReportToName, EmailAddress, LoginID, AlternateEmail, AgencyId, AgencyFundingCode, AgencyFunding From vwDirectReport where EmpID ='" & EmpID_ & "' And Type='" & EmpType_ & "'"


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
			<option value="AMER" <%if Type_ ="AMER" Then %>Selected<%End If%>>AMER</option>
			<option value="LES" <%if Type_ ="LES" Then %>Selected<%End If%>>LES</option>
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
		strsql = "Select EmpID, ISNULL(EmpName,'')+' - '+ISNULL(Office,'') As EmpName from vwPhoneCustomerList Where LEN(ISNULL(EmpName,''))<>'' AND EmpType = 'AMER' Order by EmpName"
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
</form>
</BODY>
</html>