
<div id="extra">
			<ul id="menu">
			<%  If Not rsYourStaffEmpty Then 	%>
				<li>
					<a href="#">Waiting For Your Approval</a>
					<ul>
						<%
						If rsApprovalEmpty Then
						%>
						<li class="current"><a href="#">All Approved</a></li>
						<%
						Else
							For i = 0 To UBound(arrNumberList,2)
						%>
							<li <%if arrNumberList(0,i)=MobilePhone_ and arrNumberList(1,i)=MonthP and arrNumberList(2,i)=YearP and Nav_=1 then %>class="current"<%End If%>><a href="1BillingApproval.asp?CellPhone=<%=arrNumberList(0,i)%>&MonthP=<%=arrNumberList(1,i)%>&YearP=<%=arrNumberList(2,i)%>&LoginID=<%=arrNumberList(3,i)%>&Nav=1"><%=arrNumberList(4,i)%></a></li>
						<%
							Next
						End If
						%>
					</ul>
				</li>
			<%	End If	%>
				<li>
					<a href="#">Your Staff :</a>
				</li>
				<li>
			<%  If rsYourStaffEmpty Then	%>
						<ul><li class="current"><a href="#">No Subordinates</a></li></ul>
			<%	Else
				NavName = ""
					For i = 0 To UBound(arrYourStaff,2)
						If arrYourStaff(1,i) <> NavName Then
							NavName = arrYourStaff(1,i)
							If i <> 0 Then	%>
						</ul>
				<%			End If	%>
							<a href="#"><%=arrYourStaff(1,i)%></a>
						<ul>
				<%		End If
						If arrYourStaff(0,i) <> "" Then %>
							<li <%if arrYourStaff(0,i)=MobilePhone_ and Nav_=2 then %>class="current"<%End If%>><a href="1BillingApproval.asp?CellPhone=<%=arrYourStaff(0,i)%>&EmpID=<%=arrYourStaff(2,i)%>&Nav=2"><%=arrYourStaff(0,i)%></a></li>
				<%		Else	%>
							<li><a href="#">No Number Assigned</a></li>
				<%		End If
					Next %>
					</ul>
			<%	End If	%>
				</li>
			</ul>
</div>

</div>

<div id="header"><h1><%=SiteHeader%></h1>
<table widht=100% bolder = 0 cellpadding=0 cellspacing=0>
	<tr>

		<td colspan=2  style="background: url(images/top-navigation-slice.jpg) repeat left top;">
			<div id="navigation-left">
			<ul>
            	<li><a href="1MonthlyBilling.asp">Home</a></li>
                <li id="active"><a href="1BillingApproval.asp">Approve</a></li>
				<li><a href="Default.asp">Cashier</a></li>
                <li><a href="Default.asp">Manage</a></li>
				<li><a href="Default.asp">Alerts</a></li>
				<li><a href="Default.asp">Reports</a></li>
			</ul>
 		</div>
		</td>

		<td  width=100% style="background: url(images/top-navigation-slice.jpg) repeat left top;">
		&nbsp;
		</td>
		<td widht=130px>
		<div id="navigation-right">
			<ul>
   	            <li><a href="Default.asp">Admin</a></li>
				<li><a href="mailto:zagrebisc@state.gov">Help</a></li>
			</ul>
		</div>
		</td>
	</tr>
</table>
</div>
