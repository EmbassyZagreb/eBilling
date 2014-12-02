
<div id="extra">
			<ul id="menu">
				<li>
					<a href="#">Cell Phone Numbers</a>
					<ul>
				

						<%
						If rsempty Then	
						%>
						<li class="current"><a href="#">No number</a></li>	
						<%
						Else
							For i = 0 To UBound(arrNumberList,2)	
						%>	
							<li <%if arrNumberList(0,i)=MobilePhone_ then %>class="current"<%End If%>><a href="1MonthlyBilling.asp?CellPhone=<%=arrNumberList(0,i)%>"><%=arrNumberList(0,i)%></a></li>
						<%
							Next
						End If	
						%>				
					</ul>
				</li>
				<li>
					<a href="#">Residential Phones</a>
					<ul>
						<li><a href="#">Not available yet</a></li>
					</ul>
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
            	<li id="active"><a href="1MonthlyBilling.asp">Home</a></li>
                <li><a href="1BillingApproval.asp">Approve</a></li>
				<li><a href="#">Cashier</a></li>
                <li><a href="#">Manage</a></li>
				<li><a href="#">Alerts</a></li> 	            
				<li><a href="1MonthlyBillingAll.asp">Reports</a></li>
			</ul>
 		</div> 
		</td>
		
		<td  width=100% style="background: url(images/top-navigation-slice.jpg) repeat left top;">
		&nbsp;
		</td>
		<td widht=130px>
		<div id="navigation-right">
			<ul>
   	            <li><a href="#">Admin</a></li>
				<li><a href="#">Help</a></li>
			</ul>
		</div>  
		</td>		
	</tr>
</table> 
</div>