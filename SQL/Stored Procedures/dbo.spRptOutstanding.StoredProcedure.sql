/****** Object:  StoredProcedure [dbo].[spRptOutstanding]    Script Date: 12/02/2014 15:00:22 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
CREATE Procedure [dbo].[spRptOutstanding] 
	@Outstanding Int=Null
	,@Operator varchar(3)=Null
	,@Agency Varchar(50)=Null
	,@Section Varchar(50)=Null
	,@EmpID Varchar(10)=Null
	,@Status Varchar(15)=Null
As
Declare	@curMonth Varchar(2)
	,@curYear Varchar(4)
	,@curPeriod Varchar(8)
	,@PaymentDueDate Varchar(2)
--	,@PaymentDueDate TinyInt
	,@strSQL varchar(3000)
/*
Set @Agency = NullIf(@Agency,'X') 
Set @Section = NullIf(@Section,'X') 
Set @EmpID = NullIf(@EmpID,'X') 
*/

Select @PaymentDueDate=Convert(Varchar(2),PaymentDueDate) From PaymentDueDate
--Select @PaymentDueDate=PaymentDueDate From PaymentDueDate
/*
Select EmpName, MonthP, YearP, Office, HomePhone, HomePhoneBillRp, WorkPhone, OfficePhonePrsBillRp, LoginID, TotalShuttleBillRp
	, TotalBillingRp, Datediff(d,YearP+MonthP+@PaymentDueDate,getdate()) As Aging, Status
	--, Case When ProgressID=6 Then 'Paid' Else 'Pending' End As Status
from vwMonthlyBilling
Where Datediff(d,YearP+MonthP+@PaymentDueDate,getdate())=@Outstanding
And ((Agency = @Agency) or @Agency is null )
And ((Office = @Section) or @Section is null )
And ((EmpID = @EmpID) or @EmpID is null )
*/
/*
Set @strSQL ='Select EmpName, MonthP, YearP, Office, HomePhone, HomePhoneBillRp, HomePhonePrsBillRp, WorkPhone, OfficePhonePrsBillRp, MobilePhone, CellPhonePrsBillRp,LoginID, TotalShuttleBillRp
	, TotalBillingRp, SendMailDate, Case When (Datediff(d,dateadd(d,'+@PaymentDueDate+',SendMailDate),getdate())<=0) or (SendMailStatusID=1) Then 0 
		Else Datediff(d,dateadd(d,'+@PaymentDueDate+',SendMailDate),getdate()) End  As Aging, Status, ProgressDesc
from vwMonthlyBilling
Where Case When (Datediff(d,dateadd(d,'+@PaymentDueDate+',SendMailDate),getdate())<=0) or (SendMailStatusID=1) Then 0 
		Else Datediff(d,dateadd(d,'+@PaymentDueDate+',SendMailDate),getdate()) End '+@Operator+convert(varchar(10),@Outstanding)+
' And ((Agency = '''+@Agency+''') or '''+@Agency +''' = ''X'' ) 
And ((Office = '''+@Section+''') or '''+@Section+''' = ''X'' ) 
And ((EmpID = '''+@EmpID+''') or '''+@EmpID+''' = ''X'' )
And ((ProgressID = '+@Status+') or '''+@Status+''' = 0 )'
print @strSQL
Exec(@strSQL)
*/

/*
Select EmpName, MonthP, YearP, Office, HomePhone, HomePhoneBillRp, HomePhonePrsBillRp, WorkPhone, OfficePhonePrsBillRp, MobilePhone, CellPhonePrsBillRp,LoginID, TotalShuttleBillRp
	, TotalBillingRp, SendMailDate, PaidDate
	, Case When (ProgressID=7) or (Len(ISNULL(SendMailDate,'')))=0 Then 0 When ProgressID in (5,6) Then Datediff(d,PaidDate,SendMailDate) 
		   Else  Datediff(d,GETDATE(),SendMailDate) End As Aging
, Status, ProgressDesc
from vwMonthlyBilling
*/

--Set @strSQL ='Select EmpName, MonthP, YearP, Office, HomePhone, HomePhoneBillRp, HomePhonePrsBillRp, WorkPhone, OfficePhonePrsBillRp, MobilePhone, CellPhonePrsBillRp,LoginID, TotalShuttleBillRp
Set @strSQL ='Select EmpName, MonthP, YearP, Office, MobilePhone, CellPhonePrsBillRp,LoginID
	, TotalBillingRp, SendMailDate, Case When (Datediff(d,dateadd(d,1,SendMailDate),getdate())<=0) or (SendMailStatusID=1) Then 0 When ProgressID in (6,7,8) Then Datediff(d, SendMailDate, ProgressIDDate)
	   Else Datediff(d, SendMailDate, getdate()) End As Aging, Status, ProgressDesc
from vwMonthlyBilling
Where Case When (Datediff(d,dateadd(d,1,SendMailDate),getdate())<=0) or (SendMailStatusID=1) Then 0 When ProgressID in (6,7,8) Then Datediff(d, SendMailDate, ProgressIDDate)
	  Else Datediff(d, SendMailDate, getdate()) End '+@Operator+convert(varchar(10),@Outstanding)+ ' And ((Agency = '''+@Agency+''') or '''+@Agency +''' = ''X'' ) 
And ((Office = '''+@Section+''') or '''+@Section+''' = ''X'' ) 
And ((EmpID = '''+@EmpID+''') or '''+@EmpID+''' = ''X'' )
And ((ProgressID = '+@Status+') or '''+@Status+''' = 0 ) 
And MobilePhone <> '''' Order by EmpName'
--print @strSQL
Exec(@strSQL)
GO
