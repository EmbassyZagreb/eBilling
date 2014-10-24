/****** Object:  StoredProcedure [dbo].[spRptAging]    Script Date: 08/01/2014 13:31:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[spRptAging]
	 @sMonth Varchar(2)=Null
	,@sYear Varchar(4)=Null
	,@eMonth Varchar(2)=Null
	,@eYear Varchar(4)=Null
	,@Agency Varchar(50)=Null
	,@Section Varchar(50)=Null
	,@EmpID Varchar(10)=Null
	--,@Status Varchar(15)=Null
	,@Status TinyInt=0
As
Declare	@curMonth Varchar(2)
	,@curYear Varchar(4)
	,@curPeriod Varchar(8)
	--,@PaymentDueDate Varchar(2)
	,@PaymentDueDate TinyInt

Set @Agency = NullIf(@Agency,'X') 
Set @Section = NullIf(@Section,'X') 
Set @EmpID = NullIf(@EmpID,'X') 
Set @Status = NullIf(@Status,0) 

--Select @PaymentDueDate=Right('00'+Convert(Varchar(2),PaymentDueDate),2) From PaymentDueDate
Select @PaymentDueDate=PaymentDueDate From PaymentDueDate

Set @curMonth=Right('00'+Convert(Varchar(2),month(Getdate())),2)
Set @curYear=Right('0000'+Convert(Varchar(4),year(Getdate())),4)

--Set @curPeriod=@curYear+@curMonth+@PaymentDueDate

Select EmpName, MonthP, YearP, Office, HomePhone, HomePhoneBillRp, HomePhonePrsBillRp, WorkPhone, OfficePhonePrsBillRp, MobilePhone, CellPhonePrsBillRp, LoginID, TotalShuttleBillRp
	, TotalBillingRp, Case When (Datediff(d,dateadd(d,@PaymentDueDate,SendMailDate),getdate())<=0) or (SendMailStatusID=1) Then 0 
		Else Datediff(d,dateadd(d,@PaymentDueDate,SendMailDate),getdate()) End As Aging, Status, ProgressDesc
	--, Case When ProgressID=6 Then 'Paid' Else 'Pending' End As Status
from vwMonthlyBilling
Where YearP+MonthP>=@sYear+@sMonth And YearP+MonthP<=@eYear+@eMonth
And ((Agency = @Agency) or @Agency is null )
And ((Office = @Section) or @Section is null )
And ((EmpID = @EmpID) or @EmpID is null )
And ((ProgressID = @Status) or @Status is null )
GO
