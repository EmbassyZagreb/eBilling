/****** Object:  StoredProcedure [dbo].[spGetPaymentReceipt]    Script Date: 08/01/2014 13:31:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[spGetPaymentReceipt]
	 @StartPeriod Varchar(6)
	, @EndPeriod Varchar(6)
	--,@EmpID Varchar(10)   *Marin-out
	,@EmpName Varchar(50) --*Marin-in
	,@Outstanding Int=0
	,@Section Varchar(50)=Null
	--,@Status Varchar(1)=Null
	,@Status TinyInt=0
As
Set @Section = NullIf(@Section,'X')

Set @StartPeriod=NullIf(@StartPeriod,'')
Set @EndPeriod=NullIf(@EndPeriod,'')
--Set @EmpID = NullIf(@EmpID,'X')    *Marin-out
Set @EmpName = NullIf(@EmpName,'') --*Marin-in
Set @Section=NullIf(@Section,'')
Set @Status=NullIf(@Status,0)


Declare @PaymentDueDate TinyInt
		--@PaymentDueDate Varchar(2)
	   
--Select @PaymentDueDate=Right('00'+Convert(Varchar(2),PaymentDueDate),2) From PaymentDueDate
Select @PaymentDueDate=PaymentDueDate From PaymentDueDate

Select  A.EmpID, A.EmpName, A.MonthP, A.YearP, A.Office, isNull(A.HomePhoneBillRp,0) As HomePhoneBillRp, isNull(A.HomePhonePrsBillRp,0) As HomePhonePrsBillRp
	, isNull(A.OfficePhonePrsBillRp,0) As OfficePhonePrsBillRp, A.MobilePhone, isNull(A.CellPhonePrsBillRp,0) As CellPhonePrsBillRp, A.LoginID, isNull(A.TotalShuttleBillRp,0) As TotalShuttleBillRp
	, isNull(A.TotalBillingAmountPrsRp,0) As TotalBillingRp
	, isNull(A.PaidAmountDlr,0) As PaidAmountDlr, isNull(A.PaidAmountRp,0) As PaidAmountRp
	--, Datediff(d,YearP+MonthP+@PaymentDueDate,getdate()) As Aging
	,Case When (Datediff(d,dateadd(d,@PaymentDueDate,SendMailDate),getdate())<=0) or (SendMailStatusID=1) Then 0 
		Else Datediff(d,dateadd(d,@PaymentDueDate,SendMailDate),getdate()) End As Aging
	, A.ProgressID
	--, Case When A.ProgressID <=4 Then 'Pending' Else B.ProgressDesc End As Status
	, isNull(B.ProgressDesc,'') As Status
	, A.PaidDate    -- Marin
	, A.AlternateEmailFlag
From vwMonthlyBilling A
Left Join ProgressStatus B on (A.ProgressID=B.ProgressID)
Where (A.YearP+A.MonthP>=@StartPeriod or @StartPeriod is Null)
And (A.YearP+A.MonthP<=@EndPeriod or @EndPeriod is Null)
--And ((A.EmpID = @EmpID) or @EmpID is null )    *Marin-out
And ((A.EmpName like @EmpName) or @EmpName is null )  --*Marin-in
And (Datediff(d,dateadd(d,@PaymentDueDate,SendMailDate),getdate()) =@Outstanding or @Outstanding=0)
And ((Office = @Section) or @Section is null )
--And ((Case When A.ProgressID<6 Then 'P' Else 'F' End=@Status) or @Status is Null)
And ((A.ProgressID=@Status) or @Status is Null)
Order by A.EmpName
GO
