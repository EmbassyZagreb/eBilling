/****** Object:  StoredProcedure [dbo].[spUpdateTotPersonalCall]    Script Date: 08/01/2014 13:31:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[spUpdateTotPersonalCall]
	 @Type Varchar(1)=Null
	,@EmpID Varchar(10)=Null
	,@ExtNo Varchar(15)=Null
	,@Month Varchar(2)=Null
	,@Year Varchar(4)=Null
As
Declare @TotPersonalCall Numeric(18,2)	

--Office Phone
If @Type='1'
Begin

	Select @TotPersonalCall=isNull(Sum(isNull(Cost,0)),0)
	From BillingDt 
	Where Extension=@ExtNo And MonthP=@Month and YearP=@Year And isPersonal='Y'
	
	Update MonthlyBilling Set OfficePhonePrsBillRp=@TotPersonalCall, OfficePhonePrsBillDlr=isNull(Round(@TotPersonalCall/ExchangeRate,0),0)
		, TotalBillingRp= HomePhoneBillRp+@TotPersonalCall+TotalShuttleBillRp+CellPhonePrsBillRp
		, TotalBillingDlr=HomePhoneBillDlr+Round(isNull(isNull(@TotPersonalCall,0)/ExchangeRate,0),2) +TotalShuttleBillDlr+CellPhonePrsBillDlr
	Where EmpID=@EmpID And MonthP=@Month and YearP=@Year

End
--Home Phone
Else If @Type='2'

Begin

	Select @TotPersonalCall=isNull(Sum(isNull(Cost,0)),0)
	From HomePhoneDt 
	Where PhoneNumber=@ExtNo And MonthP=@Month and YearP=@Year And isPersonal='Y'

	Update MonthlyBilling Set HomePhonePrsBillRp=@TotPersonalCall, HomePhonePrsBillDlr=isNull(Round(@TotPersonalCall/ExchangeRate,0),0)
--		, TotalBillingRp= OfficePhonePrsBillRp+@TotPersonalCall+TotalShuttleBillRp+CellPhonePrsBillRp
--		, TotalBillingDlr= OfficePhonePrsBillDlr+Round(isNull(isNull(@TotPersonalCall,0)/ExchangeRate,0),2) +TotalShuttleBillDlr+CellPhonePrsBillDlr
	Where EmpID=@EmpID And MonthP=@Month and YearP=@Year

End
--Cell Phone
Else If @Type='3'
Begin
	Select @TotPersonalCall=isNull(Sum(isNull(Cost,0)),0)
	From CellPhoneDt 
	Where PhoneNumber=@ExtNo And MonthP=@Month and YearP=@Year And isPersonal='Y'

	Update MonthlyBilling Set CellPhonePrsBillRp=@TotPersonalCall, CellPhonePrsBillDlr=isNull(Round(@TotPersonalCall/ExchangeRate,2),0)
--		, TotalBillingRp= OfficePhonePrsBillRp+HomePhonePrsBillRp+@TotPersonalCall+TotalShuttleBillRp
--		, TotalBillingDlr=OfficePhonePrsBillDlr+HomePhonePrsBillDlr+Round(isNull(isNull(@TotPersonalCall,0)/ExchangeRate,0),2) +TotalShuttleBillDlr
	Where PhoneNumber=@ExtNo And EmpID=@EmpID And MonthP=@Month and YearP=@Year
End
GO
