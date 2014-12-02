/****** Object:  StoredProcedure [dbo].[spGenerateProgressLog]    Script Date: 12/02/2014 15:00:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
CREATE Procedure [dbo].[spGenerateProgressLog]
	@Month Varchar(2)=''
	,@Year Varchar(4)=''
As
Delete ProgressLog Where MonthP=@Month And YearP=@Year

Insert ProgressLog(MonthP, YearP, ProgressID, [Description], TotalRecord, TotalBill)
Select MonthP, YearP, ProgressID, ProgressDesc, TotalRecord, TotalBill
From vwProgressSummary
Where MonthP=@Month And YearP=@Year

Insert ProgressLog(MonthP, YearP, ProgressID, [Description], TotalRecord, TotalBill)
Select MonthP ,YearP, ISNULL(ProgressID,0), ErrorType, COUNT(ID), SUM(CurrentBalance) 
From vwReconciliationRpt
Where MonthP=@Month And YearP=@Year
Group By MonthP ,YearP, ProgressID, ErrorType
GO
