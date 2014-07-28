/****** Object:  View [dbo].[vwProgressSummary]    Script Date: 07/28/2014 12:50:24 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE View [dbo].[vwProgressSummary]
As
Select A.MonthP, A.YearP, ProgressID, ProgressDesc, COUNT(EmpID) As TotalRecord, SUM(TotalBillingRp) As TotalBill
From vwMonthlyBilling A
Inner Join CellPhoneHd B on (A.MonthP=B.MonthP And A.YearP=B.YearP And A.MobilePhone=B.PHONENUMBER)
Where A.TotalBillingRp>0
Group By A.MonthP, A.YearP, A.ProgressID, A.ProgressDesc
GO
