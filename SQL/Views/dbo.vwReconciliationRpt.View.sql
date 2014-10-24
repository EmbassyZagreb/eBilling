/****** Object:  View [dbo].[vwReconciliationRpt]    Script Date: 08/01/2014 13:35:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE View [dbo].[vwReconciliationRpt]
As
Select ID , A.MonthP, A.YearP, A.ProgressID, ISNULL(C.ProgressDesc,'') As ErrorType, A.PhoneNumber, B.EmpName, B.Office, isNULL(A.Balance,0) As CurrentBalance
	, CONVERT(varchar(30), CreateDate,109) As CreateDate
From Reconciliation A
Left Join vwPhoneCustomerList B on (A.PhoneNumber=B.MobilePhone)
Left Join ProgressStatus C on (A.ProgressID=C.ProgressID)
GO
