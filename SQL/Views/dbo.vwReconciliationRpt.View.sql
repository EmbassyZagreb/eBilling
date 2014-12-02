/****** Object:  View [dbo].[vwReconciliationRpt]    Script Date: 12/02/2014 15:00:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
-- Update vwReconciliationRpt
CREATE VIEW [dbo].[vwReconciliationRpt]
AS
SELECT     A.ID, A.MonthP, A.YearP, A.ProgressID, ISNULL(C.ProgressDesc, '') AS ErrorType, A.PhoneNumber, B.EmpName, B.Office, ISNULL(B.AgencyFunding, '') 
                      AS AgencyFundingDesc, ISNULL(A.Balance, 0) AS CurrentBalance, ISNULL(A.CreateBy, '') AS CreateBy, CONVERT(varchar(30), A.CreateDate, 100) 
                      AS CreateDate
FROM         dbo.Reconciliation AS A LEFT OUTER JOIN
                      dbo.vwPhoneCustomerList AS B ON A.PhoneNumber = B.MobilePhone LEFT OUTER JOIN
                      dbo.ProgressStatus AS C ON A.ProgressID = C.ProgressID
GO
