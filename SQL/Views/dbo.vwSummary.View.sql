/****** Object:  View [dbo].[vwSummary]    Script Date: 12/02/2014 15:00:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
-- Update vwSummary
CREATE VIEW [dbo].[vwSummary]
AS
SELECT     dbo.ProgressLog.MonthP, dbo.ProgressLog.YearP, dbo.ProgressLog.ProgressID, dbo.ProgressStatus.ProgressDesc AS Description, 
                      dbo.ProgressLog.TotalRecord AS TotalRecordGenerated, dbo.ProgressLog.TotalBill AS TotalBillGenerated
FROM         dbo.ProgressLog INNER JOIN
                      dbo.ProgressStatus ON dbo.ProgressLog.ProgressID = dbo.ProgressStatus.ProgressID
UNION
SELECT     A.MonthP AS MonthPCurrent, A.YearP AS YearPCurrent, A.ProgressId AS ProgressIdCurrent, A.ProgressDesc AS ProgressDescCurrent, COUNT(A.EmpID) AS TotalRecordCurrent, SUM(A.TotalBillingRp) AS TotalBillCurrent
FROM         dbo.vwMonthlyBilling AS A INNER JOIN
                      dbo.CellPhoneHd AS B ON A.MonthP = B.MonthP AND A.YearP = B.YearP AND A.MobilePhone = B.PHONENUMBER
WHERE     (A.TotalBillingRp > 0)
GROUP BY A.MonthP, A.YearP, A.ProgressId, A.ProgressDesc
GO
