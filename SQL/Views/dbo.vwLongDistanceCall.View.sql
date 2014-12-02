/****** Object:  View [dbo].[vwLongDistanceCall]    Script Date: 12/02/2014 15:00:28 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
-- Update vwLongDistanceCall
CREATE VIEW [dbo].[vwLongDistanceCall]
AS
SELECT     B.MonthP, B.YearP, 'Cell Phone' AS PhoneType, B.PhoneNumber, ISNULL(A.Post, '') AS Post, A.EmpID, ISNULL(A.EmpName, '') AS EmpName, 
                      A.Agency, ISNULL(A.Office, '') AS Office, B.DialedNumber, B.DialedDatetime, B.CallDurationSecond, B.Cost, B.isPersonal, 
                      CASE WHEN isPersonal = 'Y' THEN 'Personal' ELSE 'Official' END AS CallType
FROM         dbo.vwCellPhoneNumberList AS A INNER JOIN
                      dbo.CellPhoneDt AS B ON A.PhoneNumber = B.PhoneNumber
WHERE     (A.BillFlag = 'Y') AND (B.DialedNumber NOT LIKE '385%' OR
                      B.DialedNumber LIKE '007%' OR
                      B.DialedNumber LIKE '008%')
GO
