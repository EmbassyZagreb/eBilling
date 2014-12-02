/****** Object:  View [dbo].[vwLineUsages]    Script Date: 12/02/2014 15:00:28 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
-- Update vwLineUsages
CREATE VIEW [dbo].[vwLineUsages]
AS
SELECT     'Cell Phone' AS PhoneType, B.PhoneNumber, ISNULL(A.Post, '') AS Post, A.EmpID, ISNULL(A.EmpName, '') AS EmpName, A.Agency, ISNULL(A.Office, '') 
                      AS Office, B.DialedDatetime, B.CallDurationSecond, B.Cost
FROM         dbo.vwCellPhoneNumberList AS A INNER JOIN
                      dbo.CellPhoneDt AS B ON A.PhoneNumber = B.PhoneNumber
WHERE     (A.BillFlag = 'Y')
GO
