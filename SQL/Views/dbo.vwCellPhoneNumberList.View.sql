/****** Object:  View [dbo].[vwCellPhoneNumberList]    Script Date: 12/02/2014 15:00:27 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
-- Update vwCellPhoneNumberList
CREATE VIEW [dbo].[vwCellPhoneNumberList]
AS
SELECT DISTINCT 
                      A.ID, A.PhoneNumber, A.PhoneType, C.PhoneTypeName, A.EmpID, CASE WHEN isNull(A.EmpID, '') 
                      = '' THEN 'Vacant' ELSE B.EmpName END AS EmpName, B.Post, ISNULL(B.Agency, '') AS Agency, B.Office, CASE WHEN B.Office IN ('CLO', 'FMO/Bud', 
                      'FMO/VOU', 'GSO', 'GSO/Motor', 'GSO/Procur', 'GSO/SH-CUS', 'GSO/SHIP', 'GSO/TRAV', 'GSO/WH-SUP', 'HR', 'IM', 'IM/Mail', 'IM/Mail/FP', 'IM/Prog', 
                      'IM/REC', 'IM/TEL/MAI', 'IM/TEL/RAD', 'ISC', 'MGMT') THEN 'MGT' WHEN B.Office IN ('FMO', 'FMO/Cash') 
                      THEN 'FMO' ELSE B.Office END AS SectionGroup, ISNULL(B.EmailAddress, '') AS EmailAddress, ISNULL(B.AlternateEmail, '') AS AlternateEmail, 
                      ISNULL(A.Remark, '') AS Remark, A.BillFlag, A.OwnerID, D.EmpName AS OwnerName, A.Discontinued, 
                      CASE WHEN A.Discontinued = 'Y' THEN 'Yes' ELSE 'No' END AS DiscontinuedDesc, CASE WHEN A.DiscontinuedDate IS NULL 
                      THEN '' ELSE CONVERT(Varchar(15), A.DiscontinuedDate, 106) END AS DiscontinuedDate, CASE WHEN
                          (SELECT     COUNT(PhoneNumber)
                            FROM          MonthlyBilling B
                            WHERE      B.PhoneNumber = A.PhoneNumber) > 0 THEN 'Y' ELSE 'N' END AS ExistInMonthlyBilling
FROM         dbo.MsCellPhoneNumber AS A LEFT OUTER JOIN
                      dbo.vwPhoneCustomerList AS B ON A.EmpID = B.EmpID LEFT OUTER JOIN
                      dbo.PhoneType AS C ON A.PhoneType = C.PhoneType LEFT OUTER JOIN
                      dbo.vwPhoneCustomerList AS D ON A.OwnerID = D.EmpID
GO
