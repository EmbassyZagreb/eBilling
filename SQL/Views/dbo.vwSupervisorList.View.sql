/****** Object:  View [dbo].[vwSupervisorList]    Script Date: 12/02/2014 15:00:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
-- Update vwSupervisorList
CREATE VIEW [dbo].[vwSupervisorList]
AS
SELECT     EmpID, EmpName, EmailAddress
FROM         dbo.vwDirectReport AS A
WHERE     (EmailAddress IN
                          (SELECT     SupervisorEmail
                            FROM          dbo.MonthlyBilling
                            WHERE      (LEN(SupervisorEmail) > 1)))
GO
