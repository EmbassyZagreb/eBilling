/****** Object:  View [dbo].[vwAgencyFunding]    Script Date: 12/02/2014 15:00:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
-- Update vwAgencyFunding
CREATE VIEW [dbo].[vwAgencyFunding]
AS
SELECT     AgencyID, AgencyFundingCode, AgencyDesc, FiscalStripVAT, FiscalStripNonVAT, Disabled, CASE WHEN
                          (SELECT     COUNT(AgencyID)
                            FROM          MonthlyBilling B
                            WHERE      B.AgencyID = A.AgencyID) > 0 THEN 'Y' ELSE 'N' END AS ExistInMonthlyBilling
FROM         dbo.AgencyFunding AS A
GO
