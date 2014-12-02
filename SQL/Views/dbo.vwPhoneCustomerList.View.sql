/****** Object:  View [dbo].[vwPhoneCustomerList]    Script Date: 12/02/2014 15:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
-- Update vwPhoneCustomerList
CREATE VIEW [dbo].[vwPhoneCustomerList]
AS
SELECT     A.EmpID, ISNULL(A.EmpName, '') AS EmpName, ISNULL(A.Post, '') AS Post, ISNULL(A.EmpType, '') AS EmpType, ISNULL(A.Agency, '') AS Agency, 
                      ISNULL(A.OfficeSection, '') AS Office, CASE WHEN A.OfficeSection IN ('CLO', 'FMO/Bud', 'FMO/VOU', 'GSO', 'GSO/Motor', 'GSO/Procur', 'GSO/SH-CUS', 
                      'GSO/SHIP', 'GSO/TRAV', 'GSO/WH-SUP', 'HR', 'IM', 'IM/Mail', 'IM/Mail/FP', 'IM/Prog', 'IM/REC', 'IM/TEL/MAI', 'IM/TEL/RAD', 'ISC', 'MGMT') 
                      THEN 'MGT' WHEN A.OfficeSection IN ('FMO', 'FMO/Cash') THEN 'FMO' ELSE A.OfficeSection END AS SectionGroup, ISNULL(A.WorkingTitle, '') 
                      AS WorkingTitle, ISNULL(D.PhoneNumber, '') AS MobilePhone, ISNULL(A.EmailAddress, '') AS EmailAddress, ISNULL(A.AlternateEmail, '') 
                      AS AlternateEmail, ISNULL(A.SupervisorId, '') AS SupervisorId, ISNULL(C.EmpName, '') AS SupervisorName, ISNULL(A.LoginID, '') AS LoginID, 
                      ISNULL(A.Remark, '') AS Remark, ISNULL(A.Status, '') AS Status, ISNULL(A.AgencyID, 0) AS AgencyID, ISNULL(B.AgencyFundingCode, '') 
                      AS AgencyFundingCode, ISNULL(B.AgencyDesc, '') AS AgencyFunding, ISNULL(B.FiscalStripVAT, '') AS FiscalStripVAT, ISNULL(B.FiscalStripNonVAT, '') 
                      AS FiscalStripNonVAT, B.Disabled AS AgencyDisabled
FROM         dbo.MsEmployee AS A LEFT OUTER JOIN
                      dbo.AgencyFunding AS B ON A.AgencyID = B.AgencyID LEFT OUTER JOIN
                      dbo.MsEmployee AS C ON A.SupervisorId = C.EmpID LEFT OUTER JOIN
                      dbo.MsCellPhoneNumber AS D ON A.EmpID = D.EmpID
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[7] 4[5] 2[17] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 18
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwPhoneCustomerList'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwPhoneCustomerList'
GO
