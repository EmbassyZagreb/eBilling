/****** Object:  View [dbo].[vwSupervisorReminderRpt]    Script Date: 08/01/2014 13:35:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[vwSupervisorReminderRpt]
AS
SELECT     A.EmpID, C.EmpName, C.Agency, C.Office, C.LoginID, A.MonthP, A.YearP, A.HomePhonePrsBillRp, A.OfficePhonePrsBillRp, 
                      A.PhoneNumber AS MobilePhone, A.CellPhonePrsBillRp, A.TotalShuttleBillRp, ISNULL(A.HomePhonePrsBillRp, 0) + ISNULL(A.OfficePhonePrsBillRp, 0) 
                      + ISNULL(A.CellPhonePrsBillRp, 0) + ISNULL(A.TotalShuttleBillRp, 0) AS TotalBillingAmountPrsRp, 
                      CASE WHEN A.ProgressId = 2 THEN ISNULL(DATEDIFF(dd, A.SendMailDate, GETDATE()), 0) ELSE ISNULL(DATEDIFF(dd, A.SendMailDate, 
                      A.SupervisorApproveDate), 0) END AS Aging, A.SendMailDate, C.EmailAddress, C.AgencyFunding, A.SendMailStatusID, F.SendMailStatusDesc, 
                      A.ProgressId, ISNULL(G.EmpName, '') AS Supervisor, ISNULL(G.EmailAddress, '') AS SupervisorEmail, A.CellPhoneBillRp
FROM         dbo.MonthlyBilling AS A LEFT OUTER JOIN
                      dbo.vwDirectReport AS C ON A.EmpID = C.EmpID LEFT OUTER JOIN
                      dbo.SendMailStatus AS F ON A.SendMailStatusID = F.SendMailStatusID LEFT OUTER JOIN
                      dbo.vwDirectReport AS G ON A.SupervisorEmail = G.EmailAddress AND G.Type = 'AMER'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
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
         Begin Table = "A"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 114
               Right = 233
            End
            DisplayFlags = 280
            TopColumn = 11
         End
         Begin Table = "C"
            Begin Extent = 
               Top = 6
               Left = 271
               Bottom = 114
               Right = 450
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "F"
            Begin Extent = 
               Top = 114
               Left = 38
               Bottom = 192
               Right = 214
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "G"
            Begin Extent = 
               Top = 6
               Left = 488
               Bottom = 114
               Right = 667
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
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
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwSupervisorReminderRpt'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwSupervisorReminderRpt'
GO
