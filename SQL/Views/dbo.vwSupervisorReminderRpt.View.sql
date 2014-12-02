/****** Object:  View [dbo].[vwSupervisorReminderRpt]    Script Date: 12/02/2014 15:00:32 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[vwSupervisorReminderRpt]
AS
SELECT DISTINCT 
                      A.EmpID, B.EmpName, B.Agency, B.Office, B.LoginID, A.MonthP, A.YearP, A.PhoneNumber AS MobilePhone, ROUND(ISNULL(A.CellPhonePrsBillRp, 0) 
                      * 100, 2) / 100 AS CellPhonePrsBillRp, ((ROUND(ISNULL(A.HomePhonePrsBillRp, 0) * 100, 0) / 100 + ROUND(ISNULL(A.OfficePhonePrsBillRp, 0) * 100, 
                      0) / 100) + ROUND(ISNULL(A.CellPhonePrsBillRp, 0) * 100, 0) / 100) + ROUND(ISNULL(A.TotalShuttleBillRp, 0) * 100, 0) 
                      / 100 AS TotalBillingAmountPrsRp, CASE WHEN A.ProgressId = 2 THEN ISNULL(DATEDIFF(dd, A.SendMailDate, GETDATE()), 0) 
                      ELSE ISNULL(DATEDIFF(dd, A.SendMailDate, A.SupervisorApproveDate), 0) END AS Aging, A.ProgressId, B.SupervisorId AS ReportTo, 
                      ISNULL(E.EmpName, '') AS Supervisor, CASE WHEN len(isNull(A.SupervisorEmail, '')) < 5 THEN isNull(E.EmailAddress, '') 
                      ELSE isNull(A.SupervisorEmail, '') END AS SupervisorEmail, CASE WHEN len(B.EmailAddress) < 5 THEN ISNULL(B.AlternateEmail, '') 
                      ELSE B.EmailAddress END AS EmailAddress, A.SendMailStatusID, F.SendMailStatusDesc, ISNULL(CONVERT(Varchar(15), A.SendMailDate, 106), '') 
                      AS SendMailDate, (ROUND(ISNULL(A.HomePhonePrsBillDlr, 0), 0) + ROUND(ISNULL(A.OfficePhonePrsBillDlr, 0), 0) 
                      + ROUND(ISNULL(A.CellPhonePrsBillDlr, 0), 2)) + ROUND(ISNULL(A.TotalShuttleBillRp, 0) * 100, 0) / 100 AS TotalBillingAmountPrsDlr, 
                      ISNULL(A.AgencyFundingDesc, '') AS AgencyFunding, ROUND(ISNULL(A.CellPhoneBillRp, 0) * 100, 2) / 100 AS CellPhoneBillRp
FROM         dbo.MonthlyBilling AS A INNER JOIN
                      dbo.vwPhoneCustomerList AS B ON A.EmpID = B.EmpID LEFT OUTER JOIN
                      dbo.MsCellPhoneNumber AS D ON A.EmpID = D.EmpID AND D.BillFlag <> 'N' LEFT OUTER JOIN
                      dbo.ProgressStatus AS C ON A.ProgressId = C.ProgressID LEFT OUTER JOIN
                      dbo.vwPhoneCustomerList AS E ON B.SupervisorId = E.EmpID LEFT OUTER JOIN
                      dbo.SendMailStatus AS F ON A.SendMailStatusID = F.SendMailStatusID
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[20] 4[54] 2[17] 3) )"
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
            TopColumn = 0
         End
         Begin Table = "F"
            Begin Extent = 
               Top = 6
               Left = 488
               Bottom = 84
               Right = 664
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "B"
            Begin Extent = 
               Top = 6
               Left = 271
               Bottom = 114
               Right = 450
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "D"
            Begin Extent = 
               Top = 6
               Left = 702
               Bottom = 114
               Right = 866
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "C"
            Begin Extent = 
               Top = 6
               Left = 904
               Bottom = 99
               Right = 1055
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "E"
            Begin Extent = 
               Top = 6
               Left = 1093
               Bottom = 114
               Right = 1272
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
         Column = 17700
         Alias = 2325
         Table = 1170
         ' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwSupervisorReminderRpt'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'Output = 720
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
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwSupervisorReminderRpt'
GO
