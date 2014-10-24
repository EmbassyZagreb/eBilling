/****** Object:  View [dbo].[vwMonthlyBilling]    Script Date: 08/01/2014 13:35:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[vwMonthlyBilling]
AS
SELECT DISTINCT 
                      A.EmpID, B.LoginID, A.MonthP, A.YearP, B.EmpName, B.Agency, B.Office, B.WorkingTitle, A.PhoneNumber AS MobilePhone, A.BillFlag, 
                      ISNULL(A.ExchangeRate, 0) AS ExchangeRate, ROUND(ISNULL(A.HomePhoneBillRp, 0) * 100, 0) / 100 AS HomePhoneBillRp, 
                      ROUND(ISNULL(A.HomePhoneBillDlr, 0), 0) AS HomePhoneBillDlr, ROUND(ISNULL(A.HomePhonePrsBillRp, 0) * 100, 0) / 100 AS HomePhonePrsBillRp, 
                      ROUND(ISNULL(A.HomePhonePrsBillDlr, 0), 0) AS HomePhonePrsBillDlr, ROUND(ISNULL(A.OfficePhonePrsBillRp, 0) * 100, 0) 
                      / 100 AS OfficePhonePrsBillRp, ROUND(ISNULL(A.OfficePhonePrsBillDlr, 0), 0) AS OfficePhonePrsBillDlr, ROUND(ISNULL(A.OfficePhoneBillRp, 0) * 100, 
                      0) / 100 AS OfficePhoneBillRp, ROUND(ISNULL(A.OfficePhoneBillDlr, 0), 0) AS OfficePhoneBillDlr, ROUND(ISNULL(A.CellPhoneBillRp, 0) * 100, 0) 
                      / 100 AS CellPhoneBillRp, ROUND(ISNULL(A.CellPhoneBillDlr, 0), 2) AS CellPhoneBillDlr, ROUND(ISNULL(A.CellPhonePrsBillRp, 0) * 100, 2) 
                      / 100 AS CellPhonePrsBillRp, ROUND(ISNULL(A.CellPhonePrsBillDlr, 0), 2) AS CellPhonePrsBillDlr, ISNULL(A.TotalAM, 0) AS TotalAM, 
                      ISNULL(A.TotalPM, 0) AS TotalPM, ISNULL(A.TotalAMPM, 0) AS TotalAMPM, ROUND(ISNULL(A.TotalShuttleBillRp, 0) * 100, 0) / 100 AS TotalShuttleBillRp,
                       ROUND(ISNULL(A.TotalShuttleBillDlr, 0), 0) AS TotalShuttleBillDlr, ROUND(ISNULL(A.TotalBillingRp, 0) * 100, 0) / 100 AS TotalBillingRp, 
                      ROUND(ISNULL(A.TotalBillingDlr, 0), 2) AS TotalBillingDlr, A.ProgressId, A.ProgressIdDate, C.ProgressDesc, ISNULL(A.Notes, '') AS Notes, 
                      B.SupervisorId AS ReportTo, ISNULL(E.EmpName, '') AS Supervisor, CASE WHEN len(isNull(A.SupervisorEmail, '')) < 5 THEN isNull(E.EmailAddress, '') 
                      ELSE isNull(A.SupervisorEmail, '') END AS SupervisorEmail, ISNULL(A.SupervisorRemark, '') AS SupervisorRemark, CASE WHEN len(B.EmailAddress) 
                      < 5 THEN ISNULL(B.AlternateEmail, '') ELSE B.EmailAddress END AS EmailAddress, ISNULL(A.ReceiptNo, '') AS ReceiptNo, ISNULL(A.PaidAmountDlr, 
                      0) AS PaidAmountDlr, ISNULL(A.PaidAmountRp, 0) AS PaidAmountRp, ISNULL(CONVERT(Varchar(15), A.PaidDate, 106), '') AS PaidDate, 
                      ISNULL(A.CashierRemark, '') AS CashierRemark, DATEDIFF(dd, A.SupervisorApproveDate, GETDATE()) AS Outstanding, 
                      CASE WHEN (A.ProgressID < 6 OR
                      A.ProgressID = 8) THEN 'Pending' ELSE 'Completed' END AS Status, A.SendMailStatusID, F.SendMailStatusDesc, ISNULL(CONVERT(Varchar(15), 
                      A.SendMailDate, 106), '') AS SendMailDate, ((ROUND(ISNULL(A.HomePhonePrsBillRp, 0) * 100, 0) / 100 + ROUND(ISNULL(A.OfficePhonePrsBillRp, 0) 
                      * 100, 0) / 100) + ROUND(ISNULL(A.CellPhonePrsBillRp, 0) * 100, 0) / 100) + ROUND(ISNULL(A.TotalShuttleBillRp, 0) * 100, 0) 
                      / 100 AS TotalBillingAmountPrsRp, (ROUND(ISNULL(A.HomePhonePrsBillDlr, 0), 0) + ROUND(ISNULL(A.OfficePhonePrsBillDlr, 0), 0) 
                      + ROUND(ISNULL(A.CellPhonePrsBillDlr, 0), 2)) + ROUND(ISNULL(A.TotalShuttleBillRp, 0) * 100, 0) / 100 AS TotalBillingAmountPrsDlr, 
                      CASE WHEN len(B.EmailAddress) < 5 AND ISNULL(B.AlternateEmail, '') <> '' THEN 'Y' ELSE 'N' END AS AlternateEmailFlag, 
                      CASE WHEN isNull(B.EmpType, '') = 'Dummy' THEN 'Y' ELSE 'N' END AS DummyFlag, CASE WHEN B.Office IN ('CLO', 'FMO/Bud', 'FMO/VOU', 'GSO', 
                      'GSO/Motor', 'GSO/Procur', 'GSO/SH-CUS', 'GSO/SHIP', 'GSO/TRAV', 'GSO/WH-SUP', 'HR', 'IM', 'IM/Mail', 'IM/Mail/FP', 'IM/Prog', 'IM/REC', 'IM/TEL/MAI', 
                      'IM/TEL/RAD', 'ISC', 'MGMT') THEN 'MGT' WHEN B.Office IN ('FMO', 'FMO/Cash') THEN 'FMO' ELSE B.Office END AS SectionGroup, B.AgencyID, 
                      ISNULL(B.AgencyFundingCode, '') AS AgencyFundingCode, ISNULL(B.AgencyFunding, '') AS AgencyFunding, B.FiscalStripNonVAT, 
                      ISNULL(B.FiscalStripVAT, '') AS FiscalStripVAT, ISNULL(E.EmpName, '') AS ApprovalSupervisor, 
                      CASE WHEN B.EmpType = 'Dummy' THEN 'Non open account/ have more than one phone number' ELSE '' END AS Note
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
         Configuration = "(H (1[41] 4[20] 2[39] 3) )"
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
               Top = 17
               Left = 35
               Bottom = 125
               Right = 230
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "B"
            Begin Extent = 
               Top = 20
               Left = 648
               Bottom = 128
               Right = 827
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "D"
            Begin Extent = 
               Top = 135
               Left = 385
               Bottom = 243
               Right = 549
            End
            DisplayFlags = 280
            TopColumn = 1
         End
         Begin Table = "C"
            Begin Extent = 
               Top = 194
               Left = 686
               Bottom = 287
               Right = 837
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "E"
            Begin Extent = 
               Top = 259
               Left = 376
               Bottom = 367
               Right = 555
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "F"
            Begin Extent = 
               Top = 222
               Left = 38
               Bottom = 300
               Right = 214
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
      Begin ColumnWidths = 63
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width ' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwMonthlyBilling'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'= 1500
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
         Width = 1950
         Width = 2580
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
         Column = 5550
         Alias = 1755
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
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwMonthlyBilling'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwMonthlyBilling'
GO
