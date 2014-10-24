/****** Object:  View [dbo].[vwCellPhoneNumberList]    Script Date: 08/01/2014 13:35:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
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
                      THEN '' ELSE CONVERT(Varchar(15), A.DiscontinuedDate, 106) END AS DiscontinuedDate
FROM         dbo.MsCellPhoneNumber AS A LEFT OUTER JOIN
                      dbo.vwPhoneCustomerList AS B ON A.EmpID = B.EmpID LEFT OUTER JOIN
                      dbo.PhoneType AS C ON A.PhoneType = C.PhoneType LEFT OUTER JOIN
                      dbo.vwPhoneCustomerList AS D ON A.OwnerID = D.EmpID
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[27] 4[45] 2[24] 3) )"
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
               Right = 202
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "B"
            Begin Extent = 
               Top = 6
               Left = 240
               Bottom = 114
               Right = 419
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "C"
            Begin Extent = 
               Top = 114
               Left = 38
               Bottom = 192
               Right = 199
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "D"
            Begin Extent = 
               Top = 114
               Left = 237
               Bottom = 222
               Right = 416
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
         Column = 6090
         Alias = 1455
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
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwCellPhoneNumberList'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vwCellPhoneNumberList'
GO
