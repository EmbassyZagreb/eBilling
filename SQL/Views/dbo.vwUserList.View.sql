/****** Object:  View [dbo].[vwUserList]    Script Date: 12/02/2014 15:00:32 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
-- Update vwUserList
CREATE View [dbo].[vwUserList]
As
Select A.LoginId, ISNULL(B.EmpName,'') As EmployeeName
	, ISNULL(B.OfficeSection,'') As OfficeLocation, A.RoleID, C.RoleName
From Users A
Left Join MsEmployee B on (A.LoginID=B.LoginID)
Left Join UserRoles C on (A.RoleID=C.RoleID)
GO
