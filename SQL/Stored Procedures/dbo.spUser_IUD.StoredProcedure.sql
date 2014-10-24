/****** Object:  StoredProcedure [dbo].[spUser_IUD]    Script Date: 08/01/2014 13:31:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[spUser_IUD] 
	(@Mode Varchar(1),
	 @LoginID varchar(50)=Null,
	 @RoleID varchar(10)=Null)
AS
If @Mode='I'
Begin
	Insert Users(LoginID,RoleID) 
	Values(@LoginID,@RoleID)
End
Else If @Mode='E'
Begin
	Update Users Set RoleID = @RoleID Where LoginID = @LoginID
End
Else If @Mode='D'
Begin
	Delete Users Where LoginID=@LoginID
End
GO
