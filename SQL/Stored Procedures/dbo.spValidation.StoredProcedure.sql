/****** Object:  StoredProcedure [dbo].[spValidation]    Script Date: 08/01/2014 13:31:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[spValidation]
	@RAMCNo varchar(20)=Null
	,@BirthDate varchar(20)=Null
As
	Select Convert(Varchar(10),FN_Emp_Key_Num)+'L' As EmpID
	FROM [JAKARTAAP02\SQL05].PassPS.dbo.FN_EMP
	Where FN_EMP_RAMC_NUM_TXT=@RAMCNo And FN_EMP_DOB_DT=@BirthDate
	--Print 	@LoginID
GO
