/****** Object:  StoredProcedure [dbo].[spValidation]    Script Date: 12/02/2014 15:00:24 ******/
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
