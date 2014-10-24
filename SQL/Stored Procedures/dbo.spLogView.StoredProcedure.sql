/****** Object:  StoredProcedure [dbo].[spLogView]    Script Date: 08/01/2014 13:31:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[spLogView]
		@MonthP Varchar(2)=''
		,@YearP Varchar(4)=''
		,@ProgressID TinyInt=0
As
/*		
Set @MonthP='09'
Set @YearP='2013'
Set @ProgressID=15
*/
Select A.MonthP, A.YearP, A.ProgressID, C.ProgressDesc, A.EmpID, A.EmpName, A.MobilePhone, A.TotalBillingRp
From vwMonthlyBilling A
Inner Join CellPhoneHd B on (A.MonthP=B.MonthP And A.YearP=B.YearP And A.MobilePhone=B.PHONENUMBER)
Left Join ProgressStatus C on (A.ProgressID=C.ProgressID)
Where A.TotalBillingRp>0 And A.ProgressID=@ProgressID And A.MonthP=@MonthP And A.YearP=@YearP
Union
(
Select A.MonthP, A.YearP, ISNULL(A.ProgressID,0), ISNULL(C.ProgressDesc,''), ISNULL(A.EmpID,''), ISNULL(B.EmpName,'') , A.PhoneNumber, A.Balance
From Reconciliation A
Left Join vwPhoneCustomerList B on (A.EmpID=B.EmpID)
Left Join ProgressStatus C on (A.ProgressID=C.ProgressID)
Where A.ProgressID=@ProgressID And MonthP=@MonthP And YearP=@YearP
)
Order by A.EmpName
GO
