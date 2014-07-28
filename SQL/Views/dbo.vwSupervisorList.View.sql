/****** Object:  View [dbo].[vwSupervisorList]    Script Date: 07/28/2014 12:50:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE view [dbo].[vwSupervisorList]
 As
 /*
 Select Distinct A.ReportTo As EmpID, B.EmpName
 From vwPhoneCustomerList A
 Inner Join vwPhoneCustomerList B on (A.ReportTo=B.EmpID)
 Where LEN(A.ReportTo)>1 
GO
*/


 Select A.EmpID, A.EmpName, A.EmailAddress
 From vwPhoneCustomerList A
 Where EmailAddress in (Select SupervisorEmail From MonthlyBilling Where LEN(SupervisorEmail)>1)
GO
