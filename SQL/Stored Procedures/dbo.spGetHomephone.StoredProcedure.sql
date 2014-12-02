/****** Object:  StoredProcedure [dbo].[spGetHomephone]    Script Date: 12/02/2014 15:00:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[spGetHomephone]
		 @Type Varchar(1)=Null
		,@HomePhone Varchar(50)=Null
		,@Month Varchar(2)=Null
		,@Year Varchar(4)=Null
As
--	If left(@HomePhone,1)<>'+'
--		Set @HomePhone='+'+@HomePhone
If @Type='1'
Begin
	Select A.[ID], A.MonthP, A.YearP, A.Nomor, NamaPelanggan, B.EmpName 
		, B.Office As OfficeLocation, B.EmailAddress As EmpEmail, B.LoginID, NoTagihan, Abonemen, Lokal, SLJJ, STB, JAPATI, SLI007, [001+008], [17], OPERATOR, AIRTIME, QUOTA, LAIN2
		, PPN, METERAI, TOTAL, Case When A.Status='P' Then 'Paid'  Else 'Pending' End Status, ReceiptNo, PaidAmount, PaidDate, CashierRemark
	From HomePhone A
	Inner Join vwHomePhoneNumberList B on (A.Nomor=B.PhoneNumber)
	Where --B.Type='AMER' And 
	 A.MonthP=@Month And A.YearP=@Year And B.PhoneNumber=@HomePhone
End
If @Type='2'
Begin
	Select PhoneNumber, MonthP, YearP, CallRecordID, DialedDatetime, DialedNumber, CallDuration, Cost, isPersonal From HomePhoneDt 
	Where PhoneNumber=@HomePhone and MonthP=@Month and YearP=@Year
End
GO
