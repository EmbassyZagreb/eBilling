/****** Object:  StoredProcedure [dbo].[spGetPaymentList]    Script Date: 12/02/2014 15:00:17 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
CREATE Procedure [dbo].[spGetPaymentList]
		@Type varchar(1)=Null
		,@Month varchar(2)=Null
		,@Year varchar(4)=Null
		,@Extension varchar(50)=Null
		,@EmpName Varchar(50)=Null
		,@OfficeSection Varchar(50)=Null
		,@Outstanding int=0
		,@Status varchar(1)=Null	
As
	
	Set @Extension=Nullif(@Extension,'')
	Set @EmpName=Nullif(@EmpName,'')
	Set @OfficeSection=Nullif(@OfficeSection,'')
	Set @Outstanding=Nullif(@Outstanding,0)
--	Set @Status=Nullif(@Status,'')
If @Type='1'
Begin
	Select A.Extension, B.EmpName, B.Office As OfficeLocation, A.MonthP, A.YearP, A.Notes, SpvEmail
		, Case When A.Status='S' Then 'Submitted' When A.Status='A' Then 'Approved' 
			When A.Status='C' Then 'Correction'  When A.Status='P' Then 'Paid'
		  Else 'Pending' End Status, B.EmailAddress As EmpEmail, SpvRemark
		, Sum(Case when C.isPersonal ='Y' then Cost else 0 end) As PersonalCost , Sum(Cost) As TotalCost, A.ReceiptNo
		, isNull(A.PaidDate,'') As PaidDate, isNull(A.PaidAmount,0) As PaidAmount, A.CashierRemark, Datediff(dd, A.SpvApprovalDate,Getdate()) As Outstanding
	From BillingHd A
	Inner Join vwOfficePhoneNumberList B on (A.Extension=B.PhoneNumber)	
	Inner Join BillingDt C on (A.Extension=C.Extension And A.MonthP=C.MonthP And A.YearP=C.YearP)
	Where (A.MonthP=@Month or @Month is Null)
	And (A.YearP=@Year or @Year is Null)
	And (A.Extension=@Extension or @Extension is Null)
	And (B.EmpName like @EmpName+'%' or @EmpName is Null)
	And (LTrim(RTrim(B.Office))=@OfficeSection or @OfficeSection is Null)
	And (Datediff(dd, A.SpvApprovalDate,Getdate()) = @Outstanding or @Outstanding is Null)
	And (isNull(A.Status,'')=@Status or @Status='X')
	Group By A.Extension, B.EmpName, B.Office, A.MonthP, A.YearP, A.Notes, SpvEmail, A.Status, B.EmailAddress, SpvRemark, A.ReceiptNo, A.PaidDate,  isNull(A.PaidAmount,0), A.CashierRemark, Datediff(dd, A.SpvApprovalDate,Getdate())
	Order By  B.EmpName
End
Else
If @Type='2'
Begin
	Declare @ExchangeRate numeric(9)
		,@PaymentDueDate Varchar(2)

	Select @ExchangeRate=isNull(ExchangeRate,0) From ExchangeRate Where ExchangeMonth=@Month And ExchangeYear=@Year

	Select @PaymentDueDate=Right('00'+Convert(Varchar(2),PaymentDueDate),2) From PaymentDueDate
--	print @Month
--	print @Year
/*
	Select A.[ID], A.MonthP, A.YearP, A.Nomor, NamaPelanggan
		, Case When isNull(B.FirstName,'')='' Then B.LastName Else B.LastName+', '+B.FirstName End As EmpName , B.Office As OfficeLocation
		, B.EmailAddress As EmpEmail, B.LoginID, NoTagihan, Abonemen, Lokal, SLJJ, STB, JAPATI, SLI007, [001+008], [17], OPERATOR
		, AIRTIME, QUOTA, LAIN2, PPN, METERAI, TOTAL, Case When A.Status='P' Then 'Paid'  Else 'Pending' End Status, ReceiptNo
		, isNull(PaidAmount,0) As PaidAmount, isNull(PaidDate,'') As PaidDate, CashierRemark, Datediff(dd, Convert(datetime, A.MonthP+'/'+@PaymentDueDate+'/'+A.YearP),Getdate()) As Outstanding
		, isNull(Round(isNull(TOTAL,0)/@ExchangeRate,0),0)  As TotalDollar, isNull(@ExchangeRate,0) As ExchangeRate, isNull(A.PaidCurrency,'') As PaidCurrency
	From HomePhone A
	Inner Join MsHomePhoneNumber C on (A.Nomor=C.PhoneNumber)
	Inner Join vwPhoneCustomerList B on (C.EmpID=B.EmpID)
	Where --B.TypeE='AMER' And 
	B.Status='C' And A.MonthP=@Month And A.YearP=@Year 
--	And (RTrim(B.LoginID)=@Extension or @Extension is Null)
	And (B.HomePhone=@Extension or @Extension is Null)
	And (B.LastName like @LastName+'%' or @LastName is Null)
	And (B.FirstName like @FirstName+'%' or @FirstName is Null)
	And (isNull(A.Status,'')=@Status or @Status='X')
	And (Datediff(dd, Convert(datetime, A.MonthP+'/'+@PaymentDueDate+'/'+A.YearP),Getdate()) = @Outstanding or @Outstanding is Null)
	And (LTrim(RTrim(B.Office))=@OfficeSection or @OfficeSection is Null)
	Order By  B.LastName, B.FirstName
*/
	Select C.[ID], A.EmpID, A.MonthP, A.YearP, B.PhoneNumber, D.EmpName
		, D.Office As OfficeLocation, isNull(A.HomePhonePrsBillRp,0) As HomePhonePrsBillRp, isNull(A.HomePhonePrsBillDlr,0) As HomePhonePrsBillDlr, A.ExchangeRate
		, Case When C.Status='P' Then 'Paid'  Else 'Pending' End Status, C.ReceiptNo, isNull(C.PaidAmount,0) As PaidAmount
		, isNull(C.PaidDate,'') As PaidDate, C.CashierRemark, Datediff(dd, Convert(datetime, A.MonthP+'/'+@PaymentDueDate+'/'+A.YearP),Getdate()) As Outstanding
		, isNull(C.PaidCurrency,'') As PaidCurrency, A.ProgressID, Case When C.Status='P' Then 'Paid'  Else 'Pending' End As Status
	From MonthlyBilling A
	Inner Join MsHomePhoneNumber B on (A.EmpID=B.EmpID)
	Inner Join HomePhone C on (B.PhoneNumber=C.Nomor And A.MonthP=C.MonthP And A.YearP=C.YearP)
	Inner Join vwHomePhoneNumberList D on (A.EmpID=D.EmpID)
	Where A.MonthP=@Month And A.YearP=@Year 
	And (B.PhoneNumber=@Extension or @Extension is Null)
	And (D.EmpName like @EmpName+'%' or @EmpName is Null)
	And (isNull(C.Status,'')=@Status or @Status='X')
	And (Datediff(dd, Convert(datetime, A.MonthP+'/'+@PaymentDueDate+'/'+A.YearP),Getdate()) = @Outstanding or @Outstanding is Null)
	And (LTrim(RTrim(D.Office))=@OfficeSection or @OfficeSection is Null)
	Order By  D.EmpName

End
GO
