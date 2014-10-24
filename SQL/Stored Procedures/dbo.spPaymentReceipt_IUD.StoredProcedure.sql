/****** Object:  StoredProcedure [dbo].[spPaymentReceipt_IUD]    Script Date: 08/01/2014 13:31:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[spPaymentReceipt_IUD]
	(@Mode varchar(1)=Null,
	 @PaymentNo int=0,
	 @EmpID varchar(10)=Null,
	 @MonthP varchar(2)=Null,
	 @YearP varchar(4)=Null,
	 @PhoneNumber varchar(30)=Null,
	 @ReceiptNo varchar(30)=Null,
	 @Currency varchar(3)=Null,
	 @PaidAmountDlr numeric(18,2)=Null,
	 @PaidAmountRp numeric(18,2)=Null,
	 @PaidDate datetime=Null,
	 @CashierRemark varchar(500)=Null,
	 @PaymentType varchar(1)=Null)
AS 
If @Mode='I'
Begin
	INSERT INTO PaymentReceipt (EmpID, MonthP, YearP, PhoneNumber, ReceiptNo, Currency, PaidAmountDlr, PaidAmountRp, PaidDate, CashierRemark, PaymentType) 
	VALUES ( @EmpID, @MonthP, @YearP, @PhoneNumber, @ReceiptNo, @Currency, @PaidAmountDlr, @PaidAmountRp, @PaidDate, @CashierRemark, @PaymentType)

	--Update MOnthly Bill Status
	Update MonthlyBilling Set PaidAmountDlr=isNull(PaidAmountDlr,0)+@PaidAmountDlr, PaidAmountRp=isNull(PaidAmountRp,0)+@PaidAmountRp
			, ProgressId=Case When @PaymentType='P' Then 5 When @PaymentType='F' Then 6 End, ReceiptNo = @ReceiptNo, PaidDate = @PaidDate
			, ProgressIdDate=GetDate()
	Where EmpID=@EmpID And PhoneNumber=@PhoneNumber And MonthP=@MonthP And YearP=@YearP
End
Else If @Mode='U'
Begin
	UPDATE PaymentReceipt SET ReceiptNo = @ReceiptNo, Currency=@Currency , PaidAmountDlr=isNull(PaidAmountDlr,0)+@PaidAmountDlr, PaidAmountRp=isNull(PaidAmountRp,0)+@PaidAmountRp, PaidDate = @PaidDate
		, CashierRemark	= @CashierRemark, PaymentType = @PaymentType
	WHERE (PaymentNo = @PaymentNo)
End
Else If @Mode='D'
Begin
	Delete PaymentReceipt WHERE PaymentNo = @PaymentNo
End
GO
