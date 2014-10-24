/****** Object:  StoredProcedure [dbo].[spGenerateMonthlyBilling]    Script Date: 08/01/2014 13:31:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[spGenerateMonthlyBilling]
	@Month varchar(2)=Null
	,@Year varchar(4)=Null
	,@MobilePhone Varchar(30)=''
	,@AlwaysExemptedCallType Varchar(255)
	,@ExemptedIfOfficialCallType Varchar(255) 
As

/*Update Cellphone Master List*/
Insert MsCellPhoneNumber(PhoneNumber, PhoneType, EmpID, BillFlag, OwnerID, CreateBy)
Select MobilePhone, 'B', EmpID, 'Y', EmpID, 'Automation Job'
From vwPhoneCustomerList
Where MobilePhone in 
(Select PhoneNumber
From dbo.CellPhoneHd A
Left Join vwPhoneCustomerList B on (A.PhoneNumber=B.MobilePhone)
Where PhoneNumber not in (Select PhoneNumber from MsCellPhoneNumber)
And A.MonthP=@Month And A.YearP=@Year 
)

--Update Status Personal Call
Declare @ExtNo varchar(10)
		,@CeilingAmount Numeric(9,2)
		,@DetailRecordAmount Numeric(9)
		,@PhoneNumber Varchar(50)
		,@PhoneNumberPersonal Varchar(50)			
Select @CeilingAmount=CeilingAmount, @DetailRecordAmount=DetailRecordAmount From PaymentDueDate

If LEN(@MobilePhone)>1 
Begin
	Update CellPhoneDt Set isPersonal='N' Where MonthP=@Month And YearP=@Year And PhoneNumber=@MobilePhone 
	Update CellPhoneDt Set isPersonal='Y' Where Cost>@DetailRecordAmount And MonthP=@Month And YearP=@Year And PhoneNumber=@MobilePhone 
												And CallType not like @ExemptedIfOfficialCallType And CallType not like @AlwaysExemptedCallType  
	Select @PhoneNumberPersonal=PhoneNumber From MsCellPhoneNumber Where PhoneNumber=@MobilePhone And BillFlag='P'
	Update CellPhoneDt Set isPersonal='Y' Where Cost>@DetailRecordAmount And MonthP=@Month And YearP=@Year And PhoneNumber=@PhoneNumberPersonal 
												And CallType not like @AlwaysExemptedCallType
End
Else
Begin
	Update CellPhoneDt Set isPersonal='N' Where MonthP=@Month And YearP=@Year 
	Update CellPhoneDt Set isPersonal='Y' Where Cost>@DetailRecordAmount And MonthP=@Month And YearP=@Year  
												And CallType not like @ExemptedIfOfficialCallType And CallType not like @AlwaysExemptedCallType  
	Update CellPhoneDt Set isPersonal='Y' Where Cost>@DetailRecordAmount And MonthP=@Month And YearP=@Year And PhoneNumber In (Select PhoneNumber From MsCellPhoneNumber Where BillFlag='P')
												And CallType not like @AlwaysExemptedCallType  
End

--Monthly Bill
Declare @ShuttleBusRate Numeric(5,2)
	,@ExchangeRate Numeric(9,2)

Select @ExchangeRate=isNull(ExchangeRate,0) from ExchangeRate
Where ExchangeMonth=@Month And ExchangeYear=@Year

Delete MonthlyBilling 
Where MonthP=@Month And YearP=@Year And (PhoneNumber=@MobilePhone or @MobilePhone='')

Insert MonthlyBilling(EmpID, MonthP, YearP, PhoneNumber, BillFlag, ExchangeRate
	, CellPhoneBillRp, CellPhoneBillDlr, CellPhonePrsBillRp, CellPhonePrsBillDlr
	, TotalBillingRp, TotalBillingDlr, ProgressId, ProgressIdDate)
SELECT DISTINCT   
                      D.EmpID AS EmpID, @Month AS MonthP, @Year AS YearP, D.MobilePhone AS PhoneNumber, A.BillFlag, @ExchangeRate AS ExchangeRate
						, isNull(B.TOTBILLAMOUNT,0) AS CellPhoneBillRp
						, ISNULL(ROUND(ISNULL(ISNULL(B.TOTBILLAMOUNT, 0) / @ExchangeRate, 0), 2), 0) AS TotalBillingDlr
						, ISNULL(SUM(CASE WHEN C.isPersonal = 'Y' THEN C.Cost ELSE 0 END), 0) AS CellPhonePrsBillRp
						, ISNULL(ROUND(ISNULL(ISNULL(SUM(CASE WHEN C.isPersonal = 'Y' THEN C.Cost ELSE 0 END), 0) / @ExchangeRate, 0), 2), 0) AS CellPhonePrsBillDlr
						, ISNULL(B.TOTBILLAMOUNT, 0) AS TotalBillingRp
						, ISNULL(ROUND(ISNULL(ISNULL(B.TOTBILLAMOUNT, 0) / @ExchangeRate, 0), 2), 0) AS TotalBillingDlr
						, 1 , GETDATE()
FROM         dbo.vwCellPhoneNumberList AS A INNER JOIN
					   dbo.CellPhoneHd AS B ON B.PhoneNumber = A.PhoneNumber AND B.MonthP = @Month AND B.YearP = @Year LEFT OUTER JOIN
                       dbo.CellPhoneDt AS C ON C.PhoneNumber = A.PhoneNumber AND C.MonthP = @Month AND C.YearP = @Year LEFT OUTER JOIN
					   vwPhoneCustomerList AS D ON D.MobilePhone = A.PhoneNumber
WHERE     ((ISNULL(B.TOTBILLAMOUNT, 0) >= 0) AND (A.Phonenumber = @MobilePhone OR @MobilePhone = '') And (A.BillFlag <> 'N')) 
GROUP BY  D.EmpID, D.MobilePhone, C.MonthP, C.YearP,  A.BillFlag,ISNULL(B.TOTBILLAMOUNT, 0), ROUND(ISNULL(ISNULL(B.TOTBILLAMOUNT, 0) / @ExchangeRate, 0), 2) 


--Update CellPhone Call duration in second
Update CellPhoneDt set CallDurationSecond=dbo.GetTotDuration(callduration)
Where MonthP=@Month and YearP=@Year


--Update payment status to 'Awaiting Payment' for employee who has personal phone
Update MonthlyBilling Set ProgressId=4, ProgressIdDate=GETDATE()
Where MonthP=@Month and YearP=@Year And (Phonenumber = @MobilePhone OR @MobilePhone = '') And BillFlag='P'


--Update payment status to 'Awaiting Off-line Review' for employee who doesn't has open net account
Update MonthlyBilling Set ProgressId=8, ProgressIdDate=GETDATE()
Where MonthP=@Month and YearP=@Year And (Phonenumber = @MobilePhone OR @MobilePhone = '')
And PhoneNumber in (Select PhoneNumber From vwCellPhoneNumberList Where BillFlag<>'N' And len(ISNULL(EmailAddress,''))<8)


--Update payment status to 'Closed - Below Threshold' for billing amount smaller than equal to Ceiling amount parameter
Update MonthlyBilling Set ProgressId=7, ProgressIdDate=GETDATE() 
Where ((CellPhonePrsBillRp)<=@CeilingAmount or (CellPhonePrsBillRp = 0))
And MonthP=@Month and YearP=@Year And (Phonenumber = @MobilePhone OR @MobilePhone = '')


--Generate Reconciliation Report

	Exec spGenerateReconRpt @Month, @Year
	
	Exec spGenerateProgressLog @Month, @Year
GO
