/****** Object:  StoredProcedure [dbo].[spGenerateMonthlyBilling]    Script Date: 07/28/2014 12:45:36 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[spGenerateMonthlyBilling]
	@Month varchar(2)=Null
	,@Year varchar(4)=Null
	,@EmpID varchar(10)=''
	,@AlwaysExemptedCallType Varchar(255) --*Marin-in
	,@ExemptedIfOfficialCallType Varchar(255) --*Marin-in
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

If LEN(@EmpID)>1
Begin
	Select @PhoneNumber=PhoneNumber From MsCellPhoneNumber Where EmpID=@EmpID
	Update CellPhoneDt Set isPersonal='N' Where MonthP=@Month And YearP=@Year And PhoneNumber=@PhoneNumber
	Update CellPhoneDt Set isPersonal='Y' Where Cost>@DetailRecordAmount And MonthP=@Month And YearP=@Year And PhoneNumber=@PhoneNumber 
												And CallType not like @ExemptedIfOfficialCallType And CallType not like @AlwaysExemptedCallType  
	Select @PhoneNumberPersonal=PhoneNumber From MsCellPhoneNumber Where EmpID=@EmpID And BillFlag='P'
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

/*
Select @ShuttleBusRate=isNull(ShuttleBusRate,0) from ShuttleBusRate
Where ShuttleBusRateID=
(Select Max(ShuttleBusRateID) from ShuttleBusRate)
*/

Select @ExchangeRate=isNull(ExchangeRate,0) from ExchangeRate
Where ExchangeMonth=@Month And ExchangeYear=@Year


Select @PhoneNumber=PhoneNumber From MsCellPhoneNumber Where EmpID=@EmpID 
Delete MonthlyBilling 
Where MonthP=@Month And YearP=@Year And (PhoneNumber=@PhoneNumber or EmpID=@EmpID or @EmpID='') 


Insert MonthlyBilling(EmpID, MonthP, YearP, PhoneNumber, BillFlag, ExchangeRate
	, CellPhoneBillRp, CellPhoneBillDlr, CellPhonePrsBillRp, CellPhonePrsBillDlr
	, TotalBillingRp, TotalBillingDlr, ProgressId, ProgressIdDate)
/*
Insert MonthlyBilling(EmpID, MonthP, YearP, PhoneNumber, BillFlag, ExchangeRate, HomePhoneBillRp, HomePhoneBillDlr, HomePhonePrsBillRp  
	, HomePhonePrsBillDlr, OfficePhonePrsBillRp, OfficePhonePrsBillDlr, OfficePhoneBillRp, OfficePhoneBillDlr
	, CellPhoneBillRp, CellPhoneBillDlr, CellPhonePrsBillRp, CellPhonePrsBillDlr
	, TotalShuttleBusUsed, TotalShuttleBillRp, TotalShuttleBillDlr, TotalBillingRp, TotalBillingDlr, ProgressId, ProgressIdDate)
*/
Select A.EmpID, @Month As MonthP, @Year As YearP, F.PhoneNumber, F.BillFlag, @ExchangeRate As ExchangeRate
--	, isNull(C.HomePhoneBillRp,0) As HomePhoneBillRp
--	, isNull(C.HomePhoneBillDlr,0) As HomePhoneBillDlr
--	, isNull(C.HomePhonePrsBillRp,0) As HomePhonePrsBillRp  
--	, isNull(C.HomePhonePrsBillDlr,0) As HomePhonePrsBillDlr
--	, isNull(B.OfficePhonePrsBillRp,0) As OfficePhonePrsBillRp
--	, isNull(B.OfficePhonePrsBillDlr,0) As OfficePhonePrsBillDlr
--	, isNull(B.OfficePhoneBillRp,0) As OfficePhoneBillRp
--	, isNull(B.OfficePhoneBillDlr,0) As OfficePhoneBillDlr
	, isNull(D.CellPhoneBillRp,0) As CellPhoneBillRp
	, isNull(D.CellPhoneBillDlr,0) As CellPhoneBillDlr
	, isNull(D.CellPhonePrsBillRp,0) As CellPhonePrsBillRp
	, isNull(D.CellPhonePrsBillDlr,0) As CellPhonePrsBillDlr
--	, ISNULL(E.Used,0)
--	, isNull((E.Used)*@ShuttleBusRate*@ExchangeRate,0) As TotalShuttleBillRp
--	, isNull((E.Used)*@ShuttleBusRate,0) As TotalShuttleBillDlr
--	, isNull(C.HomePhoneBillRp,0)+isNull(B.OfficePhoneBillRp,0)+isNull(D.CellPhoneBillRp,0)+isNull((E.Used)*@ShuttleBusRate*@ExchangeRate,0) As TotalBillingRp
	, isNull(D.CellPhoneBillRp,0) As TotalBillingRp
--	, isNull(C.HomePhoneBillDlr,0)+isNull(B.OfficePhoneBillDlr,0)+isNull(D.CellPhoneBillDlr,0)+isNull((E.Used)*@ShuttleBusRate,0) As TotalBillingDlr, 1, GETDATE()
	, isNull(D.CellPhoneBillDlr,0) As TotalBillingDlr, 1, GETDATE()
From vwPhoneCustomerList A
/* 
Left Join
(
Select A.EmpID, A.PhoneNumber As WorkPhone, D.MonthP, D.YearP, Sum(Case when D.isPersonal ='Y' then D.Cost else 0 end) As OfficePhonePrsBillRp, Sum(Case when D.isPersonal ='Y' then Round(isNull(isNull(D.Cost,0)/@ExchangeRate,0),2) else 0 end) As OfficePhonePrsBillDlr
	, Sum(isNull(D.Cost,0)) As OfficePhoneBillRp, Sum(Round(isNull(isNull(D.Cost,0)/@ExchangeRate,0),2)) As OfficePhoneBillDlr 
From BillingDt D
Inner Join MsOfficePhoneNumber A on (D.Extension=A.PhoneNumber)
Where A.BillFlag<>'N' And D.MonthP=@Month And D.YearP=@Year
Group By A.EmpID, A.PhoneNumber, D.MonthP, D.YearP
) B on (A.EmpID=B.EmpID)

Left Join 
(
Select B.EmpID, B.PhoneNumber As HomePhone, A.MonthP, A.YearP, isNull(A.Total,0) As HomePhonePrsBillRp
	, Round(isNull(isNull(A.Total,0)/@ExchangeRate,0),2) As HomePhonePrsBillDlr
	, isNull(A.Total,0) As HomePhoneBillRp, Round(isNull(isNull(A.Total,0)/@ExchangeRate,0),2) As HomePhoneBillDlr
From HomePhone A
Inner Join MsHomePhoneNumber B on (A.Nomor=B.PhoneNumber)
Where B.BillFlag<>'N' And A.MonthP=@Month And A.YearP=@Year
) C on (A.EmpID=C.EmpID)
*/
Left join(
Select B.EmpID, B.PhoneNumber As MobilePhone, A.MonthP, A.YearP
	, isNull(Sum(Case When C.isPersonal='Y' Then C.Cost Else 0 End),0) As CellPhonePrsBillRp
	, Round(isNull(isNull(Sum(Case When C.isPersonal='Y' Then C.Cost Else 0 End),0)/@ExchangeRate,0),2) As CellPhonePrsBillDlr
	, isNull(A.TOTBILLAMOUNT,0) As CellPhoneBillRp, Round(isNull(isNull(A.TOTBILLAMOUNT,0)/@ExchangeRate,0),2) As CellPhoneBillDlr
From CellphoneHd A
Left Join CellphoneDt C on (A.PhoneNumber=C.PhoneNumber And A.MonthP=C.MonthP And A.YearP=C.YearP)
Inner Join MsCellPhoneNumber B on (A.PhoneNumber=B.PhoneNumber)
Where B.BillFlag<>'N' And A.MonthP=@Month And A.YearP=@Year
And B.EmpID in
--temporary exclude employee has more than 1 cellphone number registered
(
select EmpID from MsCellPhoneNumber Where EmpId<>'' And BillFlag<>'N'
--select EmpID from MsCellPhoneNumber Where EmpId<>'' And BillFlag='Y'
Group By EmpID
Having Count(PhoneNumber)=1
)
Group By B.EmpID, B.PhoneNumber, A.MonthP, A.YearP, isNull(A.TOTBILLAMOUNT,0), Round(isNull(isNull(A.TOTBILLAMOUNT,0)/@ExchangeRate,0),2)
) D on (A.EmpID=D.EmpID)
/* 
Left Join
(

	Select EmpID, ISNULL(Used,0) As Used
	From ShuttleBusUsed 
	Where MonthP=@Month and YearP=@Year 
) E on (A.EmpID=E.EmpID)
*/
-- ***** Marin's addin
Left Join
(
Select EmpID, PhoneNumber, BillFlag
From vwCellPhoneNumberList
) F on (A.EmpID=F.EmpID)
-- *****

--Where isNull(A.MobilePhone,'')<>'' And (isNull(C.HomePhonePrsBillRp,0)>0 or isNull(B.OfficePhoneBillRp,0)>0 or isNull(E.Used,0)>0 or isNull(D.CellPhoneBillRp,0)>=0)
Where isNull(A.MobilePhone,'')<>'' And isNull(D.CellPhoneBillRp,0)>=0
And (A.EmpID=@EmpID or @EmpID='')



--Update CellPhone Call duration in second
Update CellPhoneDt set CallDurationSecond=dbo.GetTotDuration(callduration)
Where MonthP=@Month and YearP=@Year

--Update payment status to Awaiting payment for employee who doesn't has open net account
Update MonthlyBilling Set ProgressId=8, ProgressIdDate=GETDATE()
Where (EmpID in 
( Select OwnerID From MsCellPhoneNumber A
  Inner JOin vwPhoneCustomerList B on (A.OwnerID=B.EmpID)  
  Where A.BillFlag<>'N' And len(ISNULL(B.EmailAddress,''))<8
)) 

--Update payment status to complete for billing amount smaller than equal to Ceiling amount parameter
--Update MonthlyBilling Set ProgressId=6 Where TotalBillingRp<=@CeilingAmount And MonthP=@Month and YearP=@Year
Update MonthlyBilling Set ProgressId=7, ProgressIdDate=GETDATE() 
--Where ((HomePhonePrsBillRp+OfficePhonePrsBillRp+CellPhonePrsBillRp+TotalShuttleBillRp)<=@CeilingAmount or (HomePhonePrsBillRp+OfficePhonePrsBillRp+CellPhonePrsBillRp+TotalShuttleBillRp = 0))
Where ((CellPhonePrsBillRp)<=@CeilingAmount or (CellPhonePrsBillRp = 0))
And MonthP=@Month and YearP=@Year And (EmpID=@EmpID or @EmpID='')

--******************
--Update payment status to Awaiting payment for employee who has personal phone   *Marin's fix
Update MonthlyBilling Set ProgressId=4, ProgressIdDate=GETDATE()
Where (EmpID in 
( Select OwnerID From MsCellPhoneNumber A
  Inner JOin vwPhoneCustomerList B on (A.OwnerID=B.EmpID)  
  Where A.BillFlag='P'
))
--******************

--Generate Reconciliation Report

--If @EmpID=''
--Begin
	Exec spGenerateReconRpt @Month, @Year
	
	Exec spGenerateProgressLog @Month, @Year
--End
GO
