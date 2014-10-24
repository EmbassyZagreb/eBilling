/****** Object:  StoredProcedure [dbo].[spGenerateReconRpt]    Script Date: 08/01/2014 13:31:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[spGenerateReconRpt]
		@MonthP varchar(2)=''
		,@YearP varchar(4)=''
As
Delete Reconciliation Where MonthP=@MonthP And YearP=@YearP

--Unknown number
Insert Reconciliation(MonthP,YearP,ProgressID,PhoneNumber,Balance,CreateDate)
Select A.MonthP, A.YearP, 11, A.PhoneNumber, isNULL(A.TOTBILLAMOUNT,0), GETDATE()
From CellPhoneHd A
Left Join vwPhoneCustomerList B on (A.PhoneNumber=B.MobilePhone)
Where A.MonthP=@MonthP And A.YearP=@YearP
And PhoneNumber not in (Select PhoneNumber from MsCellPhoneNumber)

/*
--Duplicate names
Insert Reconciliation(MonthP,YearP,ProgressID,PhoneNumber, EmpID, Balance,CreateDate)
Select @MonthP, @YearP, 12, PhoneNumber, EmpID, 0, GETDATE()
From vwCellPhoneNumberList
Where BillFlag='Y' and  EmpID in (
Select EmpID from vwPhoneCustomerList Where LEN(EmpID)>1 and  ISNULL(EmpName,'') <>'' and isNull(MobilePhone,'')<>''
Group By EmpID
Having COUNT(EmpID)>1
)
*/

--Null or vacant assignment
Insert Reconciliation(MonthP,YearP,ProgressID,PhoneNumber, EmpID, Balance,CreateDate)
Select @MonthP, @YearP, 13, A.PhoneNumber, B.EmpID, 0, GETDATE()
From CellPhoneHd A
Inner Join MsCellPhoneNumber B on (A.PHONENUMBER=B.PhoneNumber)
Where B.EmpID not in (Select EmpID from vwPhoneCustomerList Where LEN(EmpID)>1 and  ISNULL(EmpName,'') <>'')
And A.MonthP=@MonthP And A.YearP=@YearP

--Empty Login
Insert Reconciliation(MonthP,YearP,ProgressID,PhoneNumber, EmpID, Balance,CreateDate)
Select MonthP, YearP, 14, MobilePhone, EmpID, TotalBillingRp, GETDATE()
From vwMonthlyBilling 
Where EmpID in (Select EmpID from vwPhoneCustomerList  Where loginid='' and right(EmailAddress,9)='state.gov'  and len(EmailAddress)>5 and RIGHT(EmpId,1)<>'N' And Status='C')
And MonthP=@MonthP And YearP=@YearP And MobilePhone is not null

--Wrong Bill Charged
/*
Insert Reconciliation(MonthP,YearP,ErrorType,PhoneNumber, EmpID, Balance,CreateDate)
Select MonthP, YearP, 'Wrong Bill Charged', MobilePhone, EmpID, TotalBillingAmountPrsRp, GETDATE()
From vwMonthlyBilling 
Where MobilePhone not in (Select PHONENUMBER from CellPhoneHd Where MonthP=@MonthP And YearP=@YearP)
And MonthP=@MonthP And YearP=@YearP And MobilePhone is not null


--Non OpenNet
Insert Reconciliation(MonthP,YearP,ErrorType,PhoneNumber, EmpID, Balance,CreateDate)
Select MonthP, YearP, 'Non OpenNet', MobilePhone, EmpID, TotalBillingAmountPrsRp, GETDATE()
From vwMonthlyBilling 
Where EmpID in (Select EmpID from vwPhoneCustomerList  Where loginid='' and right(EmailAddress,9)<>'state.gov' and len(EmailAddress)>5 and RIGHT(EmpId,1)<>'N' And Status='C')
And MonthP=@MonthP And YearP=@YearP And MobilePhone is not null
*/

--Zero Amount
Insert Reconciliation(MonthP,YearP,ProgressID,PhoneNumber,Balance,CreateDate)
Select A.MonthP, A.YearP, 15, A.PhoneNumber, isNULL(A.TOTBILLAMOUNT,0), GETDATE()
From CellPhoneHd A
Inner Join MsCellPhoneNumber B on (A.PHONENUMBER=B.PhoneNumber)
Where A.MonthP=@MonthP And A.YearP=@YearP And A.TOTBILLAMOUNT=0

-- Number not set as Discontinued
Insert Reconciliation(MonthP,YearP,ProgressID,PhoneNumber, EmpID, Balance,CreateDate)
Select @MonthP, @YearP, 16, PhoneNumber, EmpID, 0, GETDATE()
From vwCellPhoneNumberList
Where PhoneNumber not in (Select PhoneNumber From CellPhoneHd Where MonthP=@MonthP And YearP=@YearP) and Discontinued='N'

-- Bill not Generated
Insert Reconciliation(MonthP,YearP,ProgressID,PhoneNumber, EmpID, Balance,CreateDate)
Select A.MonthP, A.YearP, 17, A.PhoneNumber, B.EmpID, isNULL(A.TOTBILLAMOUNT,0), GETDATE()
From CellPhoneHd A
Left Join vwPhoneCustomerList B on (A.PhoneNumber=B.MobilePhone)
Where A.MonthP=@MonthP And A.YearP=@YearP
And PhoneNumber not in (SELECT PhoneNumber FROM MonthlyBilling WHERE MonthP=@MonthP AND YearP=@YearP and PhoneNumber is not null)
GO
