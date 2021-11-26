// I cannot use F# optional parameters because they end up forcing you to include the F# dll
// Instead I use overloading with parsimony to achieve similar goals
#light
namespace Excel.FinancialFunctions

open System
open System.Collections
open Excel.FinancialFunctions.Common
open Excel.FinancialFunctions.Tvm
open Excel.FinancialFunctions.Loan
open Excel.FinancialFunctions.Irr
open Excel.FinancialFunctions.Depreciation
open Excel.FinancialFunctions.DayCount
open Excel.FinancialFunctions.Bonds
open Excel.FinancialFunctions.TBill
open Excel.FinancialFunctions.Misc
open Excel.FinancialFunctions.OddBonds

/// A wrapper class to expose the Excel financial functions API to .NET clients
type Financial =
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/accrint-function-fe45d089-6722-4fb3-9379-e1f911d8dc74">ACCRINT function</a>
    /// The accrued interest for a security that pays periodic interest
    static member AccrInt (issue, firstInterest, settlement, rate, par, frequency, basis, calcMethod) =
        calcAccrInt issue firstInterest settlement rate par frequency basis calcMethod

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/accrint-function-fe45d089-6722-4fb3-9379-e1f911d8dc74">ACCRINT function</a>
    /// The accrued interest for a security that pays periodic interest, using "FromIssueToSettlement" calculation method
    static member AccrInt (issue, firstInterest, settlement, rate, par, frequency, basis) = 
        calcAccrInt issue firstInterest settlement rate par frequency basis AccrIntCalcMethod.FromIssueToSettlement
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/accrintm-function-f62f01f9-5754-4cc4-805b-0e70199328a7">ACCRINTM function</a>
    /// The accrued interest for a security that pays interest at maturity
    static member AccrIntM (issue, settlement, rate, par, basis) =
        calcAccrIntM issue settlement rate par basis
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/amordegrc-function-a14d0ca1-64a4-42eb-9b3d-b0dededf9e51">AMORDEGRC function</a>
    /// The depreciation for each accounting period by using a depreciation coefficient
    /// ExcelCompliant is used because Excel stores 13 digits. AmorDegrc algorithm rounds numbers  
    /// and returns different results unless the numbers get rounded to 13 digits before rounding them.  
    /// I.E. 22.49999999999999 is considered 22.5 by Excel, but 22.4 by the .NET framework    
    static member AmorDegrc (cost, datePurchased, firstPeriod, salvage, period, rate, basis, excelCompliant) =
        calcAmorDegrc cost datePurchased firstPeriod salvage period rate basis excelCompliant
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb">AMORLINC function</a>
    /// The depreciation for each accounting period
    static member AmorLinc (cost, datePurchased, firstPeriod, salvage, period, rate, basis) =
        calcAmorLinc cost datePurchased firstPeriod salvage period rate basis

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/coupdaybs-function-eb9a8dfb-2fb2-4c61-8e5d-690b320cf872">COUPDAYBS function</a>
    /// The number of days from the beginning of the coupon period to the settlement date
    static member CoupDaysBS (settlement, maturity, frequency, basis) =
        calcCoupDaysBS settlement maturity frequency basis

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/coupdays-function-cc64380b-315b-4e7b-950c-b30b0a76f671">COUPDAYS function</a>
    /// The number of days in the coupon period that contains the settlement date
    /// The Excel algorithm seems wrong in that it doesn't respect `coupDays = coupDaysBS + coupDaysNC`    
    /// This equality should stand. The differs from Excel by +/- one or two days when the date spans a leap year.
    static member CoupDays (settlement, maturity, frequency, basis) =
        calcCoupDays settlement maturity frequency basis
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/coupdaysnc-function-5ab3f0b2-029f-4a8b-bb65-47d525eea547">COUPDAYSNC function</a>
    /// The number of days from the settlement date to the next coupon date
    static member CoupDaysNC (settlement, maturity, frequency, basis) =
        calcCoupDaysNC settlement maturity frequency basis

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/coupncd-function-fd962fef-506b-4d9d-8590-16df5393691f">COUPNCD function</a>
    /// The next coupon date after the settlement date 
    static member CoupNCD (settlement, maturity, frequency, basis) =
        calcCoupNCD settlement maturity frequency basis

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/coupnum-function-a90af57b-de53-4969-9c99-dd6139db2522">COUPNUM function</a>
    /// The number of coupons payable between the settlement date and maturity date
    static member CoupNum (settlement, maturity, frequency, basis) =
        calcCoupNum settlement maturity frequency basis
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/couppcd-function-2eb50473-6ee9-4052-a206-77a9a385d5b3">COUPCD function</a>
    /// The previous coupon date before the settlement date
    static member CoupPCD (settlement, maturity, frequency, basis) =
        calcCoupPCD settlement maturity frequency basis

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/cumipmt-function-61067bb0-9016-427d-b95b-1a752af0e606">CUMIPMT function</a>
    /// The cumulative interest paid between two periods 
    static member CumIPmt (rate, nper, pv, startPeriod, endPeriod, typ) =
        calcCumipmt rate nper pv startPeriod endPeriod typ

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/cumprinc-function-94a4516d-bd65-41a1-bc16-053a6af4c04d">CUMPRINC function</a>
    /// The cumulative principal paid on a loan between two periods 
    static member CumPrinc (rate, nper, pv, startPeriod, endPeriod, typ) =
        calcCumprinc rate nper pv startPeriod endPeriod typ

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/db-function-354e7d28-5f93-4ff1-8a52-eb4ee549d9d7">DB function</a>
    /// The depreciation of an asset for a specified period by using the fixed-declining balance method
    static member Db (cost, salvage, life, period, month) =
        calcDb cost salvage life period month

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/db-function-354e7d28-5f93-4ff1-8a52-eb4ee549d9d7">DB function</a>
    /// The depreciation of an asset for a specified period by using the fixed-declining balance method
    static member Db (cost, salvage, life, period) =
        calcDb cost salvage life period 12.

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/ddb-function-519a7a37-8772-4c96-85c0-ed2c209717a5">DDB function</a>
    /// The depreciation of an asset for a specified period by using the double-declining balance method or some other method that you specify
    /// 
    /// Excel Ddb has two interesting characteristics:  
    /// 1. It special cases ddb for fractional periods between 0 and 1 by considering them to be 1  
    /// 2. It is inconsistent with VDB(..., True) for fractional periods, even if VDB(..., True) is defined to be the same as ddb. The algorithm for VDB is theoretically correct.  
    /// This function makes the same 1. adjustment.
    static member Ddb (cost, salvage, life, period, factor) =
        calcDdb cost salvage life period factor

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/ddb-function-519a7a37-8772-4c96-85c0-ed2c209717a5">DDB function</a>
    /// The depreciation of an asset for a specified period by using the double-declining balance method or some other method that you specify
    static member Ddb (cost, salvage, life, period) =
        calcDdb cost salvage life period 2.

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/disc-function-71fce9f3-3f05-4acf-a5a3-eac6ef4daa53">DISC function</a>
    /// The discount rate for a security 
    static member Disc (settlement, maturity, pr, redemption, basis) =
        calcDisc settlement maturity pr redemption basis

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/dollarde-function-db85aab0-1677-428a-9dfd-a38476693427">DOLLARDE function</a>
    /// Converts a dollar price, expressed as a fraction, into a dollar price, expressed as a decimal number 
    static member DollarDe (fractionalDollar, fraction) =
        calcDollarDe fractionalDollar fraction
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/dollarfr-function-0835d163-3023-4a33-9824-3042c5d4f495">DOLLARFR function</a>
    /// Converts a dollar price, expressed as a decimal number, into a dollar price, expressed as a fraction 
    static member DollarFr (decimalDollar, fraction) =
        calcDollarFr decimalDollar fraction
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/duration-function-b254ea57-eadc-4602-a86a-c8e369334038">DURATION function</a>
    /// The annual duration of a security with periodic interest payments 
    static member Duration (settlement, maturity, coupon, yld, frequency, basis) =
        calcDuration settlement maturity coupon yld frequency basis
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/effect-function-910d4e4c-79e2-4009-95e6-507e04f11bc4">EFFECT function</a>
    /// The effective annual interest rate
    static member Effect (nominalRate, npery) =
        calcEffect nominalRate npery

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/fv-function-2eef9f44-a084-4c61-bdd8-4fe4bb1b71b3">FV function</a>
    /// The future value of an investment
    static member Fv (rate, nper, pmt, pv, typ) =
        calcFv rate nper pmt pv typ

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/fvschedule-function-bec29522-bd87-4082-bab9-a241f3fb251d">FVSCHEDULE function</a>
    /// The future value of an initial principal after applying a series of compound interest rates 
    static member FvSchedule (principal, schedule) =
        calcFvSchedule principal schedule
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/intrate-function-5cb34dde-a221-4cb6-b3eb-0b9e55e1316f">INTRATE function</a>
    /// The interest rate for a fully invested security 
    static member IntRate (settlement, maturity, investment, redemption, basis) =
        calcIntRate settlement maturity investment redemption basis

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/ipmt-function-5cce0ad6-8402-4a41-8d29-61a0b054cb6f">IPMT function</a>
    /// The interest payment for an investment for a given period 
    static member IPmt (rate, per, nper, pv, fv, typ) =
        calcIpmt rate per nper pv fv typ
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/irr-function-64925eaa-9988-495b-b290-3ad0c163c1bc">IRR function</a>
    /// The internal rate of return for a series of cash flows 
    static member Irr (values, guess) =
        calcIrr values guess 

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/irr-function-64925eaa-9988-495b-b290-3ad0c163c1bc">IRR function</a>
    /// The internal rate of return for a series of cash flows 
    static member Irr (values) =
        calcIrr values 0.1 
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/ispmt-function-fa58adb6-9d39-4ce0-8f43-75399cea56cc">ISPMT function</a>
    /// Calculates the interest paid during a specific period of an investment
    static member ISPmt (rate, per, nper, pv) =
        calcIspmt rate per nper pv
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/mduration-function-b3786a69-4f20-469a-94ad-33e5b90a763c">MDURATION function</a>
    /// The Macauley modified duration for a security with an assumed par value of $100 
    static member MDuration (settlement, maturity, coupon, yld, frequency, basis) =
        calcMDuration settlement maturity coupon yld frequency basis
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/mirr-function-b020f038-7492-4fb4-93c1-35c345b53524">MIRR function</a>
    /// The internal rate of return where positive and negative cash flows are financed at different rates 
    static member Mirr (values, financeRate, reinvestRate) =
        calcMirr values financeRate reinvestRate 
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/nominal-function-7f1ae29b-6b92-435e-b950-ad8b190ddd2b">NOMINAL function</a>
    /// The annual nominal interest rate 
    static member Nominal (effectRate, npery) =
        calcNominal effectRate npery

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/nper-function-240535b5-6653-4d2d-bfcf-b6a38151d815">NPER function</a>
    /// The number of periods for an investment 
    static member NPer (rate, pmt, pv, fv, typ) =
        calcNper rate pmt pv fv typ

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/npv-function-8672cb67-2576-4d07-b67b-ac28acf2a568">NPV function</a>
    /// The net present value of an investment based on a series of periodic cash flows and a discount rate 
    static member Npv (rate, values) =
        calcNpv rate values
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/oddfprice-function-d7d664a8-34df-4233-8d2b-922bcf6a69e1">ODDFPRICE function</a>
    /// The price per $100 face value of a security with an odd first period 
    static member OddFPrice (settlement, maturity, issue, firstCoupon, rate, yld, redemption, frequency, basis) =
        calcOddFPrice settlement maturity issue firstCoupon rate yld redemption frequency basis
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/oddfyield-function-66bc8b7b-6501-4c93-9ce3-2fd16220fe37">ODDFYIELD function</a>
    /// The yield of a security with an odd first period
    static member OddFYield (settlement, maturity, issue, firstCoupon, rate, pr, redemption, frequency, basis) =
        calcOddFYield settlement maturity issue firstCoupon rate pr redemption frequency basis

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/oddlprice-function-fb657749-d200-4902-afaf-ed5445027fc4">ODDLPRICE function</a>
    /// The price per $100 face value of a security with an odd last period
    static member OddLPrice (settlement, maturity, lastInterest, rate, yld, redemption, frequency, basis) =
        calcOddLPrice settlement maturity lastInterest rate yld redemption frequency basis
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/oddlyield-function-c873d088-cf40-435f-8d41-c8232fee9238">ODDLYIELD function</a>
    /// The yield of a security with an odd last period
    static member OddLYield (settlement, maturity, lastInterest, rate, pr, redemption, frequency, basis) =
        calcOddLYield settlement maturity lastInterest rate pr redemption frequency basis

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/pmt-function-0214da64-9a63-4996-bc20-214433fa6441">PMT function</a>
    /// The periodic payment for an annuity
    static member Pmt (rate, nper, pv, fv, typ) =
        calcPmt rate nper pv fv typ

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/ppmt-function-c370d9e3-7749-4ca4-beea-b06c6ac95e1b">PPMT function</a>
    /// The payment on the principal for an investment for a given period 
    static member PPmt (rate, per, nper, pv, fv, typ) =
        calcPpmt rate per nper pv fv typ

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/price-function-3ea9deac-8dfa-436f-a7c8-17ea02c21b0a">PRICE function</a>
    /// The price per $100 face value of a security that pays periodic interest 
    static member Price (settlement, maturity, rate, yld, redemption, frequency, basis) =
        calcPrice settlement maturity rate yld redemption frequency basis
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/pricedisc-function-d06ad7c1-380e-4be7-9fd9-75e3079acfd3">PRICEDISC function</a>
    /// The price per $100 face value of a discounted security 
    static member PriceDisc (settlement, maturity, discount, redemption, basis) =
        calcPriceDisc settlement maturity discount redemption basis
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/pricemat-function-52c3b4da-bc7e-476a-989f-a95f675cae77">PRICEMAT function</a>
    /// The price per $100 face value of a security that pays interest at maturity 
    static member PriceMat (settlement, maturity, issue, rate, yld, basis) =
        calcPriceMat settlement maturity issue rate yld basis
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/pv-function-23879d31-0e02-4321-be01-da16e8168cbd">PV function</a>
    /// The present value of an investment 
    static member Pv (rate, nper, pmt, fv, typ) =
        calcPv rate nper pmt fv typ

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/rate-function-9f665657-4a7e-4bb7-a030-83fc59e748ce">RATE function</a>
    /// The interest rate per period of an annuity 
    static member Rate (nper, pmt, pv, fv, typ, guess) =
        calcRate nper pmt pv fv typ guess

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/rate-function-9f665657-4a7e-4bb7-a030-83fc59e748ce">RATE function</a>
    /// The interest rate per period of an annuity 
    static member Rate (nper, pmt, pv, fv, typ) =
        calcRate nper pmt pv fv typ 0.1

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/received-function-7a3f8b93-6611-4f81-8576-828312c9b5e5">RECEIVED function</a>
    /// The amount received at maturity for a fully invested security
    static member Received (settlement, maturity, investment, discount,basis) =
        calcReceived settlement maturity investment discount basis
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/sln-function-cdb666e5-c1c6-40a7-806a-e695edc2f1c8">SLN function</a>
    /// The straight-line depreciation of an asset for one period
    static member Sln (cost, salvage, life) =
        calcSln cost salvage life

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/syd-function-069f8106-b60b-4ca2-98e0-2a0f206bdb27">SYD function</a>
    /// The sum-of-years' digits depreciation of an asset for a specified period 
    static member Syd (cost, salvage, life, per) =
        calcSyd cost salvage life per
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/tbilleq-function-2ab72d90-9b4d-4efe-9fc2-0f81f2c19c8c">TBILLEQ function</a>
    /// The bond-equivalent yield for a Treasury bill 
    static member TBillEq (settlement, maturity, discount) =
        calcTBillEq settlement maturity discount

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/tbillprice-function-eacca992-c29d-425a-9eb8-0513fe6035a2">TBILLPRICE function</a>
    /// The price per $100 face value for a Treasury bill
    static member TBillPrice (settlement, maturity, discount) =
        calcTBillPrice settlement maturity discount

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/tbillyield-function-6d381232-f4b0-4cd5-8e97-45b9c03468ba">TBILLYIELD function</a>
    /// The yield for a Treasury bill
    static member TBillYield (settlement, maturity, pr) =
        calcTBillYield settlement maturity pr

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/vdb-function-dde4e207-f3fa-488d-91d2-66d55e861d73">VDB function</a>
    /// The depreciation of an asset for a specified or partial period by using a declining balance method.
    ///
    /// In the excel version of this algorithm the depreciation in the period (0,1) is not the same as the sum of the depreciations in periods (0,0.5) (0.5,1)  
    /// `VDB(100,10,13,0,0.5,1,0) + VDB(100,10,13,0.5,1,1,0) != VDB(100,10,13,0,1,1,0)`  
    /// Notice that in Excel by using '1' (no_switch) instead of '0' as the last parameter everything works as expected.  
    /// In truth, the last parameter should have no influence in the calculation given that in the first period there is no switch to sln depreciation.  
    /// Overall, I think my algorithm is correct, even if it disagrees with Excel when startperiod is fractional.
    static member Vdb (cost, salvage, life, startPeriod, endPeriod, factor, noSwitch) =
        calcVdb cost salvage life startPeriod endPeriod factor noSwitch

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/vdb-function-dde4e207-f3fa-488d-91d2-66d55e861d73">VDB function</a>
    /// The depreciation of an asset for a specified or partial period by using a declining balance method
    static member Vdb (cost, salvage, life, startPeriod, endPeriod, factor) =
        calcVdb cost salvage life startPeriod endPeriod factor VdbSwitch.SwitchToStraightLine

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/vdb-function-dde4e207-f3fa-488d-91d2-66d55e861d73">VDB function</a>
    /// The depreciation of an asset for a specified or partial period by using a declining balance method
    static member Vdb (cost, salvage, life, startPeriod, endPeriod) =
        calcVdb cost salvage life startPeriod endPeriod 2. VdbSwitch.SwitchToStraightLine
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/xirr-function-de1242ec-6477-445b-b11b-a303ad9adc9d">XIRR function</a>
    /// The internal rate of return for a schedule of cash flows that is not necessarily periodic
    static member XIrr (values, dates, guess) =
        calcXirr values dates guess

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/xirr-function-de1242ec-6477-445b-b11b-a303ad9adc9d">XIRR function</a>
    /// The internal rate of return for a schedule of cash flows that is not necessarily periodic
    static member XIrr (values, dates) =
        calcXirr values dates 0.1
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/xnpv-function-1b42bbf6-370f-4532-a0eb-d67c16b664b7">XNPV function</a>
    /// The net present value for a schedule of cash flows that is not necessarily periodic
    static member XNpv (rate, values, dates) =
        calcXnpv rate values dates
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/yield-function-f5f5ca43-c4bd-434f-8bd2-ed3c9727a4fe">YIELD function</a>
    /// The yield on a security that pays periodic interest
    static member Yield (settlement, maturity, rate, pr, redemption, frequency, basis) =
        calcYield settlement maturity rate pr redemption frequency basis

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/yielddisc-function-a9dbdbae-7dae-46de-b995-615faffaaed7">YIELDDISC function</a>
    /// The annual yield for a discounted security; for example, a Treasury bill
    static member YieldDisc (settlement, maturity, pr, redemption, basis) =
        calcYieldDisc settlement maturity pr redemption basis
    
    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/yieldmat-function-ba7d1809-0d33-4bcb-96c7-6c56ec62ef6f">YIELDMAT function</a>
    /// The annual yield of a security that pays interest at maturity
    static member YieldMat (settlement, maturity, issue, rate, pr, basis) =
        calcYieldMat settlement maturity issue rate pr basis

    /// <a target="_blank" href="https://support.microsoft.com/en-us/office/yearfrac-function-3844141e-c76d-4143-82b6-208454ddc6a8">YEARFRAC function</a>
    /// Calculates the fraction of the year represented by the number of whole days between two dates - not a financial function
    static member YearFrac (startDate, endDate, basis) =
        calcYearFrac startDate endDate basis    
           