// This is a wrapper class to expose the API nicely to .NET clients
// I cannot use F# optional parameters because they end up forcing you to include the F# dll
// Instead I use overloading with parsimony to achieve similar goals
#light
namespace System.Numeric

open System
open System.Collections
open System.Numeric
open System.Numeric.Common
open System.Numeric.Tvm
open System.Numeric.Loan
open System.Numeric.Irr
open System.Numeric.Depreciation
open System.Numeric.DayCount
open System.Numeric.Bonds
open System.Numeric.TBill
open System.Numeric.Misc
open System.Numeric.OddBonds
open System.Numeric

type Financial =
    /// http://office.microsoft.com/en-us/excel/HP052089791033.aspx
    static member AccrInt (issue, firstInterest, settlement, rate, par, frequency, basis, calcMethod) =
        calcAccrInt issue firstInterest settlement rate par frequency basis calcMethod
    static member AccrInt (issue, firstInterest, settlement, rate, par, frequency, basis) = 
        calcAccrInt issue firstInterest settlement rate par frequency basis AccrIntCalcMethod.FromIssueToSettlement
    /// http://office.microsoft.com/en-us/excel/HP052089801033.aspx
    static member AccrIntM (issue, settlement, rate, par, basis) =
        calcAccrIntM issue settlement rate par basis
    /// http://office.microsoft.com/en-us/excel/HP052089841033.aspx
    /// ExcelCompaint is used because Excel stores 13 digits. amorDegrc alorithm rounds numbers
    /// and returns different results unless the numbers get rounded to 13 digits before rounding them.
    /// I.E. 22.49999999999999 is considered 22.5 by Excel, but 22.4 by the .NET framework    
    static member AmorDegrc (cost, datePurchased, firstPeriod, salvage, period, rate, basis, excelCompliant) =
        calcAmorDegrc cost datePurchased firstPeriod salvage period rate basis excelCompliant
    /// http://office.microsoft.com/en-us/excel/HP052089851033.aspx
    static member AmorLinc (cost, datePurchased, firstPeriod, salvage, period, rate, basis) =
        calcAmorLinc cost datePurchased firstPeriod salvage period rate basis
    /// http://office.microsoft.com/en-us/excel/HP052090301033.aspx
    static member CoupDaysBS (settlement, maturity, frequency, basis) =
        calcCoupDaysBS settlement maturity frequency basis
    /// http://office.microsoft.com/en-us/excel/HP052090311033.aspx
    /// The Excel algorithm seems wrong in that it doesn't respect coupDays = coupDaysBS + coupDaysNC
    /// This equality should stand. By manually counting the days, I'm pretty confident that my algorithm is correct
    /// My result differs from Excel by +/- one or two days when the date spans a leap year.
    static member CoupDays (settlement, maturity, frequency, basis) =
        calcCoupDays settlement maturity frequency basis
    /// http://office.microsoft.com/en-us/excel/HP052090321033.aspx
    static member CoupDaysNC (settlement, maturity, frequency, basis) =
        calcCoupDaysNC settlement maturity frequency basis
    /// http://office.microsoft.com/en-us/excel/HP052090331033.aspx
    static member CoupNCD (settlement, maturity, frequency, basis) =
        calcCoupNCD settlement maturity frequency basis
    /// http://office.microsoft.com/en-us/excel/HP052090341033.aspx
    static member CoupNum (settlement, maturity, frequency, basis) =
        calcCoupNum settlement maturity frequency basis
    /// http://office.microsoft.com/en-us/excel/HP052090351033.aspx
    static member CoupPCD (settlement, maturity, frequency, basis) =
        calcCoupPCD settlement maturity frequency basis
    /// http://office.microsoft.com/en-us/excel/HP052090381033.aspx
    static member CumIPmt (rate, nper, pv, startPeriod, endPeriod, typ) =
        calcCumipmt rate nper pv startPeriod endPeriod typ
    /// http://office.microsoft.com/en-us/excel/HP052090391033.aspx
    static member CumPrinc (rate, nper, pv, startPeriod, endPeriod, typ) =
        calcCumprinc rate nper pv startPeriod endPeriod typ
    /// http://office.microsoft.com/en-us/excel/HP052090481033.aspx
    static member Db (cost, salvage, life, period, month) =
        calcDb cost salvage life period month
    static member Db (cost, salvage, life, period) =
        calcDb cost salvage life period 12.
    /// http://office.microsoft.com/en-us/excel/HP052090511033.aspx
    /// Excel Ddb has two interesting characteristics:
    /// 1. It special cases ddb for fractional periods between 0 and 1 by considering them to be 1
    /// 2. It is inconsistent with VDB(..., True) for fractional periods, even if VDB(..., True) is defined to be the same as ddb. The algorithm for VDB is theoretically correct.
    /// This function makes the same 1. ajdjustment.
    static member Ddb (cost, salvage, life, period, factor) =
        calcDdb cost salvage life period factor
    static member Ddb (cost, salvage, life, period) =
        calcDdb cost salvage life period 2.
    /// http://office.microsoft.com/en-us/excel/HP052090601033.aspx
    static member Disc (settlement, maturity, pr, redemption, basis) =
        calcDisc settlement maturity pr redemption basis
    /// http://office.microsoft.com/en-us/excel/HP052090641033.aspx
    static member DollarDe (fractionalDollar, fraction) =
        calcDollarDe fractionalDollar fraction
    /// http://office.microsoft.com/en-us/excel/HP052090651033.aspx
    static member DollarFr (decimalDollar, fraction) =
        calcDollarFr decimalDollar fraction
    /// http://office.microsoft.com/en-us/excel/HP052090701033.aspx
    static member Duration (settlement, maturity, coupon, yld, frequency, basis) =
        calcDuration settlement maturity coupon yld frequency basis
    /// http://office.microsoft.com/en-us/excel/HP052090741033.aspx
    static member Effect (nominalRate, npery) =
        calcEffect nominalRate npery
    /// http://office.microsoft.com/en-us/excel/HP052090991033.aspx
    static member Fv (rate, nper, pmt, pv, typ) =
        calcFv rate nper pmt pv typ
    /// http://office.microsoft.com/en-us/excel/HP052091001033.aspx
    static member FvSchedule (principal, schedule) =
        calcFvSchedule principal schedule
    /// http://office.microsoft.com/en-us/excel/HP052091441033.aspx
    static member IntRate (settlement, maturity, investment, redemption, basis) =
        calcIntRate settlement maturity investment redemption basis
    /// http://office.microsoft.com/en-us/excel/HP052091451033.aspx
    static member IPmt (rate, per, nper, pv, fv, typ) =
        calcIpmt rate per nper pv fv typ
    /// http://office.microsoft.com/en-us/excel/HP052091461033.aspx
    static member Irr (values, guess) =
        calcIrr values guess 
    static member Irr (values) =
        calcIrr values 0.1 
    /// http://office.microsoft.com/en-us/excel/HP052508401033.aspx
    static member ISPmt (rate, per, nper, pv) =
        calcIspmt rate per nper pv
    /// http://office.microsoft.com/en-us/excel/HP052091731033.aspx
    static member MDuration (settlement, maturity, coupon, yld, frequency, basis) =
        calcMDuration settlement maturity coupon yld frequency basis
    /// http://office.microsoft.com/en-us/excel/HP052091801033.aspx
    static member Mirr (values, financeRate, reinvestRate) =
        calcMirr values financeRate reinvestRate 
    /// http://office.microsoft.com/en-us/excel/HP052091911033.aspx
    static member Nominal (effectRate, npery) =
        calcNominal effectRate npery
    /// http://office.microsoft.com/en-us/excel/HP052091981033.aspx       
    static member NPer (rate, pmt, pv, fv, typ) =
        calcNper rate pmt pv fv typ
    /// http://office.microsoft.com/en-us/excel/HP052091991033.aspx
    static member Npv (rate, values) =
        calcNpv rate values
    /// http://office.microsoft.com/en-us/excel/HP052092041033.aspx
    static member OddFPrice (settlement, maturity, issue, firstCoupon, rate, yld, redemption, frequency, basis) =
        calcOddFPrice settlement maturity issue firstCoupon rate yld redemption frequency basis
    /// http://office.microsoft.com/en-us/excel/HP052092051033.aspx
    static member OddFYield (settlement, maturity, issue, firstCoupon, rate, pr, redemption, frequency, basis) =
        calcOddFYield settlement maturity issue firstCoupon rate pr redemption frequency basis
    /// http://office.microsoft.com/en-us/excel/HP052092061033.aspx
    static member OddLPrice (settlement, maturity, lastInterest, rate, yld, redemption, frequency, basis) =
        calcOddLPrice settlement maturity lastInterest rate yld redemption frequency basis
    /// http://office.microsoft.com/en-us/excel/HP052092071033.aspx
    static member OddLYield (settlement, maturity, lastInterest, rate, pr, redemption, frequency, basis) =
        calcOddLYield settlement maturity lastInterest rate pr redemption frequency basis
    /// http://office.microsoft.com/en-us/excel/HP052092151033.aspx
    static member Pmt (rate, nper, pv, fv, typ) =
        calcPmt rate nper pv fv typ
    /// http://office.microsoft.com/en-us/excel/HP052092181033.aspx
    static member PPmt (rate, per, nper, pv, fv, typ) =
        calcPpmt rate per nper pv fv typ
    /// http://office.microsoft.com/en-us/excel/HP052092191033.aspx
    static member Price (settlement, maturity, rate, yld, redemption, frequency, basis) =
        calcPrice settlement maturity rate yld redemption frequency basis
    /// http://office.microsoft.com/en-us/excel/HP052092201033.aspx
    static member PriceDisc (settlement, maturity, discount, redemption, basis) =
        calcPriceDisc settlement maturity discount redemption basis
    /// http://office.microsoft.com/en-us/excel/HP052092211033.aspx
    static member PriceMat (settlement, maturity, issue, rate, yld, basis) =
        calcPriceMat settlement maturity issue rate yld basis
    /// http://office.microsoft.com/en-us/excel/HP052092251033.aspx    
    static member Pv (rate, nper, pmt, fv, typ) =
        calcPv rate nper pmt fv typ
    /// http://office.microsoft.com/en-us/excel/HP052092321033.aspx        
    static member Rate (nper, pmt, pv, fv, typ, guess) =
        calcRate nper pmt pv fv typ guess
    static member Rate (nper, pmt, pv, fv, typ) =
        calcRate nper pmt pv fv typ 0.1
    /// http://office.microsoft.com/en-us/excel/HP052092331033.aspx
    static member Received (settlement, maturity, investment, discount,basis) =
        calcReceived settlement maturity investment discount basis
    /// http://office.microsoft.com/en-us/excel/HP052092631033.aspx
    static member Sln (cost, salvage, life) =
        calcSln cost salvage life
    /// http://office.microsoft.com/en-us/excel/HP052093021033.aspx
    static member Syd (cost, salvage, life, per) =
        calcSyd cost salvage life per
    /// http://office.microsoft.com/en-us/excel/HP052093091033.aspx
    static member TBillEq (settlement, maturity, discount) =
        calcTBillEq settlement maturity discount
    /// http://office.microsoft.com/en-us/excel/HP052093101033.aspx
    static member TBillPrice (settlement, maturity, discount) =
        calcTBillPrice settlement maturity discount
    /// http://office.microsoft.com/en-us/excel/HP052093111033.aspx
    static member TBillYield (settlement, maturity, pr) =
        calcTBillYield settlement maturity pr
    /// http://office.microsoft.com/en-us/excel/HP052093341033.aspx
    /// In the excel version of this algorithm the depreciation in the period (0,1) is not the same as the sum of the depreciations in periods (0,0.5) (0.5,1)
    /// VDB(100,10,13,0,0.5,1,0) + VDB(100,10,13,0.5,1,1,0) <> VDB(100,10,13,0,1,1,0)
    /// Notice that in Excel by using '1' (no_switch) instead of '0' as the last parameter everything works as expected.
    /// In truth, the last parameter should have no influence in the calculation given that in the first period there is no switch to sln depreciation.
    /// Overall, I think my algorithm is correct, even if it disagrees with Excel when startperiod is fractional.
    static member Vdb (cost, salvage, life, startPeriod, endPeriod, factor, noSwitch) =
        calcVdb cost salvage life startPeriod endPeriod factor noSwitch
    static member Vdb (cost, salvage, life, startPeriod, endPeriod, factor) =
        calcVdb cost salvage life startPeriod endPeriod factor VdbSwitch.SwitchToStraightLine
    static member Vdb (cost, salvage, life, startPeriod, endPeriod) =
        calcVdb cost salvage life startPeriod endPeriod 2. VdbSwitch.SwitchToStraightLine
    /// http://office.microsoft.com/en-us/excel/HP052093411033.aspx
    static member XIrr (values, dates, guess) =
        calcXirr values dates guess
    static member XIrr (values, dates) =
        calcXirr values dates 0.1
    /// http://office.microsoft.com/en-us/excel/HP052093421033.aspx
    static member XNpv (rate, values, dates) =
        calcXnpv rate values dates
    /// http://office.microsoft.com/en-us/excel/HP052093451033.aspx
    static member Yield (settlement, maturity, rate, pr, redemption, frequency, basis) =
        calcYield settlement maturity rate pr redemption frequency basis
    /// http://office.microsoft.com/en-us/excel/HP052093461033.aspx
    static member YieldDisc (settlement, maturity, pr, redemption, basis) =
        calcYieldDisc settlement maturity pr redemption basis
    /// http://office.microsoft.com/en-us/excel/HP052093471033.aspx
    static member YieldMat (settlement, maturity, issue, rate, pr, basis) =
        calcYieldMat settlement maturity issue rate pr basis
    // Not financial function (bonus material)
    /// http://office.microsoft.com/en-us/excel/HP052093441033.aspx
    static member YearFrac (startDate, endDate, basis) =
        calcYearFrac startDate endDate basis    
                         
                