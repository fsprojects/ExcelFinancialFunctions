// Wrappers for Excel functions to be used in the testcases
#light
namespace Excel.FinancialFunctions
open Microsoft.Office.Interop.Excel // You need Excel 12 to test this

open System
open System.Runtime.InteropServices // For COMException
open Excel.FinancialFunctions.Common
open Excel.FinancialFunctions.Tvm
open Excel.FinancialFunctions.DayCount

module internal ExcelTesting =
    // Need to be a singleton for perf reasons
    let app = new ApplicationClass()
    let funcs = app.WorksheetFunction
    
    let pvEx r nper pmt fv (pd:PaymentDue) = funcs.Pv(r, nper, pmt, fv, int pd) 
    let fvEx r nper pmt thePv (pd:PaymentDue) = funcs.Fv(r, nper, pmt, thePv, int pd)
    let pmtEx r nper pv fv (pd:PaymentDue) = funcs.Pmt(r, nper, pv, fv, int pd)
    let ipmtEx r per nper pv fv (pd:PaymentDue) = funcs.Ipmt(r, per, nper, pv, fv, int pd)
    let ppmtEx r per nper pv fv (pd:PaymentDue) = funcs.Ppmt(r, per, nper, pv, fv, int pd)
    let nperEx r pmt pv fv (pd:PaymentDue) = float (funcs.NPer(r, pmt, pv, fv, float pd))
    let rateEx nper pmt pv fv (pd:PaymentDue) guess = funcs.Rate (nper, pmt, pv, fv, float pd, guess)
    let fvScheduleEx pv interests = funcs.FVSchedule(pv, interests)
    let irrEx cfs guess = funcs.Irr(cfs, guess)
    let npvEx r cfs = funcs.Npv(r, cfs)
    let mirrEx cfs financeRate reinvestRate = funcs.MIrr(cfs, financeRate, reinvestRate)
    // Is there a bug in the Excel managed OM in that Worksheet.Xnpv takes 2 params instead of three??
    // Trying to make it work with app.Run, but giving up after a day. XIRR should provide enough test
    //let convertDatesToLong (dates:array<DateTime>) = dates |> Array.map (fun x -> x.ToOADate())
    //let xnpvEx r cfs dates: float = Convert.ToDouble (app.Run(Macro = "atpvbaen.xlam!XNPV", Arg1 = r, Arg2 = cfs, Arg3 = (convertDatesToLong dates) ))
    let xirrEx cfs dates guess = funcs.Xirr(cfs, dates, guess)
    let dbEx cost salvage life period month = funcs.Db(cost, salvage, life, period, month)
    let slnEx cost salvage life = funcs.Sln(cost, salvage, life)
    let sydEx cost salvage life period = funcs.Syd(cost, salvage, life, period)
    let ddbEx cost salvage life period factor = funcs.Ddb(cost, salvage, life, period, factor)
    let vdbEx cost salvage life startPeriod endPeriod factor bflag = funcs.Vdb(cost, salvage, life, startPeriod, endPeriod, factor, bflag)
    let cumipmtEx r nper pv startPeriod endPeriod (pd:PaymentDue) = funcs.CumIPmt(r, nper, pv, startPeriod, endPeriod, int pd)
    let cumprincEx r nper pv startPeriod endPeriod (pd:PaymentDue) = funcs.CumPrinc(r, nper, pv, startPeriod, endPeriod, int pd)
    let ispmtEx r per nper pv = funcs.Ispmt (r, per, nper, pv)
    let coupDaysEx settl mat (freq:Frequency) (basis:DayCountBasis) = funcs.CoupDays (settl, mat, int freq, int basis)
    let coupPCDEx settl mat (freq:Frequency) (basis:DayCountBasis) = float (DateTime.FromOADate (funcs.CoupPcd (settl, mat, int freq, int basis))).Ticks
    let coupNCDEx settl mat (freq:Frequency) (basis:DayCountBasis) = float (DateTime.FromOADate (funcs.CoupNcd (settl, mat, int freq, int basis))).Ticks
    let coupNumEx settl mat (freq:Frequency) (basis:DayCountBasis) = funcs.CoupNum (settl, mat, int freq, int basis)
    let coupDaysBSEx settl mat (freq:Frequency) (basis:DayCountBasis) = funcs.CoupDayBs (settl, mat, int freq, int basis)
    let coupDaysNCEx settl mat (freq:Frequency) (basis:DayCountBasis) = funcs.CoupDaysNc (settl, mat, int freq, int basis)
    let accrIntMEx issue settlement rate par (basis:DayCountBasis) = funcs.AccrIntM(issue, settlement, rate, par, int basis)
    let accrIntEx issue firstInterest settlement rate par (frequency:Frequency) (basis:DayCountBasis) = funcs.AccrInt(issue, firstInterest, settlement, rate, par, int frequency, int basis)
    let priceEx settlement maturity rate yld redemption (frequency:Frequency) basis = funcs.Price(settlement, maturity, rate, yld, redemption, int frequency, basis) 
    let priceMatEx settlement maturity issue rate yld basis = funcs.PriceMat(settlement, maturity, issue, rate, yld, basis)
    let yieldMatEx settlement maturity issue rate pr basis = funcs.YieldMat(settlement, maturity, issue, rate, pr, basis)
    let yearFracEx startDate endDate basis = funcs.YearFrac(startDate, endDate, basis)
    let intRateEx settlement maturity investment redemption basis = funcs.IntRate(settlement, maturity, investment, redemption, basis)
    let receivedEx settlement maturity investment discount basis = funcs.Received(settlement, maturity, investment, discount, basis)
    let discEx settlement maturity pr redemption basis = funcs.Disc(settlement, maturity, pr, redemption, basis)
    let priceDiscEx settlement maturity discount redemption basis = funcs.PriceDisc(settlement, maturity, discount, redemption, basis)
    let yieldDiscEx settlement maturity pr redemption basis = funcs.YieldDisc(settlement, maturity, pr, redemption, basis)
    let TBillEqEx settlement maturity discount = funcs.TBillEq(settlement, maturity, discount)
    let TBillYieldEx settlement maturity pr = funcs.TBillYield(settlement, maturity, pr)
    let TBillPriceEx settlement maturity discount = funcs.TBillPrice(settlement, maturity, discount)
    let dollarDeEx fractionalDollar fraction = funcs.DollarDe(fractionalDollar, fraction)
    let dollarFrEx fractionalDollar fraction = funcs.DollarFr(fractionalDollar, fraction)
    let effectEx nominalRate npery = funcs.Effect(nominalRate, npery)
    let nominalEx effectRate npery = funcs.Nominal(effectRate, npery)
    let durationEx settlement maturity coupon yld frequency basis = funcs.Duration(settlement, maturity, coupon, yld, frequency, basis)
    let mdurationEx settlement maturity coupon yld frequency basis = funcs.MDuration(settlement, maturity, coupon, yld, frequency, basis)
    let oddFPriceEx settlement maturity issue firstCoupon rate yld redemption (frequency:Frequency) basis = funcs.OddFPrice(settlement, maturity, issue, firstCoupon, rate, yld, redemption, int frequency, basis)
    let oddFYieldEx settlement maturity issue firstCoupon rate pr redemption (frequency:Frequency) basis = funcs.OddFYield(settlement, maturity, issue, firstCoupon, rate, pr, redemption, int frequency, basis)
    let oddLPriceEx settlement maturity lastInterest rate yld redemption (frequency:Frequency) basis = funcs.OddLPrice(settlement, maturity, lastInterest, rate, yld, redemption, int frequency, basis)
    let oddLYieldEx settlement maturity lastInterest rate pr redemption (frequency:Frequency) basis = funcs.OddLYield(settlement, maturity, lastInterest, rate, pr, redemption, int frequency, basis)
    let amorLincEx cost datePurchased firstPeriod salvage period rate (basis:DayCountBasis) = funcs.AmorLinc(cost, datePurchased, firstPeriod, salvage, period, rate, basis)
    let amorDegrcEx cost datePurchased firstPeriod salvage period rate (basis:DayCountBasis) = funcs.AmorDegrc(cost, datePurchased, firstPeriod, salvage, period, rate, basis)
   