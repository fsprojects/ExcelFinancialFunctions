namespace Excel.FinancialFunctions

open System
open NUnit.Framework
open TestInfrastructure
open Excel.FinancialFunctions
open Excel.FinancialFunctions.ExcelTesting
open Excel.FinancialFunctions.Common
open Excel.FinancialFunctions.Tvm
open Excel.FinancialFunctions.Loan
open Excel.FinancialFunctions.Irr
open Excel.FinancialFunctions.Depreciation
open Excel.FinancialFunctions.DayCount
open Excel.FinancialFunctions.TestInfrastructure
open Excel.FinancialFunctions.TestsDef
open Excel.FinancialFunctions.Bonds
open Excel.FinancialFunctions.TBill
open Excel.FinancialFunctions.Misc
open Excel.FinancialFunctions.OddBonds
open Excel.FinancialFunctions.TestPreconditions

[<Category("Fast")>]
[<Parallelizable(ParallelScope.Children)>]
type SpotTests () =

    let spotTest1 f1 f2 p1 =
        Assert.AreEqual(f1 p1,f2 p1,precision)
    let spotTest2 f1 f2 p1 p2 =
        Assert.AreEqual(f1 p1 p2,f2 p1 p2,precision)
    let spotTest3 f1 f2 p1 p2 p3 =
        Assert.AreEqual(f1 p1 p2 p3,f2 p1 p2 p3,precision)
    let spotTest4 f1 f2 p1 p2 p3 p4 =
        Assert.AreEqual(f1 p1 p2 p3 p4,f2 p1 p2 p3 p4,precision)
    let spotTest5 f1 f2 p1 p2 p3 p4 p5 =
        Assert.AreEqual(f1 p1 p2 p3 p4 p5,f2 p1 p2 p3 p4 p5,precision)
    let spotTest6 f1 f2 p1 p2 p3 p4 p5 p6 =
        Assert.AreEqual(f1 p1 p2 p3 p4 p5 p6,f2 p1 p2 p3 p4 p5 p6,precision)
    let spotTest7 f1 f2 p1 p2 p3 p4 p5 p6 p7 =
        Assert.AreEqual(f1 p1 p2 p3 p4 p5 p6 p7,f2 p1 p2 p3 p4 p5 p6 p7,precision)
    let spotTest8 f1 f2 p1 p2 p3 p4 p5 p6 p7 p8 =
        Assert.AreEqual(f1 p1 p2 p3 p4 p5 p6 p7 p8,f2 p1 p2 p3 p4 p5 p6 p7 p8,precision)
    let spotTest9 f1 f2 p1 p2 p3 p4 p5 p6 p7 p8 p9 =
        Assert.AreEqual(f1 p1 p2 p3 p4 p5 p6 p7 p8 p9,f2 p1 p2 p3 p4 p5 p6 p7 p8 p9,precision)
       
    [<Test>]    
    member __.calcPv() = 
        spotTest5 calcPv pvEx 0.3 10. 20. 100. PaymentDue.EndOfPeriod
    [<Test>]    
    member __.calcFv() = 
        spotTest5 calcFv fvEx 0.3 10. 20. 100. PaymentDue.EndOfPeriod
    [<Test>]    
    member __.calcPmt() = 
        spotTest5 calcPmt pmtEx 0.3 10. -20. 100. PaymentDue.EndOfPeriod
    [<Test>]    
    member __.calcIpmt() = 
        spotTest6 calcIpmt ipmtEx 0.3 3. 10. -20. 100. PaymentDue.EndOfPeriod 
    [<Test>]    
    member __.calcPpmt() = 
        spotTest6 calcPpmt ppmtEx 0.3 4. 10. -20. 100. PaymentDue.EndOfPeriod 
    [<Test>]    
    member __.calcCumipmt() = 
        spotTest6 calcCumipmt cumipmtEx 0.2 10. 100. 2. 5. PaymentDue.EndOfPeriod
    [<Test>]    
    member __.calcCumprinc() = 
        spotTest6 calcCumprinc cumprincEx 0.2 10. 100. 2. 5. PaymentDue.EndOfPeriod
    [<Test>]    
    member __.calcNper() = 
        spotTest5 calcNper nperEx 0.3 10. 20. -100. PaymentDue.EndOfPeriod
    [<Test>]    
    member __.calcFvSchedule() = 
        spotTest2 calcFvSchedule fvScheduleEx 100. [|0.13;0.14;-0.2;0.34;-0.12|]
    [<Test>]    
    member __.calcIrr() = 
        spotTest2 calcIrr irrEx [|-123.; 12.; 15.; 50.; 200.|] 0.14
    [<Test>]    
    member __.calcNpv() = 
        spotTest2 calcNpv npvEx 0.14 [|-123.; 12.; 15.; 50.; 200.|]
    [<Test>]    
    member __.calcMirr() = 
        spotTest3 calcMirr mirrEx [|-123.; 12.; 15.; 50.; 200.|] 0.14 0.12          
    [<Test>]    
    member __.calcXirr() = 
        spotTest3 calcXirr xirrEx [|-1.;3.;4.|] [|date 1970 3 2; date 1988 2 3; date 1999 3 5|] 0.14
    [<Test>]    
    member __.calcDb() = 
        spotTest5 calcDb dbEx 122. 12. 12. 2. 3.
    [<Test>]    
    member __.calcSln() = 
        spotTest3 calcSln slnEx 122. 20. 12.
    [<Test>]    
    member __.calcSyd() = 
        spotTest4 calcSyd sydEx 130. 10. 10. 4.
    [<Test>]    
    member __.calcDdb() = 
        spotTest5 calcDdb ddbEx 120. 20. 10. 4. 3.
    [<Test>]    
    member __.vdbWrap() = 
        spotTest7 vdbWrap vdbEx 100. 20. 20. 2. 3. 3. VdbSwitch.DontSwitchToStraightLine
    [<Test>]    
    member __.calcIspmt() = 
        spotTest4 calcIspmt ispmtEx 0.15 3. 10. 100.
    [<Test>]    
    member __.calcCoupDays() = 
        spotTest4 calcCoupDays coupDaysEx (date 1984 3 4) (date 1990 4 5) Frequency.Quarterly DayCountBasis.UsPsa30_360  
    [<Test>]    
    member __.coupPCDWrapper() = 
        spotTest4 coupPCDWrapper coupPCDEx (date 1984 3 4) (date 1990 4 5) Frequency.Quarterly DayCountBasis.UsPsa30_360  
    [<Test>]    
    member __.coupNCDWrapper() = 
        spotTest4 coupNCDWrapper coupNCDEx (date 1984 3 4) (date 1990 4 5) Frequency.Quarterly DayCountBasis.UsPsa30_360  
    [<Test>]    
    member __.calcCoupNum() = 
        spotTest4 calcCoupNum coupNumEx (date 1984 3 4) (date 1990 4 5) Frequency.Quarterly DayCountBasis.UsPsa30_360  
    [<Test>]    
    member __.calcCoupDaysBS() = 
        spotTest4 calcCoupDaysBS coupDaysBSEx (date 1984 3 4) (date 1990 4 5) Frequency.Quarterly DayCountBasis.UsPsa30_360  
    [<Test>]    
    member __.calcCoupDaysNC() = 
        spotTest4 calcCoupDaysNC coupDaysNCEx (date 1984 3 4) (date 1990 4 5) Frequency.Quarterly DayCountBasis.UsPsa30_360  
    [<Test>]    
    member __.calcAccrIntM() = 
        spotTest5 calcAccrIntM accrIntMEx (date 1984 3 4) (date 1991 4 5) 0.07 120. DayCountBasis.UsPsa30_360    
    [<Test>]    
    member __.calcAccrIntWrap() = 
        spotTest7 calcAccrIntWrap accrIntEx (date 1984 3 4) (date 1994 3 4) (date 1991 4 5) 0.07 120. Frequency.Quarterly DayCountBasis.UsPsa30_360    
    [<Test>]    
    member __.calcPrice() = 
        spotTest7 calcPrice priceEx (date 1984 3 4) (date 1990 3 4) 0.07 0.1 110. Frequency.Quarterly DayCountBasis.ActualActual        
    [<Test>]    
    member __.calcPriceMat() = 
        spotTest6 calcPriceMat priceMatEx (date 2008 2 13) (date 2009 4 13) (date 2007 11 11) 0.061 0.061 DayCountBasis.UsPsa30_360
    [<Test>]    
    member __.calcYieldMat() = 
        spotTest6 calcYieldMat yieldMatEx (date 2008 2 13) (date 2009 4 13) (date 2007 11 11) 0.061 120. DayCountBasis.UsPsa30_360
    [<Test>]    
    member __.calcYearFrac() = 
        spotTest3 calcYearFrac yearFracEx (date 2008 2 13) (date 2009 4 13) DayCountBasis.ActualActual
    [<Test>]    
    member __.calcIntRate1() = 
        spotTest5 calcIntRate intRateEx (date 2008 2 13) (date 2010 4 13) 100. 150. DayCountBasis.UsPsa30_360
    [<Test>]    
    member __.calcIntRate2() = 
        spotTest5 calcIntRate intRateEx (date 2008 3 13) (date 2010 5 13) 100. 0.15 DayCountBasis.UsPsa30_360
    [<Test>]    
    member __.calcDisc() = 
        spotTest5 calcDisc discEx (date 2008 2 13) (date 2011 5 13) 75. 100. DayCountBasis.UsPsa30_360
    [<Test>]    
    member __.calcPriceDisc() = 
        spotTest5 calcPriceDisc priceDiscEx (date 2008 2 13) (date 2013 5 13) 0.25 100. DayCountBasis.UsPsa30_360
    [<Test>]    
    member __.calcYieldDisc() = 
        spotTest5 calcYieldDisc yieldDiscEx (date 2008 2 28) (date 2011 5 13) 75. 100. DayCountBasis.UsPsa30_360
    [<Test>]    
    member __.calcTBillEq() = 
        spotTest3 calcTBillEq TBillEqEx (date 2008 2 13) (date 2009 1 11) 0.25
    [<Test>]    
    member __.calcTBillYield() = 
        spotTest3 calcTBillYield TBillYieldEx (date 2008 2 28) (date 2009 2 27) 0.25
    [<Test>]    
    member __.calcTBillPrice() = 
        spotTest3 calcTBillPrice TBillPriceEx (date 2008 2 29) (date 2009 2 27) 0.25
    [<Test>]    
    member __.calcDollarDe() = 
        spotTest2 calcDollarDe dollarDeEx   1.125 16.
    [<Test>]    
    member __.calcDollarFr() = 
        spotTest2 calcDollarFr dollarFrEx   1.125 16.
    [<Test>]    
    member __.calcEffect() = 
        spotTest2 calcEffect effectEx   0.0525 4.
    [<Test>]    
    member __.calcNominal() = 
        spotTest2 calcNominal nominalEx   0.053543 4.
    [<Test>]    
    member __.calcDuration() = 
        spotTest6 calcDuration durationEx (date 2008 2 13) (date 2011 5 13) 100. 0.07 Frequency.Quarterly DayCountBasis.UsPsa30_360
    [<Test>]    
    member __.calcMDuration() = 
        spotTest6 calcMDuration mdurationEx (date 2008 2 13) (date 2011 5 13) 100. 0.07 Frequency.Quarterly DayCountBasis.UsPsa30_360   
    [<Test>]    
    member __.calcOddFPrice() = 
        spotTest9 calcOddFPrice oddFPriceEx (date 2008 11 11) (date 2021 3 1) (date 2008 10 15) (date 2009 3 1) 0.0785 0.0625 100. Frequency.SemiAnnual DayCountBasis.ActualActual 
    [<Test>]    
    member __.calcOddFYield() = 
        spotTest9 calcOddFYield oddFYieldEx (date 2008 11 11) (date 2021 3 1) (date 2008 10 15) (date 2009 3 1) 0.0575 84.5 100. Frequency.SemiAnnual DayCountBasis.UsPsa30_360 
    [<Test>]    
    member __.calcOddLPrice() = 
        spotTest8 calcOddLPrice oddLPriceEx (date 2008 2 7) (date 2008 6 15) (date 2007 10 15) 0.0375 0.0405 100. Frequency.SemiAnnual DayCountBasis.UsPsa30_360 
    [<Test>]    
    member __.calcOddLYield() = 
        spotTest8 calcOddLYield oddLYieldEx (date 2008 4 20) (date 2008 6 15) (date 2007 12 24) 0.0375 99.875 100. Frequency.SemiAnnual DayCountBasis.UsPsa30_360 
    [<Test>]    
    member __.calcAmorLinc() = 
        spotTest7 calcAmorLinc amorLincEx 2400. (date 2008 8 19) (date 2008 12 31) 300. 1. 0.15 DayCountBasis.ActualActual
    [<Test>]    
    member __.amorDegrcWrapper() = 
        spotTest7 amorDegrcWrapper amorDegrcEx 2400. (date 2008 8 19) (date 2008 12 31) 300. 1. 0.15 DayCountBasis.ActualActual

    // Need to test this manually because of the differences in root finding algos
    [<Test>]    
    member __.testRate() =
        spotTest6 calcRate rateEx 1. 10. 100. -100. PaymentDue.EndOfPeriod 0.15
        spotTest6 calcRate rateEx 5. 20. 120. -50. PaymentDue.BeginningOfPeriod 0.
        spotTest6 calcRate rateEx 10. -10. 0. 100. PaymentDue.EndOfPeriod -0.15
        spotTest6 calcRate rateEx 25. -40. -200. 100. PaymentDue.BeginningOfPeriod 0.15
    
    [<Test>]    
    member __.calcXnpv() = 
        let result = calcXnpv 0.14 [1.;3.;4.] [date 1970 3 2; date 1988 2 3; date 1999 3 5]
        Assert.AreEqual(1.375214,result,precision)

    [<Test>]    
    member __.calcYield() = 
        let result = calcYield (date 2008 2 15) (date 2016 11 15) 0.0575 95.04287 100. Frequency.SemiAnnual DayCountBasis.UsPsa30_360
        Assert.AreEqual(0.065,result,precision)
    
    [<Test>]
    member __.testXirrBugs() =
        let t = @"-185550.98 5/15/2008
        -231887.53 5/19/2008
        -26756.74 5/30/2008
        -384010.86 6/20/2008
        -27114.54 6/26/2008
        -458667.97 8/21/2008
        -217853.67 9/8/2008
        -424924.25 10/13/2008
        -75076.01 10/14/2008
        -389630.32 10/24/2008
        -112094.2 11/19/2008
        -25646.4 11/21/2008
        -24164.69 11/21/2008
        -1222.08 11/21/2008
        -556.91 12/3/2008
        1204954.004 12/5/2008"

        let pairs = t.Split([|' '; '\n'; '\r'|]) |> Array.filter (fun x -> not(x = ""))
        let dates = pairs |> Array.filter (fun x -> fst(DateTime.TryParse(x))) |> Array.map (fun x -> DateTime.Parse(x))  
        let values = pairs |> Array.filter (fun x -> not(fst(DateTime.TryParse(x)))) |> Array.map (fun x -> float x)
        let guess = - 0.1
        spotTest3 calcXirr xirrEx values dates guess
    
        let values = [|105091006.;-103250941.864729|]
        let dates = [|date 2000 4 10; date 2000 4 30|]
        spotTest3 calcXirr xirrEx values dates guess
    
        let values = [|206101714.849377;-156650972.54265|]
        let dates = [|date 2001 2 28; date 2001 3 31|]
        spotTest3 calcXirr xirrEx values dates guess
    
        let values = [|15108163.3840923;-75382259.6628424|]
        let dates = [|date 2000 2 29; date 2000 3 31|]
        let result = calcXirr values dates guess

        Assert.AreEqual(165601346.13484925,result,precision)
              
    [<Test>]
    member __.testOddFYield () =
        spotTest9 calcOddFYield oddFYieldEx (date 2008 12 11) (date 2021 4 1) (date 2008 10 15) (date 2009 4 1) 0.06 100. 100. Frequency.Quarterly DayCountBasis.ActualActual 
        spotTest9 calcOddFYield oddFYieldEx (date 2009 2 28) (date 2020 5 30) (date 2008 9 15) (date 2009 5 30) 0.05 75. 89. Frequency.Annual DayCountBasis.Actual360 
        spotTest9 calcOddFYield oddFYieldEx (date 2009 10 31) (date 2021 12 31) (date 2009 10 15) (date 2009 12 31) 0.06 100. 100. Frequency.Quarterly DayCountBasis.ActualActual 
