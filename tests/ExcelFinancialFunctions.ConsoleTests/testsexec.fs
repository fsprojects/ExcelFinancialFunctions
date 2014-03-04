// Running the testcases. Two things:
//    1. Running the full suite
//    2. Running some spot tests
// Running 1. can take time. So the patterns is to run just one of 1. when developing a feature and run all of it at the end.
// The spot test are a quick check that you are not breaking anything else as you work on the feature.
#light

open System
open System.Collections
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

try         
    let tests = [|
        // --------------------------------------------------------------------------------------------------
        //"RATE", Excel uses a different root finding algo. Sometimes mine is better, sometimes Excel's is. Using the TestRate function instead.
        //"ODDFYIELD", Excel uses a different root finding algo. Sometimes mine is better, sometimes Excel's is. Using the TestOddFYield function instead.        
        //"XNPV", the excel object model has an xnpv function with a different number of args. Tested through XIRR and separate testcase.
        //"YIELD", the excel object model lacks this function. Tested through the price function and separate testcase.
        // --------------------------------------------------------------------------------------------------
        "PV", (fun _ -> test5 calcPv pvEx rates npers pmts fvs dues tryPv);
        "FV", (fun _ -> test5 calcFv fvEx rates npers pmts pvs dues tryFv);  
        "PMT", (fun _ -> test5 calcPmt pmtEx rates npers pvs fvs dues tryPmt);
        "NPER", (fun _ -> test5 calcNper nperEx rates pmts pvs fvs dues tryNper);        
        "IPMT", (fun _ -> test6 calcIpmt ipmtEx rates pers npers pvs fvs dues tryIpmt); 
        "PPMT", (fun _ -> test6 calcPpmt ppmtEx rates pers npers pvs fvs dues tryPpmt);
        "CUMIPMT", (fun _ -> test6 calcCumipmt cumipmtEx rates npers pvs pers endPers dues tryCumipmt);
        "CUMPRINC", (fun _ -> test6 calcCumprinc cumprincEx rates npers pvs pers endPers dues tryCumprinc);
        "ISPMT", (fun _ -> test4 calcIspmt ispmtEx rates pers npers pvs tryIspmt);      
        "FVSCHEDULE", (fun _ -> test2 calcFvSchedule fvScheduleEx pvs testInterests precondOk2);
        "IRR", (fun _ -> test2 calcIrr irrEx testCfs guesses precondOk2);
        "NPV", (fun _ -> test2 calcNpv npvEx rates testCfs tryNpv);
        "MIRR", (fun _ -> test3 calcMirr mirrEx testCfs rates rates tryMirr);
        "XIRR", fun _-> test3 calcXirr xirrEx testCfs testDates guesses tryXirr;
        "DB", fun _ -> test5 calcDb dbEx testCosts testSalvages testLives testPeriods testMonths tryDb;
        "SLN", fun _ -> test3 calcSln slnEx testCosts testSalvages testLives trySln;
        "SYD", fun _ -> test4 calcSyd sydEx testCosts testSalvages testLives testPeriods trySyd;
        "DDB", fun _ -> test5 calcDdb ddbEx testCosts testSalvages testLives testDdbPeriods testFactors tryDdb;
        "VDB excluding fractional startdates", fun _ -> test7 vdbWrap vdbEx testCosts testSalvages testLives testPeriods testEndPeriods testFactors testVdbSwitch tryVdb; 
        "AMORLINC", fun _ -> test7 calcAmorLinc amorLincEx testCosts testIssueDates testFirstInterestDates testSalvages testPeriods testBondRates testDayCountBasis tryAmorLinc;
        "AMORDEGRC", fun _ -> test7 amorDegrcWrapper amorDegrcEx testCosts testIssueDates testFirstInterestDates testSalvages testPeriods testDeprRates testDayCountBasis tryAmorDegrc;
        "COUPDAYS excluding leap years", fun _ -> test4 calcCoupDays coupDaysEx testSettlDates testMatDates testFrequency testDayCountBasis tryCoupDays;
        "COUPDAYSBS", fun _ -> test4 calcCoupDaysBS coupDaysBSEx testSettlDates testMatDates testFrequency testDayCountBasis tryCoupDaysBS;
        "COUPDAYSNC", fun _ -> test4 calcCoupDaysNC coupDaysNCEx testSettlDates testMatDates testFrequency testDayCountBasis tryCoupDaysNC;
        "COUPNUM", fun _ -> test4 calcCoupNum coupNumEx testSettlDates testMatDates testFrequency testDayCountBasis tryCoupNum;
        "COUPPCD", fun _ -> test4 coupPCDWrapper coupPCDEx testSettlDates testMatDates testFrequency testDayCountBasis tryCoupPCD;      
        "COUPNCD", fun _ -> test4 coupNCDWrapper coupNCDEx testSettlDates testMatDates testFrequency testDayCountBasis tryCoupNCD;     
        "ACCRINTM", fun _ -> test5 calcAccrIntM accrIntMEx testIssue testSettl testBondRates testPars testDayCountBasis tryAccrIntM;
        "ACCRINT", fun _ -> test7 calcAccrIntWrap accrIntEx testIssue testFirstInt testSettl testBondRates testPars testFrequency testDayCountBasis tryAccrInt;
        "PRICE", fun _ -> test7 calcPrice priceEx testSettlDates testMatDates testBondRates testYlds testRedemptions testFrequency testDayCountBasis tryPrice 
        "PRICEMAT", fun _ -> test6 calcPriceMat priceMatEx testSettlDates testMatDates testIssue testBondRates testYlds testDayCountBasis tryPriceMat 
        "YIELDMAT", fun _ -> test6 calcYieldMat yieldMatEx testSettlDates testMatDates testIssue testBondRates testPrices testDayCountBasis tryYieldMat 
        "YEARFRAC", fun _ -> test3 calcYearFrac yearFracEx testSDates testEDates testDayCountBasis tryYearFrac; 
        "INTRATE", fun _ -> test5 calcIntRate intRateEx testSettlDates testMatDates testInvestments testRedemptions testDayCountBasis tryIntRate; 
        "RECEIVED", fun _ -> test5 calcReceived receivedEx testSettlDates testMatDates testInvestments testDiscounts testDayCountBasis tryReceived; 
        "DISC", fun _ -> test5 calcDisc discEx testSettlDates testMatDates testInvestments testRedemptions testDayCountBasis tryDisc; 
        "PRICEDISC", fun _ -> test5 calcPriceDisc priceDiscEx testSettlDates testMatDates testDiscounts testRedemptions testDayCountBasis tryPriceDisc; 
        "YIELDDISC", fun _ -> test5 calcYieldDisc yieldDiscEx testSettlDates testMatDates testInvestments testRedemptions testDayCountBasis tryYieldDisc;
        "TBILLEQ", fun _ -> test3 calcTBillEq TBillEqEx testSettlDates testTBillMat testDiscounts tryTBillEq; 
        "TBILLYIELD", fun _ -> test3 calcTBillYield TBillYieldEx testSettlDates testTBillMat testPrices tryTBillYield; 
        "TBILLPrice", fun _ -> test3 calcTBillPrice TBillPriceEx testSettlDates testTBillMat testDiscounts tryTBillPrice; 
        "DOLLARDE", fun _ -> test2 calcDollarDe dollarDeEx testFractionalDollars testFractions tryDollarDe; 
        "DOLLARFR", fun _ -> test2 calcDollarFr dollarFrEx testFractionalDollars testFractions tryDollarFr; 
        "EFFECT", fun _ -> test2 calcEffect effectEx rates testPeriods tryEffect;
        "NOMINAL", fun _ -> test2 calcNominal nominalEx rates testPeriods tryNominal; 
        "DURATION", fun _ -> test6 calcDuration durationEx testSettlDates testMatDates testInvestments testYlds testFrequency testDayCountBasis tryDuration; 
        "MDURATION", fun _ -> test6 calcMDuration mdurationEx testSettlDates testMatDates testInvestments testYlds testFrequency testDayCountBasis tryMDuration;
        "ODDFPRICE", fun _ -> test9 calcOddFPrice oddFPriceEx testSettlDates2 testMatDates testIssueDates testFirstInterestDates testBondRates testYlds testRedemptions testFrequency testDayCountBasis tryOddFPrice;
        "ODDLPRICE", fun _ -> test8 calcOddLPrice oddLPriceEx testSettlDates2 testMatDates testIssueDates testBondRates testYlds testRedemptions testFrequency testDayCountBasis tryOddLPrice;
        "ODDLYIELD", fun _ -> test8 calcOddLYield oddLYieldEx testSettlDates2 testMatDates testIssueDates testBondRates testRedemptions testRedemptions testFrequency testDayCountBasis tryOddLYield;
        |]
    
    let results = cpu_map tests
    printResults results
    
    spotTest5 calcPv pvEx 0.3 10. 20. 100. PaymentDue.EndOfPeriod
    spotTest5 calcFv fvEx 0.3 10. 20. 100. PaymentDue.EndOfPeriod
    spotTest5 calcPmt pmtEx 0.3 10. -20. 100. PaymentDue.EndOfPeriod
    spotTest6 calcIpmt ipmtEx 0.3 3. 10. -20. 100. PaymentDue.EndOfPeriod 
    spotTest6 calcPpmt ppmtEx 0.3 4. 10. -20. 100. PaymentDue.EndOfPeriod 
    spotTest6 calcCumipmt cumipmtEx 0.2 10. 100. 2. 5. PaymentDue.EndOfPeriod
    spotTest6 calcCumprinc cumprincEx 0.2 10. 100. 2. 5. PaymentDue.EndOfPeriod
    spotTest5 calcNper nperEx 0.3 10. 20. -100. PaymentDue.EndOfPeriod
    testRate () // Need to test this manually because of the differences in root finding algos
    spotTest2 calcFvSchedule fvScheduleEx 100. [|0.13;0.14;-0.2;0.34;-0.12|]
    spotTest2 calcIrr irrEx [|-123.; 12.; 15.; 50.; 200.|] 0.14
    spotTest2 calcNpv npvEx 0.14 [|-123.; 12.; 15.; 50.; 200.|]
    spotTest3 calcMirr mirrEx [|-123.; 12.; 15.; 50.; 200.|] 0.14 0.12

    let xnpvRes = calcXnpv 0.14 [1.;3.;4.] [date 1970 3 2; date 1988 2 3; date 1999 3 5]
    if not(areEqual xnpvRes 1.375214) then printfn "%f different from %f" xnpvRes 1.375214
         
    spotTest3 calcXirr xirrEx [|-1.;3.;4.|] [|date 1970 3 2; date 1988 2 3; date 1999 3 5|] 0.14
    testXirrBugs ()
    spotTest5 calcDb dbEx 122. 12. 12. 2. 3.
    spotTest3 calcSln slnEx 122. 20. 12.
    spotTest4 calcSyd sydEx 130. 10. 10. 4.
    spotTest5 calcDdb ddbEx 120. 20. 10. 4. 3.
    spotTest7 vdbWrap vdbEx 100. 20. 20. 2. 3. 3. VdbSwitch.DontSwitchToStraightLine
    spotTest4 calcIspmt ispmtEx 0.15 3. 10. 100.
    spotTest4 calcCoupDays coupDaysEx (date 1984 3 4) (date 1990 4 5) Frequency.Quarterly DayCountBasis.UsPsa30_360  
    spotTest4 coupPCDWrapper coupPCDEx (date 1984 3 4) (date 1990 4 5) Frequency.Quarterly DayCountBasis.UsPsa30_360  
    spotTest4 coupNCDWrapper coupNCDEx (date 1984 3 4) (date 1990 4 5) Frequency.Quarterly DayCountBasis.UsPsa30_360  
    spotTest4 calcCoupNum coupNumEx (date 1984 3 4) (date 1990 4 5) Frequency.Quarterly DayCountBasis.UsPsa30_360  
    spotTest4 calcCoupDaysBS coupDaysBSEx (date 1984 3 4) (date 1990 4 5) Frequency.Quarterly DayCountBasis.UsPsa30_360  
    spotTest4 calcCoupDaysNC coupDaysNCEx (date 1984 3 4) (date 1990 4 5) Frequency.Quarterly DayCountBasis.UsPsa30_360  
    spotTest5 calcAccrIntM accrIntMEx (date 1984 3 4) (date 1991 4 5) 0.07 120. DayCountBasis.UsPsa30_360    
    spotTest7 calcAccrIntWrap accrIntEx (date 1984 3 4) (date 1994 3 4) (date 1991 4 5) 0.07 120. Frequency.Quarterly DayCountBasis.UsPsa30_360    
    spotTest7 calcPrice priceEx (date 1984 3 4) (date 1990 3 4) 0.07 0.1 110. Frequency.Quarterly DayCountBasis.ActualActual
    
    let yieldRes = calcYield (date 2008 2 15) (date 2016 11 15) 0.0575 95.04287 100. Frequency.SemiAnnual DayCountBasis.UsPsa30_360
    if not(areEqual yieldRes 0.065) then printfn "%f different from %f" yieldRes 0.065

    spotTest6 calcPriceMat priceMatEx (date 2008 2 13) (date 2009 4 13) (date 2007 11 11) 0.061 0.061 DayCountBasis.UsPsa30_360
    spotTest6 calcYieldMat yieldMatEx (date 2008 2 13) (date 2009 4 13) (date 2007 11 11) 0.061 120. DayCountBasis.UsPsa30_360
    spotTest3 calcYearFrac yearFracEx (date 2008 2 13) (date 2009 4 13) DayCountBasis.ActualActual
    spotTest5 calcIntRate intRateEx (date 2008 2 13) (date 2010 4 13) 100. 150. DayCountBasis.UsPsa30_360
    spotTest5 calcIntRate intRateEx (date 2008 3 13) (date 2010 5 13) 100. 0.15 DayCountBasis.UsPsa30_360
    spotTest5 calcDisc discEx (date 2008 2 13) (date 2011 5 13) 75. 100. DayCountBasis.UsPsa30_360
    spotTest5 calcPriceDisc priceDiscEx (date 2008 2 13) (date 2013 5 13) 0.25 100. DayCountBasis.UsPsa30_360
    spotTest5 calcYieldDisc yieldDiscEx (date 2008 2 28) (date 2011 5 13) 75. 100. DayCountBasis.UsPsa30_360
    spotTest3 calcTBillEq TBillEqEx (date 2008 2 13) (date 2009 1 11) 0.25
    spotTest3 calcTBillYield TBillYieldEx (date 2008 2 28) (date 2009 2 27) 0.25
    spotTest3 calcTBillPrice TBillPriceEx (date 2008 2 29) (date 2009 2 27) 0.25
    spotTest2 calcDollarDe dollarDeEx   1.125 16.
    spotTest2 calcDollarFr dollarFrEx   1.125 16.
    spotTest2 calcEffect effectEx   0.0525 4.
    spotTest2 calcNominal nominalEx   0.053543 4.
    spotTest6 calcDuration durationEx (date 2008 2 13) (date 2011 5 13) 100. 0.07 Frequency.Quarterly DayCountBasis.UsPsa30_360
    spotTest6 calcMDuration mdurationEx (date 2008 2 13) (date 2011 5 13) 100. 0.07 Frequency.Quarterly DayCountBasis.UsPsa30_360   
    spotTest9 calcOddFPrice oddFPriceEx (date 2008 11 11) (date 2021 3 1) (date 2008 10 15) (date 2009 3 1) 0.0785 0.0625 100. Frequency.SemiAnnual DayCountBasis.ActualActual 
    spotTest9 calcOddFYield oddFYieldEx (date 2008 11 11) (date 2021 3 1) (date 2008 10 15) (date 2009 3 1) 0.0575 84.5 100. Frequency.SemiAnnual DayCountBasis.UsPsa30_360 
    testOddFYield()
    spotTest8 calcOddLPrice oddLPriceEx (date 2008 2 7) (date 2008 6 15) (date 2007 10 15) 0.0375 0.0405 100. Frequency.SemiAnnual DayCountBasis.UsPsa30_360 
    spotTest8 calcOddLYield oddLYieldEx (date 2008 4 20) (date 2008 6 15) (date 2007 12 24) 0.0375 99.875 100. Frequency.SemiAnnual DayCountBasis.UsPsa30_360 
    spotTest7 calcAmorLinc amorLincEx 2400. (date 2008 8 19) (date 2008 12 31) 300. 1. 0.15 DayCountBasis.ActualActual
    spotTest7 amorDegrcWrapper amorDegrcEx 2400. (date 2008 8 19) (date 2008 12 31) 300. 1. 0.15 DayCountBasis.ActualActual
    printEndSpotTests ()
with
    | ex -> printfn "%s" ex.Message        

endTests () |> ignore
