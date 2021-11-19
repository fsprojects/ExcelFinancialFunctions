namespace Excel.FinancialFunctions

open NUnit.Framework
open TestInfrastructure
open Excel.FinancialFunctions.ExcelTesting
open Excel.FinancialFunctions.Tvm
open Excel.FinancialFunctions.Loan
open Excel.FinancialFunctions.Irr
open Excel.FinancialFunctions.Depreciation
open Excel.FinancialFunctions.DayCount
open Excel.FinancialFunctions.TestsDef
open Excel.FinancialFunctions.Bonds
open Excel.FinancialFunctions.TBill
open Excel.FinancialFunctions.Misc
open Excel.FinancialFunctions.OddBonds
open Excel.FinancialFunctions.TestPreconditions

[<Parallelizable(ParallelScope.Children)>]
type MatrixTests () =
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
        //"XIRR", fun _-> test3 calcXirr xirrEx testCfs testDates guesses tryXirr;
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

    [<TestCase("PV")>]
    [<TestCase("FV")>]
    [<TestCase("PMT")>]
    [<TestCase("NPER")>]
    [<TestCase("IPMT")>]
    [<TestCase("PPMT")>]
    [<TestCase("CUMIPMT",Category="Fast")>]
    [<TestCase("CUMPRINC",Category="Fast")>]
    [<TestCase("ISPMT")>]
    [<TestCase("DB")>]
    [<TestCase("SLN",Category="Fast")>]
    [<TestCase("SYD",Category="Fast")>]
    [<TestCase("DDB")>]
    [<TestCase("VDB excluding fractional startdates")>]
    [<TestCase("FVSCHEDULE",Category="Fast")>]
    [<TestCase("MIRR",Category="Fast")>]
    [<TestCase("NPV")>]
    [<TestCase("ISPMT")>]
    [<TestCase("IRR",Category="Fast")>]
    [<TestCase("AMORLINC")>]
    [<TestCase("AMORDEGRC")>]
    [<TestCase("COUPDAYS excluding leap years")>]
    [<TestCase("COUPDAYSBS")>]
    [<TestCase("COUPDAYSNC")>]
    [<TestCase("COUPNUM")>]
    [<TestCase("COUPPCD")>]
    [<TestCase("COUPNCD")>]
    [<TestCase("ACCRINTM")>]
    [<TestCase("ACCRINT")>]
    [<TestCase("PRICE")>]
    [<TestCase("PRICEMAT")>]
    [<TestCase("YIELDMAT")>]
    [<TestCase("YEARFRAC")>]
    [<TestCase("INTRATE")>]
    [<TestCase("RECEIVED")>]
    [<TestCase("DISC")>]
    [<TestCase("PRICEDISC")>]
    [<TestCase("YIELDDISC")>]
    [<TestCase("TBILLEQ")>]
    [<TestCase("TBILLYIELD",Category="Fast")>]
    [<TestCase("TBILLPrice",Category="Fast")>]
    [<TestCase("DOLLARDE",Category="Fast")>]
    [<TestCase("DOLLARFR",Category="Fast")>]
    [<TestCase("EFFECT",Category="Fast")>]
    [<TestCase("NOMINAL",Category="Fast")>]
    [<TestCase("DURATION")>]
    [<TestCase("MDURATION")>]
    [<TestCase("ODDFPRICE")>]
    [<TestCase("ODDLPRICE")>]
    [<TestCase("ODDLYIELD")>]
    member __.RunMatrix test =
        let found = Array.tryFind (fun x -> fst x = test) tests
        let (_,func) = found.Value
        let (tries,success,_,_) = func ()
        Assert.AreEqual(success,tries)