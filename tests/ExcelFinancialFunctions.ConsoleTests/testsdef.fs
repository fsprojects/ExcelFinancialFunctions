// All the test values that I want to throw at the testcase infrastructure. Just add values here if you want to run 1,000,0000,000 tests
#light

namespace Excel.FinancialFunctions

open System
open System.Collections
open Excel.FinancialFunctions.ExcelTesting
open Excel.FinancialFunctions.Common
open Excel.FinancialFunctions.Tvm
open Excel.FinancialFunctions.Irr
open Excel.FinancialFunctions.TestInfrastructure
open Excel.FinancialFunctions.Depreciation
open Excel.FinancialFunctions.DayCount
open Excel.FinancialFunctions.Bonds
open Excel.FinancialFunctions.OddBonds

module internal TestsDef =

    // Test values
    let pvs = [ -300.; -100.; -5.4;  0.;  100.;  150.5]
    let fvs = pvs
    let rates = [ -1.5; -1.; -2.0;  -0.4;  -0.1;  0.;  0.6;  1.5]
    let npers = [ -2.; 0.;  1.;  2.;  2.7; 10.]
    let pers = [ 0.; 1.; 1.3; 2.; 2.5; 5.]
    let endPers = [ 1.2; 2.; 3.; 5.; 7.]
    let dues = [PaymentDue.BeginningOfPeriod; PaymentDue.EndOfPeriod]
    let pmts = [ 50.;  30.;  -10.;  0.]
    let guesses = [ 0.15; 0.5; 0.]
    let pvRates = [0.; -10.; -20.; -100.5]
    let fvRates = [0.; 10.; 100.; -20.]
    let testInterests = [ [|0.3; -0.5; 0.2; 1.3; -0.2|]; [|0.3; -0.5; 0.2; 0.; -1.2|]]       
    let testCfs = [[| -100.;10.;10.;100.|]; [| -100.; -10.; 10.; 100.|]; [| -200.; 0.; 10.; -10.; 300.|]]
    let testDates = [
        [|date 1970 4 1; date 1972 2 12; date 1980 4 23; date 1983 3 30|];
        [|date 1970 4 1; date 1973 4 12; date 1983 5 23; date 1987 4 30|];
        [|date 1970 4 1; date 1974 2 14; date 1985 6 26; date 1989 5 4|]]     
    let testCosts = [100.; 200.; -100.]
    let testSalvages = [10.;50.; 0.; -20.]
    let testLives = [0.; 1.; 13.; 12.7; 40.]
    let testPeriods = [0.; 0.3; 1.; 1.7; 2.; 10.; 11.3; 13.]
    let testDdbPeriods = [0.3; 1.; 2.; 10.; 11.; 13.]
    let testMonths = [1.; 4.; 9.]
    let testFactors = [1.;3. ; 4.5; 50.3]
    let testEndPeriods = [0.8; 1.; 3.; 4.2; 13.; 3.3; 20.]
    let testVdbSwitch = [ VdbSwitch.DontSwitchToStraightLine; VdbSwitch.SwitchToStraightLine]
    let testDayCountBasis = [DayCountBasis.Actual360;DayCountBasis.Actual365;DayCountBasis.ActualActual;DayCountBasis.Europ30_360;DayCountBasis.UsPsa30_360]
    let testSettlDates = [date 1980 2 15; date 1980 3 15; date 1993 12 31; date 2003 2 14; date 2007 10 31; date 1993 2 28; date 1981 3 31; date 2004 3 31]
    let testMatDates = [date 2000 2 28; date 1995 11 30; date 1980 5 4; date 2010 6 30; date 2008 2 29;  date 1994 1 31;  date 2003 5 14; date 2009 10 1; date 2010 6 5; date 2004 3 31]
    let testMatDates2 = [date 2000 2 28; date 1995 11 30; date 1980 5 4; date 2010 6 30;  date 2003 5 14; date 2004 3 31]
    let testFirstInterestDates = testMatDates2 |> List.map (fun x -> x.AddYears(-1)) |> List.append [date 2000 2 29]
    let testSettlDates2 = testFirstInterestDates |> List.map (fun x -> x.AddYears(-1))
    let testIssueDates = testSettlDates2 |> List.map (fun x -> x.AddYears(-1))
    let testFrequency = [Frequency.Annual;Frequency.SemiAnnual;Frequency.Quarterly]
    let testIssue = [date 1990 3 4; date 1993 2 28; date 1995 5 31; date 2000 3 28; date 1999 4 2]
    let testSettl = [date 1992 3 4; date 1995 3 1; date 1995 2 28; date 1996 3 30; date 2010 6 5; date 2000 7 2]
    let testTBillMat1 = testSettlDates |> List.map (fun x -> x.AddDays(+190.))
    let testTBillMat2 = testSettlDates |> List.map (fun x -> x.AddDays(+45.))
    let testTBillMat3 = testSettlDates |> List.map idem
    let testTBillMat = testTBillMat1 |> List.append testTBillMat2 |> List.append testTBillMat3
    let testBondRates = [0.07; 0.1]
    let testDeprRates = [0.15; 0.3; 0.5; 0.1; 0.07]
    let testYlds = [ 0.03; 0.1]
    let testRedemptions = [100.; 67.; 130.]
    let testInvestments = [100.; 23.; 200.]
    let testPars = [ 10000.; 12030.34]
    let testFirstInt = [date 1990 1 4; date 1993 3 31; date 1988 2 28; date 1986 3 30; date 2010 7 5; date 2002 1 2];
    let testSDates =  [date 1980 3 4; date 1993 12 31; date 2003 2 14; date 2007 10 31; date 1993 2 28; date 1981 3 31; date 2000 2 28;
        date 1992 1 4; date 1995 3 1; date 1995 2 28; date 1996 3 30; date 2010 6 5; date 2000 1 2;
        date 1992 3 4; date 1995 3 1; date 1998 3 30; date 2010 10 5; date 2004 7 2;
        date 1990 3 4; date 1993 2 28; date 1995 5 31; date 2000 2 28; date 1999 3 31]
    let testEDates = testSDates |> List.map (fun x -> x.AddDays(+1.))
    let testPrices = [ 75. ; 100.; 130.]
    let testDiscounts = [ 0.01; 0.25; 0.75; 2.]
    let testFractionalDollars = [0.34; 1.02; 2.34; -1.5]
    let testFractions = [ 1.; 17.; 20.] 

    let testRate () =
        spotTest6 calcRate rateEx 1. 10. 100. -100. PaymentDue.EndOfPeriod 0.15
        spotTest6 calcRate rateEx 5. 20. 120. -50. PaymentDue.BeginningOfPeriod 0.
        spotTest6 calcRate rateEx 10. -10. 0. 100. PaymentDue.EndOfPeriod -0.15
        spotTest6 calcRate rateEx 25. -40. -200. 100. PaymentDue.BeginningOfPeriod 0.15
        
    let testXirrBugs () =
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
        let res = calcXirr values dates guess
        areEqual res 165601346.13484925 |> elseThrow "XIrr manual test fails"
                   
    let testOddFYield () =
        spotTest9 calcOddFYield oddFYieldEx (date 2008 12 11) (date 2021 4 1) (date 2008 10 15) (date 2009 4 1) 0.06 100. 100. Frequency.Quarterly DayCountBasis.ActualActual 
        spotTest9 calcOddFYield oddFYieldEx (date 2009 2 28) (date 2020 5 30) (date 2008 9 15) (date 2009 5 30) 0.05 75. 89. Frequency.Annual DayCountBasis.Actual360 
        spotTest9 calcOddFYield oddFYieldEx (date 2009 10 31) (date 2021 12 31) (date 2009 10 15) (date 2009 12 31) 0.06 100. 100. Frequency.Quarterly DayCountBasis.ActualActual 

    let vdbWrap cost salvage life startPeriod endPeriod factor bflag = calcVdb cost salvage life startPeriod endPeriod factor bflag  

    let coupPCDWrapper settlement maturity (frequency:Frequency) basis =
        float (calcCoupPCD settlement maturity frequency basis).Ticks

    let coupNCDWrapper settlement maturity (frequency:Frequency) basis =
        float (calcCoupNCD settlement maturity frequency basis).Ticks

    let calcAccrIntWrap issue firstInterest settlement rate par (frequency:Frequency) basis =
        calcAccrInt issue firstInterest settlement rate par frequency basis AccrIntCalcMethod.FromIssueToSettlement    

    let amorDegrcWrapper cost datePurchased firstPeriod salvage period rate basis =
        calcAmorDegrc cost datePurchased firstPeriod salvage period rate basis true
