namespace Excel.FinancialFunctions

[<assembly:System.Runtime.CompilerServices.InternalsVisibleTo "ExcelFinancialFunctions.ConsoleTests">]
[<assembly:System.Runtime.CompilerServices.InternalsVisibleTo "ExcelFinancialFunctions.Tests">]
do()

module internal TestPreconditions =
    open System
    open Excel.FinancialFunctions.Common
    open Excel.FinancialFunctions.Tvm
    open Excel.FinancialFunctions.Irr
    open Excel.FinancialFunctions.DayCount

    // Functions to define valid parameter combinations to testcases given that the permutations can create invalid parameters combinations
    let tryPv r nper pmt fv pd =
        ( raisable r nper)                  &&
        ( r <> -1.)                         &&
        ( pmt <> 0. || fv <> 0. )   
    let tryFv r nper pmt pv pd =
        ( raisable r nper)                      &&
        ( r <> -1. || (r = -1. && nper > 0.) )  &&
        ( pmt <> 0. || pv <> 0. )       
    let tryPmt r nper pv fv pd =
        ( raisable r nper)          &&
        ( r > -1. )                 &&
        ( fv <> 0. || pv <> 0. )    &&
        ( annuityCertainPvFactor r nper pd <> 0. )
    let tryNper r pmt pv fv pd =
        ( r > -1.)                                  &&
        ( nperFactor r pmt pv pd <> 0.)             &&
        ( nperFactor r pmt (-fv) pd <> 0.)          &&
        ( nperFactor r pmt (-fv) pd / nperFactor r pmt pv pd > 0.) 
    let tryRate nper pmt pv fv pd guess =
        ( pmt <> 0. || pv <> 0. )                   &&
        ( nper > 0.)                                &&
        not (sign pmt = sign pv && sign pv = sign fv) &&
        not (sign pmt = sign pv && areEqual fv 0.) &&
        not (sign pmt = sign fv && areEqual pv 0.) &&
        not (sign pv = sign fv && areEqual pmt 0.)

    let tryNpv r cfs = r <> -1.   
    let tryIrr cfs = validCfs cfs
    let tryMirr cfs financeRate reinvestRate =
        ( financeRate  <> -1.)      &&
        ( reinvestRate <> -1.)      &&
        ( Seq.length cfs <> 1)      &&
        ( (npv financeRate (cfs |> Seq.map (fun cf -> if cf < 0. then cf else 0.)))  <> 0. ) 
    let tryXirr cfs dates guess =
        validCfs cfs                                        &&
        not(Seq.exists (fun x -> x < Seq.head dates) dates)   &&
        (Seq.length cfs = Seq.length dates)                     
    let tryDb cost salvage life period month =
        ( cost >= 0. )      &&
        ( salvage >= 0. )   &&
        ( life > 0. )       &&
        ( month > 0. )      &&
        ( period <= life )  &&
        ( period > 0. )     &&
        ( month <= 12. )    
    let trySln cost salvage life =
        ( cost >= 0. )      &&
        ( salvage >= 0. )   &&
        ( life > 0. )       
    let trySyd cost salvage life period =
        ( cost >= 0. )      &&
        ( salvage >= 0. )   &&
        ( life > 0. )       &&
        ( period > 0. )     &&
        ( period <= life )  
    let tryDdb cost salvage life period factor =
        ( cost >= 0. )      &&
        ( salvage >= 0. )   &&
        ( life > 0. )       &&
        ( factor > 0. )     &&
        ( period > 0. )     &&
        ( period <= life )  
    let tryVdb cost salvage life startPeriod endPeriod factor bflag =
        ( cost >= 0. )              &&
        ( salvage >= 0. )           &&
        ( life > 0. )               &&
        ( factor > 0. )             &&
        ( startPeriod <= life )     &&
        ( endPeriod <= life )       &&
        ( endPeriod > 0. )          &&
        ( startPeriod <= endPeriod )&&
        ( startPeriod = float (int startPeriod )) && // This is introduced to workaround the issue with fractional startDate
        ( bflag = VdbSwitch.DontSwitchToStraightLine || not(life = startPeriod && startPeriod = endPeriod) )
    let tryIpmt r per nper pv fv pd =
        ( raisable r nper)                              &&
        ( raisable r (per - 1.))                        &&
        ( fv <> 0. || pv <> 0. )                        &&
        ( r > -1. )                                     &&
        ( annuityCertainPvFactor r nper pd <> 0. )      &&
        ( per >= 1. && per <= nper )                    &&
        ( nper > 0. )                       
    let tryPpmt r per nper pv fv pd =
        ( raisable r nper)                              &&
        ( raisable r (per - 1.))                        &&
        ( fv <> 0. || pv <> 0. )                        &&
        ( r > -1.)                                      &&
        ( annuityCertainPvFactor r nper pd <> 0. )      &&
        ( per >= 1. && per <= nper )                    &&
        ( nper > 0. )                       
    let tryCumipmt r nper pv startPeriod endPeriod pd =
        ( raisable r nper)                              &&
        ( raisable r (startPeriod - 1.))                &&
        ( pv > 0. )                                     &&
        ( r > 0. )                                      &&
        ( nper > 0. )                                   &&
        ( annuityCertainPvFactor r nper pd <> 0. )      &&
        ( startPeriod <= endPeriod )                    &&
        ( endPeriod <= nper )                           &&
        ( startPeriod >= 1. )                                                       
    let tryCumprinc r nper pv startPeriod endPeriod pd =
        ( raisable r nper)                              &&
        ( raisable r (startPeriod - 1.))                &&
        ( pv > 0. )                                     &&
        ( r > 0. )                                      &&
        ( nper > 0. )                                   &&
        ( annuityCertainPvFactor r nper pd <> 0. )      &&
        ( startPeriod <= endPeriod )                    &&
        ( endPeriod <= nper )                           &&
        ( startPeriod >= 1. )                           
    let tryIspmt r per nper pv =
        ( per >= 1. && per <= nper )                    &&
        ( nper > 0. )                                                               
    let tryCoupDays (settlement:DateTime) maturity (frequency:Frequency) basis =
        maturity > settlement &&
        [settlement.Year .. maturity.Year] |> List.exists (fun year -> DateTime.IsLeapYear(year)) |> not &&
        settlement.Year <> 1993 // This is to workaround the issue with Coupdays ...
    let tryCoupNum settlement maturity (frequency:Frequency) basis =
        maturity > settlement
    let tryCoupDaysBS settlement maturity (frequency:Frequency) basis =
        maturity > settlement
    let tryCoupDaysNC settlement maturity (frequency:Frequency) basis =
        maturity > settlement
    let tryCoupPCD settlement maturity (frequency:Frequency) basis =
        maturity > settlement
    let tryCoupNCD settlement maturity (frequency:Frequency) basis =
        maturity > settlement

    let tryAccrIntM issue settlement rate par basis =
        (settlement > issue)            &&
        (rate > 0.)                     &&
        (par > 0.)              
    let tryAccrInt issue firstInterest settlement rate par frequency basis  =
        (settlement > issue)            &&
        (firstInterest > settlement)    &&
        (rate > 0.)                     &&
        (par > 0.)
    let tryPrice settlement maturity rate yld redemption (frequency:Frequency) basis =
        (maturity > settlement)         &&
        (rate > 0.)                     &&
        (yld > 0.)                      &&
        (redemption > 0.)               
    let tryYield settlement maturity rate pr redemption (frequency:Frequency) basis =
        (maturity > settlement)         &&
        (rate > 0.)                     &&
        (pr > 0.)                       &&
        (redemption > 0.)               
    let tryPriceMat settlement maturity issue rate yld basis =
        (maturity > settlement)         &&
        (maturity > issue)              &&
        (settlement > issue)            &&
        (rate > 0.)                     &&
        (yld > 0.)                      
    let tryYieldMat settlement maturity issue rate pr basis =
        (maturity > settlement)         &&
        (maturity > issue)              &&
        (settlement > issue)            &&
        (rate > 0.)                     &&
        (pr > 0.)                      
    let tryYearFrac startDate endDate basis =
        startDate < endDate       
    let tryIntRate settlement maturity investment redemption basis =
        (maturity > settlement)         &&
        (investment > 0.)               &&
        (redemption > 0.)               
    let tryReceived settlement maturity investment discount basis =
        let dc = dayCount basis
        let dim = dc.DaysBetween settlement maturity NumDenumPosition.Numerator
        let b = dc.DaysInYear settlement maturity
        let discountFactor = discount * dim / b
        discountFactor < 1.             &&
        (maturity > settlement)         &&
        (investment > 0.)               &&
        (discount > 0.)               
    let tryDisc settlement maturity pr redemption basis =
        (maturity > settlement)         &&
        (pr > 0.)                       &&
        (redemption > 0.)               
    let tryPriceDisc settlement maturity discount redemption basis =
        (maturity > settlement)         &&
        (discount > 0.)                 &&
        (redemption > 0.)               
    let tryYieldDisc settlement maturity pr redemption basis =
        (maturity > settlement)         &&
        (pr > 0.)                       &&
        (redemption > 0.)               
    let tryTBillEq settlement maturity discount =
        let dc = dayCount DayCountBasis.Actual360
        let dsm = dc.DaysBetween settlement maturity NumDenumPosition.Numerator
        let price = (100. - discount * 100. * dsm / 360.) / 100.
        let days = if dsm = 366. then 366. else 365.
        let tempTerm2 = (pow (dsm / days) 2.) - (2. * dsm / days - 1.) * (1. - 1. / price)
        (tempTerm2 >= 0.)                       &&
        (2. * dsm / days - 1. <> 0.)            &&   
        (maturity > settlement)                 &&
        (maturity <= (addYears settlement 1))   &&
        (discount > 0.)                         
    let tryTBillYield settlement maturity pr =
        (maturity > settlement)                 &&
        (maturity <= (addYears settlement 1))   &&
        (pr > 0.)                         
    let tryTBillPrice settlement maturity discount =
        let dc = dayCount DayCountBasis.ActualActual
        let dsm = dc.DaysBetween settlement maturity NumDenumPosition.Numerator
        (100. * (1. - discount * dsm / 360.)) > 0.  &&
        (maturity > settlement)                     &&
        (maturity <= (addYears settlement 1))       &&
        (discount > 0.)                         
    let tryDollarFr fractionalDollar fraction =
        (fraction > 0.) &&
        (pow 10. (ceiling (log10 (floor fraction))) <> 0.) 
    let tryDollarDe fractionalDollar fraction =
        (fraction > 0.)
        
    let tryEffect nominalRate npery =
        (nominalRate > 0.)  &&
        (npery >= 1.)       
    let tryNominal effectRate npery =
        (effectRate > 0.)   &&
        (npery >= 1.)       

    let tryDuration settlement maturity coupon yld frequency basis =
        (maturity > settlement)                 &&
        (coupon >= 0.)                          &&
        (yld >= 0.)                             
    let tryMDuration settlement maturity coupon yld frequency basis =
        (maturity > settlement)                 &&
        (coupon >= 0.)                          &&
        (yld >= 0.)                             

    let tryOddFPrice settlement maturity issue firstCoupon rate yld redemption (frequency:Frequency) basis =
        let (Date(my, mm, md)) = maturity
        let endMonth = lastDayOfMonth my mm md
        let numMonths = int (12. / (float(int frequency)))
        let numMonthsNeg = - numMonths
        let mutable dateT = maturity
        let mutable mat = maturity
        dateT <- changeMonth mat numMonthsNeg basis endMonth
        while dateT > firstCoupon do
            mat <- dateT
            dateT <- changeMonth mat numMonthsNeg basis endMonth
        // This is not in the docs !!!!!
        (dateT = firstCoupon)           &&
        (maturity > firstCoupon)        &&
        (firstCoupon > settlement)      &&
        (settlement > issue)            &&
        (rate >= 0.)                    &&
        (yld >= 0.)                     &&
        (redemption >= 0.)              

    let tryCalcOddFYield settlement maturity issue firstCoupon rate pr redemption (frequency:Frequency) basis =
        (maturity > firstCoupon)        &&
        (firstCoupon > settlement)      &&
        (settlement > issue)            &&
        (rate >= 0.)                    &&
        (pr >= 0.)                      &&
        (redemption >= 0.)              
    let tryOddLPrice settlement maturity lastInterest rate yld redemption (frequency:Frequency) basis =
        (maturity > settlement)         &&
        (settlement > lastInterest)     &&
        (rate >= 0.)                    &&
        (yld >= 0.)                     &&
        (redemption >= 0.)              
    let tryOddLYield settlement maturity lastInterest rate pr redemption (frequency:Frequency) basis =
        (maturity > settlement)         &&
        (settlement > lastInterest)     &&
        (rate >= 0.)                    &&
        (pr >= 0.)                      &&
        (redemption >= 0.)              
    let tryAmorLinc cost datePurchased firstPeriod salvage period rate basis =
        ( cost >= 0. )                      &&
        ( salvage >= 0. )                   &&
        ( salvage < cost )                  &&
        ( period >= 0. )                    &&
        ( rate >= 0. )                      &&
        (datePurchased < firstPeriod)       &&
        (basis <> DayCountBasis.Actual360 )  
    let tryAmorDegrc cost datePurchased firstPeriod salvage period rate basis =
        let assetLife = 1. / rate
        let between x1 x2 = assetLife >= x1 && assetLife <= x2
        ( not(between 0. 3.) )              &&
        ( not(between 4. 5.) )              &&
        ( cost >= 0. )                      &&
        ( salvage >= 0. )                   &&
        ( salvage < cost )                  &&
        ( period >= 0. )                    &&
        ( rate >= 0. )                      &&
        (datePurchased < firstPeriod)       &&
        (basis <> DayCountBasis.Actual360 )
