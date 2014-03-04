// Bonds mathematics the Excel way 
#light
namespace Excel.FinancialFunctions

open Excel.FinancialFunctions.Common
open Excel.FinancialFunctions.DayCount

module internal Bonds =

    // Main formulas
         
    let accrIntM issue settlement rate par basis =
         let dc = dayCount basis
         let days = dc.DaysBetween issue settlement NumDenumPosition.Numerator
         let daysInYear = dc.DaysInYear issue settlement
         par * rate * (days/daysInYear)         
    let accrInt issue (Date(fiy, fim, fid) as firstInterest) settlement rate par (frequency:Frequency) basis calcMethod =
        let dc = dayCount basis 
        let freq = float (int frequency)
        let numMonths = freq2months freq
        let numMonthsNeg = - numMonths
        let endMonthBond = lastDayOfMonth fiy fim fid
        let pcd =
            if settlement > firstInterest && calcMethod = AccrIntCalcMethod.FromIssueToSettlement
            then findPcdNcd firstInterest settlement numMonths basis endMonthBond |> fst                
            else dc.ChangeMonth firstInterest numMonthsNeg endMonthBond
        let firstDate = if issue > pcd then issue else pcd
        let days = dc.DaysBetween firstDate settlement NumDenumPosition.Numerator
        let coupDays = dc.CoupDays pcd firstInterest freq
        let aggFunction pcd ncd =
            let firstDate = if issue > pcd then issue else pcd
            let days =
                if basis = DayCountBasis.UsPsa30_360
                then
                    let psaMethod = if issue > pcd then Method360Us.ModifyStartDate else Method360Us.ModifyBothDates
                    float (dateDiff360Us firstDate ncd psaMethod)
                else dc.DaysBetween firstDate ncd NumDenumPosition.Numerator
            let coupDays =
                if basis = DayCountBasis.UsPsa30_360
                then float (dateDiff360Us pcd ncd Method360Us.ModifyBothDates)
                elif basis = DayCountBasis.Actual365 then 365. / freq
                else dc.DaysBetween pcd ncd NumDenumPosition.Denumerator
            if issue <= pcd then float (int calcMethod) else days / coupDays
        let _, _, a = datesAggregate1 pcd issue numMonthsNeg basis aggFunction (days / coupDays) endMonthBond
        par * rate / freq * a
    
    let getPriceYieldFactors settlement maturity frequency basis =
        let dc = dayCount basis 
        let n = dc.CoupNum settlement maturity frequency 
        let pcd = dc.CoupPCD settlement maturity frequency 
        let a = dc.DaysBetween pcd settlement NumDenumPosition.Numerator
        let e = dc.CoupDays settlement maturity frequency 
        let dsc = e - a
        n, pcd, a, e, dsc            
    let price settlement maturity rate yld redemption frequency basis =
        let n, pcd, a, e, dsc = getPriceYieldFactors settlement maturity frequency basis
        let coupon = 100. * rate / frequency
        let accrInt = 100. * rate / frequency * a / e
        let pvFactor k = pow (1. + yld / frequency) (k - 1. + dsc / e)
        let pvOfRedemption = redemption / pvFactor n
        let mutable pvOfCoupons = 0.
        for k = 1 to int n do pvOfCoupons <- pvOfCoupons + coupon / pvFactor (float k)
        if n = 1. then
            (redemption + coupon) / (1. + dsc / e * yld / frequency) - accrInt  
        else
            pvOfRedemption + pvOfCoupons - accrInt          
    let yieldFunc settlement maturity rate pr redemption frequency basis =
        let n, pcd, a, e, dsr = getPriceYieldFactors settlement maturity frequency basis
        if n <= 1. then
            let par = 100. // Logical inference from Excel's docs
            let firstNum = (redemption / 100. + rate / frequency) - (par / 100. + (a / e * rate /frequency))
            let firstDen = par / 100. + (a / e * rate / frequency)
            firstNum / firstDen * frequency * e / dsr
        else
            findRoot (fun yld -> price settlement maturity rate yld redemption frequency basis - pr) 0.05
    let getMatFactors settlement maturity issue basis =
        let dc = dayCount basis 
        let b = dc.DaysInYear issue settlement
        let dim = dc.DaysBetween issue maturity NumDenumPosition.Numerator
        let a = dc.DaysBetween issue settlement NumDenumPosition.Numerator
        let dsm = dim - a
        b, dim, a, dsm    
    let priceMat settlement maturity issue rate yld basis =
        let b, dim, a, dsm = getMatFactors settlement maturity issue basis 
        let num1 = 100. + (dim / b * rate * 100.)
        let den1 = 1. + (dsm / b * yld)
        let fact2 = (a / b * rate * 100.)
        num1 / den1 - fact2
    let yieldMat settlement maturity issue rate pr basis =
        let b, dim, a, dsm = getMatFactors settlement maturity issue basis 
        let term1 = dim / b * rate + 1. - pr / 100. - a / b * rate
        let term2 = pr / 100. + a / b * rate
        let term3 = b / dsm
        term1 / term2 * term3 
    let getCommonFactors settlement maturity basis =
        let dc = dayCount basis
        let dim = dc.DaysBetween settlement maturity NumDenumPosition.Numerator
        let b = dc.DaysInYear settlement maturity
        dim, b           
    let intRate settlement maturity investment redemption basis =
        let dim, b = getCommonFactors settlement maturity basis
        (redemption - investment) / investment * b /dim
    let received settlement maturity investment discount basis =
        let dim, b = getCommonFactors settlement maturity basis
        let discountFactor = discount * dim / b
        // To get the following check into the precondition testing requires calculating the discountFactor twice, so I don't do it ... 
        // discountFactor < 1.   |> elseThrow "discount * dim / b must be different from 1"
        investment / ( 1. - discountFactor )
    let disc settlement maturity pr redemption basis =
        let dim, b = getCommonFactors settlement maturity basis
        (- pr / redemption + 1.) * b / dim
    let priceDisc settlement maturity discount redemption basis =
        let dim, b = getCommonFactors settlement maturity basis
        redemption - discount * redemption * dim / b
    let yieldDisc settlement maturity pr redemption basis =
        let dim, b = getCommonFactors settlement maturity basis
        (redemption - pr) / pr * b / dim
    let duration settlement maturity coupon yld frequency basis isMDuration =
        let dc = dayCount basis 
        let dbc = dc.CoupDaysBS settlement maturity frequency
        let e = dc.CoupDays settlement maturity frequency
        let n = dc.CoupNum settlement maturity frequency
        let dsc = e - dbc
        let x1 = dsc / e
        let x2 = x1 + n - 1.
        let x3 = yld / frequency + 1.
        let x4 = pow x3 x2
        ( x4 <> 0.) |> elseThrow "(yld / frequency + 1)^((dsc / e) + n -1) must be different from 0)"
        let term1 = x2 * 100. / x4
        let term3 = 100. / x4
        let aggrFunction acc index =
            let x5 = float index - 1. + x1
            let x6 = pow x3 x5
            ( x6 <> 0.) |> elseThrow "x6 must be different from 0)"
            let x7 = (100. * coupon / frequency) / x6
            let a, b = acc
            a + x7 * x5, b + x7
        let term2, term4 = aggrBetween 1 (int n) aggrFunction (0., 0.)
        
        let term5 = term1 + term2
        let term6 = term3 + term4
        ( term6 <> 0.) |> elseThrow "term6 must be different from 0)"
        if not(isMDuration) then (term5 / term6) / frequency else ((term5 / term6) / frequency) / x3
               
    // Preconditions and special cases
    let calcAccrIntM issue settlement rate par (basis:DayCountBasis) =
        (settlement > issue)    |> elseThrow "settlement must be after issue"
        (rate > 0.)             |> elseThrow "rate must be more than 0"
        (par > 0.)              |> elseThrow "par must be more than 0"
        accrIntM issue settlement rate par basis
    let calcAccrInt issue firstInterest settlement rate par (frequency:Frequency) basis (calcMethod:AccrIntCalcMethod) =
        (settlement > issue)            |> elseThrow "settlement must be after issue"
        (firstInterest > settlement)    |> elseThrow "firstInterest must be after settlement"
        (rate > 0.)                     |> elseThrow "rate must be more than 0"
        (par > 0.)                      |> elseThrow "par must be more than 0"
        accrInt issue firstInterest settlement rate par frequency basis calcMethod
    let calcPrice settlement maturity rate yld redemption (frequency:Frequency) basis =
        (maturity > settlement)         |> elseThrow "maturity must be after settlement"
        (rate > 0.)                     |> elseThrow "rate must be more than 0"
        (yld > 0.)                      |> elseThrow "yld must be more than 0"
        (redemption > 0.)               |> elseThrow "redemption must be more than 0"
        price settlement maturity rate yld redemption (float (int frequency)) basis
    let calcYield settlement maturity rate pr redemption (frequency:Frequency) basis =
        (maturity > settlement)         |> elseThrow "maturity must be after settlement"
        (rate > 0.)                     |> elseThrow "rate must be more than 0"
        (pr > 0.)                       |> elseThrow "pr must be more than 0"
        (redemption > 0.)               |> elseThrow "redemption must be more than 0"
        yieldFunc settlement maturity rate pr redemption (float (int frequency)) basis        
    let calcPriceMat settlement maturity issue rate yld basis =
        (maturity > settlement)         |> elseThrow "maturity must be after settlement"
        (maturity > issue)              |> elseThrow "maturity must be after issue"
        (settlement > issue)            |> elseThrow "settlement must be after issue"
        (rate > 0.)                     |> elseThrow "rate must be more than 0"
        (yld > 0.)                      |> elseThrow "yld must be more than 0"
        priceMat settlement maturity issue rate yld basis
    let calcYieldMat settlement maturity issue rate pr basis =
        (maturity > settlement)         |> elseThrow "maturity must be after settlement"
        (maturity > issue)              |> elseThrow "maturity must be after issue"
        (settlement > issue)            |> elseThrow "settlement must be after issue"
        (rate > 0.)                     |> elseThrow "rate must be more than 0"
        (pr > 0.)                       |> elseThrow "price must be more than 0"
        yieldMat settlement maturity issue rate pr basis
    let calcIntRate settlement maturity investment redemption basis =
        (maturity > settlement)         |> elseThrow "maturity must be after settlement"
        (investment > 0.)               |> elseThrow "investment must be more than 0"
        (redemption > 0.)               |> elseThrow "redemption must be more than 0"
        intRate settlement maturity investment redemption basis
    let calcReceived settlement maturity investment discount basis =
        let dc = dayCount basis
        let dim = dc.DaysBetween settlement maturity NumDenumPosition.Numerator
        let b = dc.DaysInYear settlement maturity
        let discountFactor = discount * dim / b
        (discountFactor < 1.)           |> elseThrow "discount * dim / b must be different from 1"
        (maturity > settlement)         |> elseThrow "maturity must be after settlement"
        (investment > 0.)               |> elseThrow "investment must be more than 0"
        (discount > 0.)                 |> elseThrow "redemption must be more than 0"
        received settlement maturity investment discount basis
    let calcDisc settlement maturity pr redemption basis =
        (maturity > settlement)         |> elseThrow "maturity must be after settlement"
        (pr > 0.)                       |> elseThrow "investment must be more than 0"
        (redemption > 0.)               |> elseThrow "redemption must be more than 0"
        disc settlement maturity pr redemption basis
    let calcPriceDisc settlement maturity discount redemption basis =
        (maturity > settlement)         |> elseThrow "maturity must be after settlement"
        (discount > 0.)                 |> elseThrow "investment must be more than 0"
        (redemption > 0.)               |> elseThrow "redemption must be more than 0"
        priceDisc settlement maturity discount redemption basis
    let calcYieldDisc settlement maturity pr redemption basis =
        (maturity > settlement)         |> elseThrow "maturity must be after settlement"
        (pr > 0.)                       |> elseThrow "investment must be more than 0"
        (redemption > 0.)               |> elseThrow "redemption must be more than 0"
        yieldDisc settlement maturity pr redemption basis
    let calcDuration settlement maturity coupon yld (frequency:Frequency) basis =
        (maturity > settlement)                 |> elseThrow "maturity must be after settlement"
        (coupon >= 0.)                          |> elseThrow "coupon must be more than 0"
        (yld >= 0.)                             |> elseThrow "yld must be more than 0"
        duration settlement maturity coupon yld (float (int frequency)) basis false
    let calcMDuration settlement maturity coupon yld (frequency:Frequency) basis =
        (maturity > settlement)                 |> elseThrow "maturity must be after settlement"
        (coupon >= 0.)                          |> elseThrow "coupon must be more than 0"
        (yld >= 0.)                             |> elseThrow "yld must be more than 0"
        duration settlement maturity coupon yld (float (int frequency)) basis true
         
 