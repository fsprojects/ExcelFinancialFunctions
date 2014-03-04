// Difficult to digest formulas for odd bonds calculations. Trust the testcases that I got these ones right.
#light
namespace Excel.FinancialFunctions

open System
open Excel.FinancialFunctions.Common
open Excel.FinancialFunctions.DayCount

module internal OddBonds =

    // Main formulas
    let coupNumber (Date(my, mm, md) as mat) (Date(sy, sm, sd) as settl) numMonths basis isWholeNumber =
        let couponsTemp = if isWholeNumber then 0. else 1.
        let endOfMonthTemp = lastDayOfMonth my mm md
        let endOfMonth = if not(endOfMonthTemp) && mm <> 2 && md > 28 && md < daysOfMonth my mm then lastDayOfMonth sy sm sd else endOfMonthTemp
        let startDate = changeMonth settl 0 basis endOfMonth
        let coupons = if settl < startDate then couponsTemp + 1. else couponsTemp
        let date = changeMonth startDate numMonths basis endOfMonth
        let _, _, result = datesAggregate1 date mat numMonths basis (fun pcd ncd -> 1.) coupons endOfMonth
        result
    
    let daysBetweenNotNeg (dc:IDayCount) startDate endDate =
        let result = dc.DaysBetween startDate endDate NumDenumPosition.Numerator
        if result > 0. then result else 0.
    let daysBetweenNotNegPsaHack startDate endDate =
        let result = float (dateDiff360Us startDate endDate Method360Us.ModifyBothDates)
        if result > 0. then result else 0.
    let daysBetweenNotNegWithHack dc startDate endDate basis =
        if basis = DayCountBasis.UsPsa30_360
        then daysBetweenNotNegPsaHack startDate endDate
        else daysBetweenNotNeg dc startDate endDate
            
    let oddFPrice settlement (Date(my, mm, md) as maturity) issue firstCoupon rate yld redemption frequency basis =
        let dc = dayCount basis 
        let endMonth = lastDayOfMonth my mm md
        let numMonths = freq2months frequency
        let numMonthsNeg = - numMonths
        let e = dc.CoupDays settlement firstCoupon frequency
        let n = dc.CoupNum settlement maturity frequency
        let m = frequency
        let dfc = daysBetweenNotNeg dc issue firstCoupon
        if dfc < e then
            let dsc = daysBetweenNotNeg dc settlement firstCoupon
            let a = daysBetweenNotNeg dc issue settlement
            let x = yld / m + 1.
            let y = dsc / e
            let p1 = x
            let p3 = pow p1 (n - 1. + y)
            let term1 = redemption / p3
            let term2 = 100. * rate / m * dfc / e / pow p1  y
            let term3 = aggrBetween 2 (int n) (fun acc index -> acc + 100. * rate / m / (pow p1 (float index - 1. + y))) 0.
            let p2 = rate / m
            let term4 = a / e * p2 * 100.
            term1 + term2 + term3 - term4
        else // dfc >= e
            let nc = dc.CoupNum issue firstCoupon frequency
            let lateCoupon = ref firstCoupon            
            let aggrFunction acc index =
                let earlyCoupon = changeMonth !lateCoupon numMonthsNeg basis false
                let nl =
                    if basis = DayCountBasis.ActualActual
                    then daysBetweenNotNeg dc earlyCoupon !lateCoupon
                    else e
                let dci = if index > 1 then nl else daysBetweenNotNeg dc issue !lateCoupon
                let startDate = if issue > earlyCoupon then issue else earlyCoupon
                let endDate = if settlement < !lateCoupon then settlement else !lateCoupon
                let a = daysBetweenNotNeg dc startDate endDate
                lateCoupon := earlyCoupon
                let dcnl, anl = acc
                dcnl + dci / nl, anl + a / nl
            let dcnl, anl = aggrBetween (int nc) 1 aggrFunction (0., 0.)                         
            let dsc =
                if basis = DayCountBasis.Actual360 || basis = DayCountBasis.Actual365
                then
                    let date = dc.CoupNCD settlement firstCoupon frequency
                    daysBetweenNotNeg dc settlement date
                else
                    let date = dc.CoupPCD settlement firstCoupon frequency
                    let a = dc.DaysBetween date settlement NumDenumPosition.Numerator
                    e - a
            let nq = coupNumber firstCoupon settlement numMonths basis true
            let n = dc.CoupNum firstCoupon maturity frequency
            let x = yld / m + 1.
            let y = dsc / e
            let p1 = x
            let p3 = pow p1 (y + nq + n)
            let term1 = redemption / p3
            let term2 = 100. * rate / m * dcnl / pow p1 (nq + y)
            let term3 = aggrBetween 1 (int n) (fun acc index -> acc + 100. * rate / m / (pow p1 (float index + nq + y))) 0.
            let term4 = 100. * rate / m * anl
            term1 + term2 + term3 - term4 
    let oddFYield settlement maturity issue firstCoupon rate pr redemption frequency basis =
        let dc = dayCount basis
        let years = dc.DaysBetween settlement maturity NumDenumPosition.Numerator
        let m = frequency
        let px = pr - 100.
        let num = rate * years * 100. - px
        let denum = px / 4. + years * px / 2. + years * 100.
        let guess = num / denum
        findRoot (fun yld -> pr - oddFPrice settlement maturity issue firstCoupon rate yld redemption frequency basis) guess                      
    let oddLFunc settlement maturity lastInterest rate prOrYld redemption frequency basis isLPrice =
        let dc = dayCount basis
        let m = frequency
        let numMonths = int (12. / frequency)
        let lastCoupon = lastInterest
        let nc = dc.CoupNum lastCoupon maturity frequency
        let earlyCoupon = ref lastCoupon
        let aggrFunction acc index =
            let lateCoupon = changeMonth !earlyCoupon numMonths basis false          
            let nl = daysBetweenNotNegWithHack dc !earlyCoupon lateCoupon basis
            let dci = if index < int nc then nl else daysBetweenNotNegWithHack dc !earlyCoupon maturity basis
            let a =
                if lateCoupon < settlement
                then dci
                elif !earlyCoupon < settlement
                    then daysBetweenNotNeg dc !earlyCoupon settlement
                else 0.
            let startDate = if settlement > !earlyCoupon then settlement else !earlyCoupon
            let endDate = if maturity < lateCoupon then maturity else lateCoupon
            let dsc = daysBetweenNotNeg dc startDate endDate
            earlyCoupon := lateCoupon
            let dcnl, anl, dscnl = acc
            dcnl + dci / nl, anl + a / nl , dscnl + dsc / nl        
        let dcnl, anl, dscnl = aggrBetween 1 (int nc) aggrFunction (0., 0., 0.)
        let x = 100. * rate / m
        let term1 = dcnl * x + redemption
        if isLPrice then
            let term2 = dscnl * prOrYld / m + 1.
            let term3 = anl * x
            term1 / term2 - term3
        else
            let term2 = anl * x + prOrYld    
            let term3 = m / dscnl
            (term1 - term2) / term2 * term3
        
    // Preconditions and special cases
    let calcOddFPrice settlement (Date(my, mm, md) as maturity) issue firstCoupon rate yld redemption (frequency:Frequency) basis =
        let endMonth = lastDayOfMonth my mm md
        let numMonths = int (12. / (float(int frequency)))
        let numMonthsNeg = - numMonths
        let pcd, ncd = findPcdNcd (changeMonth maturity numMonthsNeg basis endMonth) firstCoupon numMonthsNeg basis endMonth
        // The next condition is not in the docs, but nevertheless is needed !!!!!
        (pcd = firstCoupon)             |> elseThrow "maturity and firstCoupon must have the same month and day (except for February when leap years are considered)"
        (maturity > firstCoupon)        |> elseThrow "maturity must be after firstCoupon"
        (firstCoupon > settlement)      |> elseThrow "firstCoupon must be after settlement"
        (settlement > issue)            |> elseThrow "settlement must be after issue"
        (rate >= 0.)                    |> elseThrow "rate must be more than 0"
        (yld >= 0.)                     |> elseThrow "yld must be more than 0"
        (redemption >= 0.)              |> elseThrow "redemption must be more than 0"
        oddFPrice settlement maturity issue firstCoupon rate yld redemption (float (int frequency)) basis
    let calcOddFYield settlement maturity issue firstCoupon rate pr redemption (frequency:Frequency) basis =
        (maturity > firstCoupon)        |> elseThrow "maturity must be after firstCoupon"
        (firstCoupon > settlement)      |> elseThrow "firstCoupon must be after settlement"
        (settlement > issue)            |> elseThrow "settlement must be after issue"
        (rate >= 0.)                    |> elseThrow "rate must be more than 0"
        (pr >= 0.)                      |> elseThrow "pr must be more than 0"
        (redemption >= 0.)              |> elseThrow "redemption must be more than 0"
        oddFYield settlement maturity issue firstCoupon rate pr redemption (float (int frequency)) basis
    let calcOddLPrice settlement maturity lastInterest rate yld redemption (frequency:Frequency) basis =
        (maturity > settlement)         |> elseThrow "maturity must be after settlement"
        (settlement > lastInterest)     |> elseThrow "settlement must be after lastInterest"
        (rate >= 0.)                    |> elseThrow "rate must be more than 0"
        (yld >= 0.)                     |> elseThrow "yld must be more than 0"
        (redemption >= 0.)              |> elseThrow "redemption must be more than 0"
        oddLFunc settlement maturity lastInterest rate yld redemption (float (int frequency)) basis true
    let calcOddLYield settlement maturity lastInterest rate pr redemption (frequency:Frequency) basis =
        (maturity > settlement)         |> elseThrow "maturity must be after settlement"
        (settlement > lastInterest)     |> elseThrow "settlement must be after lastInterest"
        (rate >= 0.)                    |> elseThrow "rate must be more than 0"
        (pr >= 0.)                      |> elseThrow "pr must be more than 0"
        (redemption >= 0.)              |> elseThrow "redemption must be more than 0"
        oddLFunc settlement maturity lastInterest rate pr redemption (float (int frequency)) basis false
       
   