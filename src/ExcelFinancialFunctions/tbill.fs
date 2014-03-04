// Very simple TBill mathematics. The only interesting thing is the 'if' statement on line 17.
#light
namespace Excel.FinancialFunctions

open System
open Excel.FinancialFunctions.Common
open Excel.FinancialFunctions.DayCount

module internal TBill =
    
    // Main formulas
    let getDsm settlement maturity basis = 
        let dc = dayCount basis
        dc.DaysBetween settlement maturity NumDenumPosition.Numerator
    let TBillEq settlement maturity discount =
        let dsm = getDsm settlement maturity DayCountBasis.Actual360
        if dsm > 182. then
            let price = (100. - discount * 100. * dsm / 360.) / 100.
            let days = if dsm = 366. then 366. else 365.
            let tempTerm2 = (pow (dsm / days) 2.) - (2. * dsm / days - 1.) * (1. - 1. / price)
            let term2 = sqr tempTerm2
            let term3 = 2. * dsm / days - 1.
            2. * (term2 - dsm / days) / term3
        else
            // This is the algo in the docs, but it is valid just above 182 ...
            365. * discount / (360. - discount * dsm)
    let TBillYield settlement maturity pr =
        let dsm = getDsm settlement maturity DayCountBasis.ActualActual
        (100. - pr) / pr * 360. / dsm
    let TBillPrice settlement maturity discount =
        let dsm = getDsm settlement maturity DayCountBasis.ActualActual
        100. * (1. - discount * dsm / 360.)
                
    // Preconditions and special cases
    let calcTBillEq settlement maturity discount =
        let dsm = getDsm settlement maturity DayCountBasis.Actual360
        let price = (100. - discount * 100. * dsm / 360.) / 100.
        let days = if dsm = 366. then 366. else 365.
        let tempTerm2 = (pow (dsm / days) 2.) - (2. * dsm / days - 1.) * (1. - 1. / price)
        (tempTerm2 >= 0.)                       |> elseThrow "(dsm / days)^2 - (2. * dsm / days - 1.) * (1. - 1. / (100. - discount * 100. * dsm / 360.) / 100.) must be positive"
        (2. * dsm / days - 1. <> 0.)            |> elseThrow "2. * dsm / days - 1. must be different from 0"    
        (maturity > settlement)                 |> elseThrow "maturity must be after settlement"
        (maturity <= (addYears settlement 1))   |> elseThrow "maturity must be less than one year after settlement"
        (discount > 0.)                         |> elseThrow "investment must be more than 0"
        TBillEq settlement maturity discount
    let calcTBillYield settlement maturity pr =
        (maturity > settlement)                 |> elseThrow "maturity must be after settlement"
        (maturity <= (addYears settlement 1))   |> elseThrow "maturity must be less than one year after settlement"
        (pr > 0.)                               |> elseThrow "pr must be more than 0"
        TBillYield settlement maturity pr
    let calcTBillPrice settlement maturity discount =
        let dsm = getDsm settlement maturity DayCountBasis.ActualActual
        (100. * (1. - discount * dsm / 360.)) > 0. |> elseThrow "a result less than zero triggers an exception"
        (maturity > settlement)                 |> elseThrow "maturity must be after settlement"
        (maturity <= (addYears settlement 1))   |> elseThrow "maturity must be less than one year after settlement"
        (discount > 0.)                         |> elseThrow "discount must be more than 0"
        TBillPrice settlement maturity discount
    


