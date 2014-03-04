// Loan related calculation routines, small variations on TVM
#light
namespace Excel.FinancialFunctions

open System
open Excel.FinancialFunctions.Common
open Excel.FinancialFunctions.Tvm

module internal Loan =
    let inline approxEqual x y = abs (x - y) < 1e-10
    
    // Main formulas
    let ipmt r per nper pv fv pd = 
        let result = - ( pv * fvFactor r (per - 1.) * r + (pmt r nper pv fv PaymentDue.EndOfPeriod) * (fvFactor r (per - 1.) - 1.) )
        if pd = PaymentDue.EndOfPeriod then result else result / (1. + r)
    let ppmt r per nper pv fv pd =
        pmt r nper pv fv pd - ipmt r per nper pv fv pd
    let ispmt r per nper pv =
        let coupon = - pv * r
        coupon - (coupon / nper * per)

    // Preconditions and special cases
    let calcIpmt r per nper pv fv pd =
        ( raisable r nper)                                                          |> elseThrow "r is not raisable to nper (r is negative and nper not an integer)"
        ( raisable r (per - 1.))                                                    |> elseThrow "r is not raisable to (per - 1) (r is negative and nper not an integer)"
        ( fv <> 0. || pv <> 0. )                                                    |> elseThrow "fv or pv need to be different from 0"
        ( r > -1.)                                                                  |> elseThrow "r must be more than -100%"
        ( annuityCertainPvFactor r nper pd <> 0. )                                  |> elseThrow "1 * pd + 1 - (1 / (1 + r)^nper) / nper has to be <> 0"
        ( per >= 1. && per <= nper )                                                |> elseThrow "per must be in the range 1 to nper"
        ( nper > 0. )                                                               |> elseThrow "nper must be more than 0"
        if approxEqual per 1. && pd = PaymentDue.BeginningOfPeriod then 0.
        elif r = -1. then -fv
        else ipmt r per nper pv fv pd
    let calcPpmt r per nper pv fv pd =
        ( raisable r nper)                                                          |> elseThrow "r is not raisable to nper (r is negative and nper not an integer)"
        ( raisable r (per - 1.))                                                    |> elseThrow "r is not raisable to (per - 1) (r is negative and nper not an integer)"
        ( fv <> 0. || pv <> 0. )                                                    |> elseThrow "fv or pv need to be different from 0"
        ( r > -1. )                                                                 |> elseThrow "r must be more than -100%"
        ( annuityCertainPvFactor r nper pd <> 0. )                                  |> elseThrow "1 * pd + 1 - (1 / (1 + r)^nper) / nper has to be <> 0"
        ( per >= 1. && per <= nper )                                                |> elseThrow "per must be in the range 1 to nper"
        ( nper > 0. )                                                               |> elseThrow "nper must be more than 0"
        if approxEqual per 1. && pd = PaymentDue.BeginningOfPeriod then pmt r nper pv fv pd
        elif r = -1. then 0.
        else ppmt r per nper pv fv pd       
    let calcCumipmt r nper pv startPeriod endPeriod pd =
        ( raisable r nper)                                                          |> elseThrow "r is not raisable to nper (r is negative and nper not an integer)"
        ( raisable r (startPeriod - 1.))                                            |> elseThrow "r is not raisable to (per - 1) (r is negative and nper not an integer)"
        ( pv > 0. )                                                                 |> elseThrow "pv must be more than 0"
        ( r > 0. )                                                                  |> elseThrow "r must be more than 0"
        ( nper > 0. )                                                               |> elseThrow "nper must be more than 0"
        ( annuityCertainPvFactor r nper pd <> 0. )                                  |> elseThrow "1 * pd + 1 - (1 / (1 + r)^nper) / nper has to be <> 0"
        ( startPeriod <= endPeriod )                                                |> elseThrow "startPeriod must be less or equal to endPeriod"
        ( startPeriod >= 1. )                                                       |> elseThrow "startPeriod must be more or equal to 1"
        ( endPeriod <= nper )                                                       |> elseThrow "startPeriod and endPeriod must be less or equal to nper"
        aggrBetween (int (ceiling startPeriod)) (int endPeriod) (fun acc per -> acc + calcIpmt r (float per) nper pv 0. pd) 0.
    let calcCumprinc r nper pv startPeriod endPeriod pd =
        ( raisable r nper)                                                          |> elseThrow "r is not raisable to nper (r is negative and nper not an integer)"
        ( raisable r (startPeriod - 1.))                                            |> elseThrow "r is not raisable to (per - 1) (r is negative and nper not an integer)"
        ( pv > 0. )                                                                 |> elseThrow "pv must be more than 0"
        ( r > 0. )                                                                  |> elseThrow "r must be more than 0"
        ( nper > 0. )                                                               |> elseThrow "nper must be more than 0"
        ( annuityCertainPvFactor r nper pd <> 0. )                                  |> elseThrow "1 * pd + 1 - (1 / (1 + r)^nper) / nper has to be <> 0"
        ( startPeriod <= endPeriod )                                                |> elseThrow "startPeriod must be less or equal to endPeriod"
        ( startPeriod >= 1. )                                                       |> elseThrow "startPeriod must be more or equal to 1"
        ( endPeriod <= nper )                                                       |> elseThrow "startPeriod and endPeriod must be less or equal to nper"
        aggrBetween (int (ceiling startPeriod)) (int endPeriod) (fun acc per -> acc + calcPpmt r (float per) nper pv 0. pd) 0.
    let calcIspmt r per nper pv =
        ( per >= 1. && per <= nper )                                                |> elseThrow "per must be in the range 1 to nper"
        ( nper > 0. )                                                               |> elseThrow "nper must be more than 0"
        ispmt r per nper pv