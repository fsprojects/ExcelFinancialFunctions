// Time value of money routines, note the extensive treatment of error condition to help the user with sensible error messages
#light
namespace Excel.FinancialFunctions

open System
open Excel.FinancialFunctions.Common

module internal Tvm =
        
    // Main formulas       
    let fvFactor r nper = (1. + r) ** nper
    let pvFactor r nper = 1. / fvFactor r nper
    let annuityCertainPvFactor r nper (pd:PaymentDue) = if r = 0. then nper else (1. + r * float pd) * (1. - pvFactor r nper) / r 
    let annuityCertainFvFactor r nper (pd:PaymentDue) = annuityCertainPvFactor r nper pd * fvFactor r nper
    let nperFactor r pmt v (pd:PaymentDue) = v * r + pmt * ( 1. + r * float pd )

    let pv r nper pmt fv pd = - (fv * pvFactor r nper + pmt * annuityCertainPvFactor r nper pd)
    let fv r nper pmt pv pd = - (pv * fvFactor r nper + pmt * annuityCertainFvFactor r nper pd)
    let pmt r nper pv fv pd = - (pv + fv * pvFactor r nper) / annuityCertainPvFactor r nper pd
    let nper r pmt pv fv (pd:PaymentDue) = ln ( nperFactor r pmt (-fv) pd / nperFactor r pmt pv pd ) / ln (r+1.)
             
    // Preconditions and special cases    
    let calcPv r nper pmt fv pd =
        ( raisable r nper)          |> elseThrow "r is not raisable to nper (r is less than -1 and nper not an integer"
        ( pmt <> 0. || fv <> 0. )   |> elseThrow "pmt or fv need to be different from 0"
        ( r <> -1.)                 |> elseThrow "r cannot be -100%"
        pv r nper pmt fv pd
    let calcFv r nper pmt pv pd =
        ( raisable r nper)                          |> elseThrow "r is not raisable to nper (r is negative and nper not an integer"
        ( r <> -1. || (r = -1. && nper > 0.) )      |> elseThrow "r cannot be -100% when nper is <= 0"
        ( pmt <> 0. || pv <> 0. )                   |> elseThrow "pmt or pv need to be different from 0"
        if r = -1. && pd = PaymentDue.BeginningOfPeriod then - (pv * fvFactor r nper)
        elif r = -1. && pd = PaymentDue.EndOfPeriod then - (pv * fvFactor r nper + pmt)
        else fv r nper pmt pv pd
    let calcPmt r nper pv fv pd =
        ( raisable r nper)                                                          |> elseThrow "r is not raisable to nper (r is negative and nper not an integer"
        ( fv <> 0. || pv <> 0. )                                                    |> elseThrow "fv or pv need to be different from 0"
        ( r <> -1. || (r = -1. && nper > 0. && pd = PaymentDue.EndOfPeriod) )       |> elseThrow "r cannot be -100% when nper is <= 0"
        ( annuityCertainPvFactor r nper pd <> 0. )                                  |> elseThrow "1 * pd + 1 - (1 / (1 + r)^nper) / nper has to be <> 0"
        if r = -1. then -fv
        else pmt r nper pv fv pd
    let calcNper r pmt pv fv pd =
        if r = 0. && pmt <> 0. then
            - (fv + pv) / pmt
        else
            nper r pmt pv fv pd
    let calcRate nper pmt pv fv pd guess =
        let haveRightSigns x y z =
            not( sign x = sign y && sign y = sign z) &&
            not (sign x = sign y && z = 0.) &&
            not (sign x = sign z && y = 0.) &&
            not (sign y = sign z && x = 0.)
            
        ( pmt <> 0. || pv <> 0. )                   |> elseThrow "pmt or pv need to be different from 0"
        ( nper > 0.)                                |> elseThrow "nper needs to be more than 0"
        ( haveRightSigns pmt pv fv )                |> elseThrow "There must be at least a change in sign in pv, fv and pmt"
        
        if fv = 0. && pv = 0. then
            if pmt < 0. then -1. else 1.
        else
            let f = fun r -> calcFv r nper pmt pv pd - fv            
            findRoot f guess
    let calcFvSchedule (pv:float) interests =
        let mutable result = pv
        for i in interests do result <- result * (1. + i)
        result