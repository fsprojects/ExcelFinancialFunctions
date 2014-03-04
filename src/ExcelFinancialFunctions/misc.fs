// Various routings that don't have an obvious classification in other files.
#light
namespace Excel.FinancialFunctions

open System
open Excel.FinancialFunctions.Common

module internal Misc =
    
    // Main formulas
    let dollar fractionalDollar fraction f =
        let  aBase = floor fraction
        let dollar = if fractionalDollar > 0. then floor fractionalDollar else ceiling fractionalDollar
        let remainder = fractionalDollar - dollar
        let digits = pow 10. (ceiling (log10  aBase))
        f  aBase dollar remainder digits        
    let dollarDe  aBase dollar remainder digits =
        remainder * digits /  aBase + dollar
    let dollarFr  aBase dollar remainder digits =
        let  absDigits = abs  digits
        remainder *  aBase / absDigits + dollar        
    let effect nominalRate npery =
        let periods = floor npery
        pow (nominalRate / periods + 1.) periods - 1.
    let nominal effectRate npery =
        let periods = floor npery
        (pow (effectRate + 1.) (1. / periods) - 1.) * periods
        
    
    // Preconditions and special cases
    let calcDollarDe fractionalDollar fraction =
        (fraction > 0.) |> elseThrow "fraction must be more than 0"
        dollar fractionalDollar fraction dollarDe
    let calcDollarFr fractionalDollar fraction =
        (fraction > 0.) |> elseThrow "fraction must be more than 0"
        (pow 10. (ceiling (log10 (floor fraction))) <> 0.) |> elseThrow "10^(ceiling (log10 (floor fraction))) must be different from 0"
        dollar fractionalDollar fraction dollarFr
    let calcEffect nominalRate npery =
        (nominalRate > 0.)  |> elseThrow "nominal rate must be more than zero"
        (npery >= 1.)       |> elseThrow "npery must be more or equal to one"
        effect nominalRate npery
    let calcNominal effectRate npery =
        (effectRate > 0.)   |> elseThrow "effective rate must be more than zero"
        (npery >= 1.)       |> elseThrow "npery must be more or equal to one"
        nominal effectRate npery
