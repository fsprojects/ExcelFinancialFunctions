// Finding internal rate of return routines. I use a different algo then excel. The results might be different.
#light
namespace Excel.FinancialFunctions

open System
open Excel.FinancialFunctions.Common
open Excel.FinancialFunctions.Tvm

module internal Irr =

    // Main formulas
    let npv r cfs = cfs |> Seq.mapi (fun i cf -> cf * pvFactor r (float (i+1))) |> Seq.sumBy idem        
    let irr cfs guess = findRoot (fun r -> npv r cfs) guess
    let mirr cfs financeRate reinvestRate =
        let n = float (Seq.length cfs)
        let positives = cfs |> Seq.map (fun cf -> if cf > 0. then cf else 0.)
        let negatives = cfs |> Seq.map (fun cf -> if cf < 0. then cf else 0.)
        (((- npv reinvestRate positives) * ((1. + reinvestRate) ** n))/
         ((  npv financeRate negatives)  * ( 1. + financeRate ))) ** (1./(n - 1.)) - 1.
    let xnpv r cfs dates =
        let d0 = Seq.head dates
        cfs |> Seq.map2 (fun d cf -> cf / ((1. + r) ** (float (days d d0) / 365.))) dates |> Seq.sumBy idem
    let xirr cfs dates guess = findRoot (fun r -> xnpv r cfs dates) guess
       
    // Preconditions and special cases
    let validCfs cfs =
        let rec _validCfs cfs pos neg =
            if pos && neg then true
            else match cfs with
                    | h::t when h > 0.  -> _validCfs t true neg
                    | h::t when h <= 0. -> _validCfs t pos true
                    | []                -> false
                    | _                 -> failwith "Should never get here"
        _validCfs (Seq.toList cfs) false false  
    let calcIrr cfs guess =
        validCfs cfs                |> elseThrow "There must be one positive and one negative cash flow"
        irr cfs guess 
    let calcNpv r cfs =
        ( r <> -1.)                 |> elseThrow "r cannot be -100%"
        npv r cfs
    let calcMirr cfs financeRate reinvestRate =
        ( financeRate  <> -1.)      |> elseThrow "financeRate cannot be -100%"
        ( reinvestRate <> -1.)      |> elseThrow "reinvestRate cannot be -100%"
        ( Seq.length cfs <> 1)      |> elseThrow "cfs must contain more than one cashflow"
        ( (npv financeRate (cfs |> Seq.map (fun cf -> if cf < 0. then cf else 0.)))  <> 0. ) |> elseThrow "The NPV calculated using financeRate and the negative cashflows in cfs must be different from zero"
        mirr cfs financeRate reinvestRate        
    let calcXnpv r cfs dates =
        ( r <> -1.)                                         |> elseThrow "r cannot be -100%"
        not(Seq.exists (fun x -> x < Seq.head dates) dates)   |> elseThrow "In dates, one date is less than the first date"
        (Seq.length cfs = Seq.length dates)                 |> elseThrow "cfs and dates must have the same length"
        xnpv r cfs dates    
    let calcXirr cfs dates guess =
        validCfs cfs                                        |> elseThrow "There must be one positive and one negative cash flow"
        not(Seq.exists (fun x -> x < Seq.head dates) dates)   |> elseThrow "In dates, one date is less than the first date"
        (Seq.length cfs = Seq.length dates)                 |> elseThrow "cfs and dates must have the same length"
        xirr cfs dates guess
        
                          
