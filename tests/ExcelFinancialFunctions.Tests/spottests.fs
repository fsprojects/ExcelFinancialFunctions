namespace ExcelFinancialFunctions.Tests

open FsCheck
open NUnit.Framework

[<SetCulture("en-US")>]
module SpotTests =
    open System 
    open System.Numerics
    open TestPreconditions

    [<Test>]
    let spotYield() =
        let param = 
            DateTime(2008, 2, 15), DateTime(2016, 11, 15), 0.0575, 95.04287, 100.0,
            Frequency.SemiAnnual, DayCountBasis.UsPsa30_360
        Financial.Yield param
        |> shouldEqual (sprintf "spotYield(%A)" param) 0.065
    
    [<Test>]
    let spotXnpv() =
        let param = 0.14, [1.;3.;4.], [DateTime(1970, 3, 2); DateTime(1988, 2, 3); DateTime(1999, 3, 5)]
        Financial.XNpv param
        |> shouldEqual (sprintf "xnpv(%A)" param) 1.375214

    [<Test>]
    let ``duration shouldn't be greater than maturity``() =
        fsCheck (fun (sd: DateTime) yrs cpn' yld' freq basis ->
            let md, cpn, yld = sd.AddYears yrs, toFloat cpn', toFloat yld'

            tryDuration sd md cpn yld freq basis
            ==>
            lazy (let duration = Financial.Duration(sd, md, cpn, yld, freq, basis)
                  duration - float yrs < PRECISION))

    [<Test>]
    let ``mduration shouldn't be greater than maturity``() =
        fsCheck (fun (sd: DateTime) yrs cpn' yld' freq basis ->
            let md, cpn, yld = sd.AddYears yrs, toFloat cpn', toFloat yld'
            
            tryMDuration sd md cpn yld freq basis
            ==>
            lazy (let duration = Financial.MDuration(sd, md, cpn, yld, freq, basis)
                  duration - float yrs < PRECISION))

    [<Test>]
    let ``tbill price is less than 100``() =
         fsCheck (fun (sd: DateTime) t disc' ->
            let md, disc = sd.AddDays (toFloat t * 365.), toFloat disc'

            tryTBillPrice sd md disc
            ==>
            lazy (Financial.TBillPrice(sd, md, disc) - 100. < PRECISION))
