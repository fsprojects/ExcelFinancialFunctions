namespace Excel.FinancialFunctions.Tests

open FsCheck
open NUnit.Framework

[<SetCulture("en-US")>]
module SpotTests =
    open System 
    open Excel.FinancialFunctions
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

    [<Test>]
    let ``vdb in the period should be the same as the sum by subperiods``() =
        fsCheck (fun c s len switchToStraightLine ->
            let start, life = 0., 10.
            let cost, salvage, period = toFloat c * 1000., toFloat s * 100., toFloat len
            let p1, p2 = period, period * 2.
            
            tryVdb cost salvage life start p2 1. switchToStraightLine
            ==>
            lazy (let inline vdb sd ed = Financial.Vdb(cost, salvage, life, sd, ed, 1., switchToStraightLine)
                  abs (vdb start p1 + vdb p1 p2 - vdb start p2) < PRECISION))

    [<Test>]
    let ``cumulative accrint should be the same as the sum of payments``() =
        fsCheck (fun (issue: DateTime) p r freq basis frac ->
            let par, rate = toFloat p * 1000000., toFloat r
            let fd, md = issue.AddDays 30., issue.AddYears 10            
            let ncd = Financial.CoupNCD(fd, md, freq, basis)

            tryAccrInt issue ncd fd rate par freq basis
            ==>
            lazy (let inline accrint interest settlement =
                      Financial.AccrInt(issue, interest, settlement, rate, par, freq, basis)

                  let daysTillInterest = max (360. / float freq * toFloat frac) 1.
                  let interest = ncd.AddDays daysTillInterest
                  abs (accrint ncd fd + accrint interest ncd - accrint interest fd) < PRECISION))
