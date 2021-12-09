#nowarn "25"

namespace Excel.FinancialFunctions.Tests

open NUnit.Framework
open System.IO
open System

[<SetCulture("en-US")>]
module CrossTests = 
    open Excel.FinancialFunctions

    // I am prototyping a new way of doing cross tests. The idea is to let the NUnit framework
    // do more of the work. So we read in the test data into a TestCaseData object, and then 
    // let NUnit work with that as normal.
    //
    // Also these files will be slightly different. They'll have a header, because I exported
    // them from Excel, and they'll have .csv file extension.

    let readTestCaseData fname goodcases =
        Path.Combine(__SOURCE_DIRECTORY__, "testdata", fname + ".csv")
        |> File.ReadLines
        |> Seq.tail // Skip header
        |> Seq.filter (fun line -> not (String.IsNullOrEmpty line))
        |> Seq.filter (fun line -> line.Contains("#NUM!") <> goodcases )
        |> Seq.map (fun line -> line.Split [| ',' |] )
        |> Seq.map TestCaseData

    let inline shouldEqual exp act =
        Assert.AreEqual(exp, float act, PRECISION)
      
    let inline elseThrow s c = if not(c) then failwith s

    let pduration_testdata_fromfile =
        readTestCaseData "pduration" true

    [<TestCaseSource( nameof pduration_testdata_fromfile)>]
    let pduration inputs =
        let (param,expected) = parse4 inputs
        Financial.Pduration param
        |> shouldEqual expected

    let pduration_failures_fromfile =
        readTestCaseData "pduration" false

    [<TestCaseSource( nameof pduration_failures_fromfile)>]
    let pduration_fail inputs =
        let (param,expected) = parse4 inputs
        ( expected = 0.0 ) |> elseThrow "Failure test must not have an expected value"
        Assert.Throws(fun () -> Financial.Pduration param |> ignore) |> ignore

    let rri_testdata_fromfile =
        readTestCaseData "rri" true

    [<TestCaseSource( nameof rri_testdata_fromfile)>]
    let rri inputs =
        let (param,expected) = parse4 inputs
        Financial.Rri param
        |> shouldEqual expected

    let rri_failures_fromfile =
        readTestCaseData "rri" false

    [<TestCaseSource( nameof rri_failures_fromfile)>]
    let rri_fail inputs =
        let (param,expected) = parse4 inputs
        ( expected = 0.0 ) |> elseThrow "Failure test must not have an expected value"
        Assert.Throws(fun () -> Financial.Rri param |> ignore) |> ignore

    // Prices for negative yielding bonds match reverse-figured Excel results to 1e-6 until yields get lower than -0.05%
    // Down to -1.0%, the difference is usually pennies. Once the yield reaches -10%, the difference can be up to $150, but let's hope
    // we never get there!
    //
    // See the file yieldnegativefails.csv for the cases which fail

    let yieldnegative_testdata_fromfile =
        readTestCaseData "yieldnegative" true

    [<TestCaseSource( nameof yieldnegative_testdata_fromfile)>]
    let yieldnegative inputs =
        let (param,expected) = parse8 inputs
        Financial.PriceAllowNegativeYield param
        |> shouldEqual expected

    [<Test>]
    let accrint() = runTests "accrint" parse8 Financial.AccrInt
        
    [<Test>]
    let accrintm() = runTests "accrintm" parse6 Financial.AccrIntM

    [<Test>]
    let amordegrc() = runTests "amordegrc" parse9 Financial.AmorDegrc

    [<Test>]
    let amorlinc() = runTests "amorlinc" parse8 Financial.AmorLinc
    
    [<Test>]
    let coupdays() = runTests "coupdays" parse5 Financial.CoupDays

    [<Test>]
    let coupdaysbs() = runTests "coupdaysbs" parse5 Financial.CoupDaysBS

    [<Test>]
    let coupdaysnc() = runTests "coupdaysnc" parse5 Financial.CoupDaysNC

    [<Test>]
    let coupncd() = 
        // compare DateTime result as ticks to use universal runTests method
        runTests "coupncd" parse5 (fun args -> (Financial.CoupNCD args).Ticks)

    [<Test>]
    let coupnum() = runTests "coupnum" parse5 Financial.CoupNum

    [<Test>]
    let couppcd() = runTests "couppcd" parse5 (fun args -> (Financial.CoupPCD args).Ticks)

    [<Test>]
    let cumipm() = runTests "cumipmt" parse7 Financial.CumIPmt

    [<Test>]
    let cumprinc() = runTests "cumprinc" parse7 Financial.CumPrinc

    [<Test>]
    let db() = runTests "db" parse6 Financial.Db

    [<Test>]
    let ddb() = runTests "ddb" parse6 Financial.Ddb

    [<Test>]
    let disc() = runTests "disc" parse6 Financial.Disc

    [<Test>]
    let dollarde() = runTests "dollarde" parse3 Financial.DollarDe

    [<Test>]
    let dollarfr() = runTests "dollarfr" parse3 Financial.DollarFr

    [<Test>]
    let duration() = runTests "duration" parse7 Financial.Duration

    [<Test>]
    let effect() = runTests "effect" parse3 Financial.Effect

    [<Test>]
    let fv() = runTests "fv" parse6 Financial.Fv

    [<Test>]
    let ipmt() = runTests "ipmt" parse7 Financial.IPmt

    [<Test>]
    let ispmt() = runTests "ispmt" parse5 Financial.ISPmt

    [<Test>]
    let intrate() = runTests "intrate" parse6 Financial.IntRate

    [<Test>]
    let mduration() = runTests "mduration" parse7 Financial.MDuration

    [<Test>]
    let nper() = runTests "nper" parse6 Financial.NPer

    [<Test>]
    let nominal() = runTests "nominal" parse3 Financial.Nominal

    [<Test>]
    let oddfprice() = runTests "oddfprice" parse10 Financial.OddFPrice

    [<Test>]
    let oddfyield() = runTests "oddfyield" parse10 Financial.OddFYield

    [<Test>]
    let ppmt() = runTests "ppmt" parse7 Financial.PPmt

    [<Test>]
    let pmt() = runTests "pmt" parse6 Financial.Pmt

    [<Test>]
    let price() = runTests "price" parse8 Financial.Price

    [<Test>]
    let pricedisc() = runTests "pricedisc" parse6 Financial.PriceDisc

    [<Test>]
    let pricemat() = runTests "pricemat" parse7 Financial.PriceMat

    [<Test>]
    let pv() = runTests "pv" parse6 Financial.Pv

    [<Test>]
    let rate() = runTests "rate" parse7 Financial.Rate

    [<Test>]
    let received() = runTests "received" parse6 Financial.Received

    [<Test>]
    let sln() = runTests "sln" parse4 Financial.Sln

    [<Test>]
    let syd() = runTests "syd" parse5 Financial.Syd

    [<Test>]
    let tbilleq() = runTests "tbilleq" parse4 Financial.TBillEq

    [<Test>]
    let tbillprice() = runTests "tbillprice" parse4 Financial.TBillPrice

    [<Test>]
    let tbillyield() = runTests "tbillyield" parse4 Financial.TBillYield

    [<Test>]
    let vdb() = runTests "vdb" parse8 Financial.Vdb

    [<Test>]
    let yearfrac() = runTests "yearfrac" parse4 Financial.YearFrac

    [<Test>]
    let yielddisc() = runTests "yielddisc" parse6 Financial.YieldDisc

    [<Test>]
    let yieldmat() = runTests "yieldmat" parse7 Financial.YieldMat

    [<Test>]
    let fvschedule() = runTests "fvschedule" (fun [| pv; interests; res |] ->
        (parse pv, parseArray interests), parse res) Financial.FvSchedule

    [<Test>]
    let irr() = runTests "irr" (fun [| cfs; guess; res |] ->
        (parseArray cfs, parse guess), parse res) Financial.Irr

    [<Test>]
    let npv() = runTests "npv" (fun [| r; cfs; res |] ->
        (parse r, parseArray cfs), parse res) Financial.Npv

    [<Test>]
    let ``npv given irr should be zero``() = 
        runTests "irr" (fun [| cfs; _; res |] ->
            (parse res, parseArray cfs), 0.) Financial.Npv
        
    [<Test>]
    let mirr() = runTests "mirr" (fun [| cfs; fr; rr; res |] ->
        (parseArray cfs, parse fr, parse rr), parse res) Financial.Mirr
        
    [<Test>]
    let xirr() = runTests "xirr" (fun [| cfs; dates; guess; res |] ->
        (parseArray cfs, parseDateArray dates, parse guess), parse res) Financial.XIrr
  