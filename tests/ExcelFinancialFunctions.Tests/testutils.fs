#nowarn "25"

namespace Excel.FinancialFunctions.Tests

open FsCheck
open NUnit.Framework
open System
open System.IO

[<AutoOpen>]
module internal TestUtils =
    [<Literal>]
    let PRECISION = 1e-6

    let readTestData fname =
        Path.Combine(__SOURCE_DIRECTORY__, "testdata", fname + ".test")
        |> File.ReadAllLines
        |> Seq.filter (fun line -> not (String.IsNullOrEmpty line))
        |> Seq.map (fun line -> line.Split [| ',' |])
    
    let inline shouldEqual msg exp act =
        Assert.AreEqual(exp, float act, PRECISION, msg)

    let inline runTests fname parsef f =
        readTestData fname    
        |> Seq.iteri (fun i data ->
            let param, expected = parsef data
            let actual = f param
            shouldEqual (sprintf "%d - %s(%A)" i fname param) expected actual)

    let inline parse str =
        let mutable res = Unchecked.defaultof<_>
        let _ = (^a: (static member TryParse: string * byref< ^a > -> bool) (str, &res))
        res

    let inline parseArray (str: string) = str.Split [| ';' |] |> Array.map parse

    // universal parse methods for function arguments and result
    let inline parse3 [| a; b; c |] =
        (parse a, parse b), parse c
    let inline parse4 [| a; b; c; d |] =
        (parse a, parse b, parse c), parse d
    let inline parse5 [| a; b; c; d; e |] =
        (parse a, parse b, parse c, parse d), parse e
    let inline parse6 [| a; b; c; d; e; f |] =
        (parse a, parse b, parse c, parse d, parse e), parse f
    let inline parse7 [| a; b; c; d; e; f; g |] =
        (parse a, parse b, parse c, parse d, parse e, parse f), parse g
    let inline parse8 [| a; b; c; d; e; f; g; h |] =
        (parse a, parse b, parse c, parse d, parse e, parse f, parse g), parse h
    let inline parse9 [| a; b; c; d; e; f; g; h; i |] =
        (parse a, parse b, parse c, parse d, parse e, parse f, parse g, parse h), parse i
    let inline parse10 [| a; b; c; d; e; f; g; h; i; j |] =
        (parse a, parse b, parse c, parse d, parse e, parse f, parse g, parse h, parse i), parse j


    let private nUnitRunner =
        { new IRunner with
            member x.OnStartFixture t = ()
            member x.OnArguments(ntest, args, every) = ()
            member x.OnShrink(args, everyShrink) = ()
            member x.OnFinished(name, result) =
                match result with
                | TestResult.True data ->
                    printfn "%s" (Runner.onFinishedToString name result)
                | _ -> Assert.Fail(Runner.onFinishedToString name result) }
   
    let private nUnitConfig = { Config.Default with Runner = nUnitRunner }

    let fsCheck testable =
        FsCheck.Check.One (nUnitConfig, testable)

    let inline toFloat x = float x / float Int32.MaxValue