namespace Excel.FinancialFunctions.Tests

open NUnit.Framework

// These tests check that nothing goes wrong with the memoization which is used in the internal dayCount function
// that is called by the functions tested below - when used concurrently.

[<SetCulture("en-US")>]
module ConcurrencyTests =
    open System 
    open Excel.FinancialFunctions
    open TestPreconditions

    let parallelCount = 8

    let TestParallel (f : unit -> unit) =
        let expected = true
        let actual =
            try
                [|1..parallelCount|]
                |> Array.Parallel.iter (fun _ -> f())
                true
            with    
            | _ -> false

        Assert.AreEqual(expected, actual)

    let startDate = DateTime(2000, 1, 1)
    let endDate = DateTime(2010, 1, 1)

    [<Test>]
    let YearFracWorksConcurrently() =
        let f = fun () -> (Financial.YearFrac(startDate, endDate, DayCountBasis.Actual365) |> ignore)
        TestParallel f

    [<Test>]
    let CoupDaysWorksConcurrently() =
        let f = fun () -> (Financial.CoupDays(startDate, endDate, Frequency.Quarterly, DayCountBasis.Actual365) |> ignore)
        TestParallel f

    [<Test>]
    let CoupPCDWorksConcurrently() =
        let f = fun () -> (Financial.CoupPCD(startDate, endDate, Frequency.Quarterly, DayCountBasis.Actual365) |> ignore)
        TestParallel f

    [<Test>]
    let CoupNCDWorksConcurrently() =
        let f = fun () -> (Financial.CoupNCD(startDate, endDate, Frequency.Quarterly, DayCountBasis.Actual365) |> ignore)
        TestParallel f

    [<Test>]
    let CoupNumWorksConcurrently() =
        let f = fun () -> (Financial.CoupNum(startDate, endDate, Frequency.Quarterly, DayCountBasis.Actual365) |> ignore)
        TestParallel f

    [<Test>]
    let CoupDaysBSWorksConcurrently() =
        let f = fun () -> (Financial.CoupDaysBS(startDate, endDate, Frequency.Quarterly, DayCountBasis.Actual365) |> ignore)
        TestParallel f

    [<Test>]
    let CoupDaysNCWorksConcurrently() =
        let f = fun () -> (Financial.CoupDaysNC(startDate, endDate, Frequency.Quarterly, DayCountBasis.Actual365) |> ignore)
        TestParallel f

