// Test infrastructure. The idea is that Excel is the oracle.
// I test a whole bunch of different values for parameters against Excel and check that the result is the same.
#light
namespace Excel.FinancialFunctions

open System
open System.Collections
open Excel.FinancialFunctions.ExcelTesting
open Excel.FinancialFunctions.Common
open Excel.FinancialFunctions.Tvm

module internal TestInfrastructure =
    
    let totTries = ref 0
    let totSuccess = ref 0
    // Closing down Excel singleton
    let endTests () =
        printfn "###### Total Test Cases Succeeded %i/%i" !totSuccess !totTries |> ignore
        app.Quit() 
        Console.ReadKey ()
                   
    let tuplize f = fun (x, y) -> f x y

    type Result = 
        | Ok 
        | Fail of (float * float * obj) 
        | Exn of (exn * obj)

    let check1 f1 f2 inp = 
        try 
            let r1,r2 = f1 inp, f2 inp 
            if areEqual r1 r2 then
               Ok
            else             
               Fail(r1,r2,box(inp))
        with exn -> 
            Exn (exn,box(inp))
           
    let test1 f1 f2 inputs precond =
        let tries = ref 0
        let successes = ref 0
        let failures = new ResizeArray<_>()
        let exceptions = new ResizeArray<_>()
        for inp in inputs  do
            if precond inp then
                incr tries;
                incr totTries
                let res = check1 f1 f2 inp
                match res with 
                | Ok ->
                    incr successes
                    incr totSuccess
                | Fail(data) -> failures.Add(data)
                | Exn(data) -> exceptions.Add(data)
        !successes, !tries, failures.ToArray(), exceptions.ToArray()

    let square s1 s2 = 
            seq { for x in s1 do for y in s2 do yield (x,y) }

    /// Pack an pair of sequences together
    let pack checker f1 f2 s1 s2    = 
        checker (fun (x,y) -> f1 x y) (fun (x,y) -> f2 x y) (square s1 s2)
    let test2 f1 f2 s1 s2 precond                   = pack test1 f1 f2 s1 s2 (tuplize precond)
    let test3 f1 f2 s1 s2 s3 precond                = pack test2 f1 f2 s1 s2 s3 (tuplize precond)
    let test4 f1 f2 s1 s2 s3 s4 precond             = pack test3 f1 f2 s1 s2 s3 s4 (tuplize precond)
    let test5 f1 f2 s1 s2 s3 s4 s5 precond          = pack test4 f1 f2 s1 s2 s3 s4 s5 (tuplize precond)
    let test6 f1 f2 s1 s2 s3 s4 s5 s6 precond       = pack test5 f1 f2 s1 s2 s3 s4 s5 s6 (tuplize precond) 
    let test7 f1 f2 s1 s2 s3 s4 s5 s6 s7 precond    = pack test6 f1 f2 s1 s2 s3 s4 s5 s6 s7 (tuplize precond)
    let test8 f1 f2 s1 s2 s3 s4 s5 s6 s7 s8 precond = pack test7 f1 f2 s1 s2 s3 s4 s5 s6 s7 s8 (tuplize precond)
    let test9 f1 f2 s1 s2 s3 s4 s5 s6 s7 s8 s9 precond = pack test8 f1 f2 s1 s2 s3 s4 s5 s6 s7 s8 s9 (tuplize precond)

    let precondOk1 _ = true
    let precondOk2 _ _ = true
    
    let spotTest1 f1 f2 p1 =
        (areEqual (f1 p1) (f2 p1)) |> elseThrow (sprintf "%f <> %f in a spot test" (f1 p1) (f2 p1))
    let spotTest2 f1 f2 p1 p2 =
        (areEqual (f1 p1 p2) (f2 p1 p2)) |> elseThrow (sprintf "%f <> %f in a spot test" (f1 p1 p2) (f2 p1 p2))
    let spotTest3 f1 f2 p1 p2 p3 =
        (areEqual (f1 p1 p2 p3) (f2 p1 p2 p3)) |> elseThrow (sprintf "%f <> %f in a spot test" (f1 p1 p2 p3) (f2 p1 p2 p3))
    let spotTest4 f1 f2 p1 p2 p3 p4 =
        (areEqual (f1 p1 p2 p3 p4) (f2 p1 p2 p3 p4)) |> elseThrow (sprintf "%f <> %f in a spot test" (f1 p1 p2 p3 p4) (f2 p1 p2 p3 p4))
    let spotTest5 f1 f2 p1 p2 p3 p4 p5 =
        (areEqual (f1 p1 p2 p3 p4 p5) (f2 p1 p2 p3 p4 p5)) |> elseThrow (sprintf "%f <> %f in a spot test" (f1 p1 p2 p3 p4 p5) (f2 p1 p2 p3 p4 p5))
    let spotTest6 f1 f2 p1 p2 p3 p4 p5 p6 =
        (areEqual (f1 p1 p2 p3 p4 p5 p6) (f2 p1 p2 p3 p4 p5 p6)) |> elseThrow (sprintf "%f <> %f in a spot test" (f1 p1 p2 p3 p4 p5 p6) (f2 p1 p2 p3 p4 p5 p6))
    let spotTest7 f1 f2 p1 p2 p3 p4 p5 p6 p7 =
        (areEqual (f1 p1 p2 p3 p4 p5 p6 p7) (f2 p1 p2 p3 p4 p5 p6 p7)) |> elseThrow (sprintf "%f <> %f in a spot test" (f1 p1 p2 p3 p4 p5 p6 p7) (f2 p1 p2 p3 p4 p5 p6 p7))
    let spotTest8 f1 f2 p1 p2 p3 p4 p5 p6 p7 p8 =
        (areEqual (f1 p1 p2 p3 p4 p5 p6 p7 p8) (f2 p1 p2 p3 p4 p5 p6 p7 p8)) |> elseThrow (sprintf "%f <> %f in a spot test" (f1 p1 p2 p3 p4 p5 p6 p7 p8) (f2 p1 p2 p3 p4 p5 p6 p7 p8))
    let spotTest9 f1 f2 p1 p2 p3 p4 p5 p6 p7 p8 p9 =
        (areEqual (f1 p1 p2 p3 p4 p5 p6 p7 p8 p9) (f2 p1 p2 p3 p4 p5 p6 p7 p8 p9)) |> elseThrow (sprintf "%f <> %f in a spot test" (f1 p1 p2 p3 p4 p5 p6 p7 p8 p9) (f2 p1 p2 p3 p4 p5 p6 p7 p8 p9))
           
    // Pretty (?) print functions
    let banner name s t =
        printfn "## Succeeded %A/%A for %s" s t name
        
    let printSummary suc tr = printfn "### Successes/tries %i/%i" suc tr
    
    let printEndSpotTests () = printfn "#### Spot tests completed succesfully as well!!"
        
    let printExceptions exns =
        for e in exns do
            let (ex:Exception), parms = e
            printfn "\tException < %s > thrown when passing %A" (ex.Message) parms
             
    let printErrors errors =
        for r in errors do
            let (r1:float), (r2:float), parms = r
            if r1 > 10000000. then
                printfn "\t%A <> %A | when passing %A" ((new DateTime(int64 r1)).ToShortDateString()) ((new DateTime(int64 r2)).ToShortDateString()) parms
            else            
                printfn "\t%A <> %A | when passing %A" r1 r2 parms
                        
    let printResults results =
        let mutable succ = 0
        let mutable tr = 0
        for r in results do
            let name, (s, t, errors, exns) = r
            banner name s t
            printErrors errors
            printExceptions exns
            succ <- succ + s
            tr <- tr + t


// Multithreading away, adapted from "F# for scientists" ...              
    open System.Threading
    let spawn (f: unit -> unit) =
        let thread = new Thread(f)
        thread.Start ()
        thread
                               
    let execute n f =
        [|for i in 1 .. n ->
            spawn f|]
        |> Array.iter (fun t -> t.Join())
        
    let next i n () =
        if !i = n then None else
            incr i
            Some(!i - 1)
            
    let map max_threads a =
        let n = Array.length a
        let b = Array.create n None
        let i = ref 0
        let rec apply i =
            let (name, func) = a.[i]
            b.[i] <- Some(name, func ())
            loop ()
        and loop () =
            Option.iter apply (lock i (next i n))
        execute max_threads loop
        Array.map Option.get b
     
    let cpu_map a = map System.Environment.ProcessorCount a

