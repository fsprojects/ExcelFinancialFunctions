// Common math, error management, zero finding, etc... routines used in all the rest of the library
#light
namespace Excel.FinancialFunctions

open System
open System.Collections.Generic

module internal Common =
    
    // Error management functions
    let mutable precision = 0.0001
    let areEqual x y = if abs(x - y) < precision then true else false
    let throw s = failwith s
    let elseThrow s c = if not(c) then throw s   
    let raisable b p = not( (1. + b) < 0. && (p - float (int p)) <> 0. )   
    // Mathematical functions
    let ln x = Math.Log(x)
    let sign (x:float) = Math.Sign(x)
    let idem x = x
    let min (x:float) y = Math.Min(x, y)
    let max (x:float) y = Math.Max(x, y)
    let rest x = x - float (int x)
    let ceiling (x:float) = Math.Ceiling(x)
    let floor (x:float) = Math.Floor(x)
    let pow x y = Math.Pow(x, y)
    let sqr x = Math.Sqrt(x)
    let log10 x = Math.Log(x, 10.)
    let round excelComplaint (x:float) =
        // Excel precision is of 13 digits so to be Excel compatible you need to preround to 13 digits ...
        if excelComplaint then
            let k = Math.Round(x, 13, MidpointRounding.AwayFromZero)
            Math.Round(k, MidpointRounding.AwayFromZero)
        else Math.Round(x, MidpointRounding.AwayFromZero)

    // Don't want to use fold directly as it is hard to read. Building simpler utility func instead ...
    let aggrBetween startPeriod endPeriod f initialValue=
        let s = if startPeriod <= endPeriod then {startPeriod .. 1 .. endPeriod} else {startPeriod .. -1 .. endPeriod}
        s |> Seq.fold f initialValue
           
    // Date functions
    let days (after:DateTime) (before:DateTime) = (after - before).Days
    let date y m d = new DateTime(y, m, d)
    let (|Date|) (d1:DateTime) = (d1.Year,d1.Month,d1.Day)
    let isLeapYear (Date(y,_,_) as d) = DateTime.IsLeapYear(y)
    let leapYear y = DateTime.IsLeapYear(y)
    let lastDayOfMonth y m d = DateTime.DaysInMonth(y, m) = d
    let lastDayOfFebruary (Date(y, m, d) as dt) = m = 2 && lastDayOfMonth y m d
    let daysOfMonth y m = DateTime.DaysInMonth(y, m)
    let addYears (d: DateTime) n = d.AddYears(n)
       
    // Find an interval that bounds the root, (shift, factor, maxtTries) are guesses
    let findBounds f guess minBound maxBound precision =
        if guess <= minBound || guess >= maxBound then throw (sprintf "guess needs to be between %f and %f" minBound maxBound) 
        let shift = 0.01
        let factor = 1.6
        let maxTries = 60
        let adjValueToMin value = if value <= minBound then minBound + precision else value
        let adjValueToMax value = if value >= maxBound then maxBound - precision else value
        let rec rfindBounds low up tries =
            let tries = tries - 1
            if tries = 0 then throw (sprintf "Not found an interval comprising the root after %i tries, last tried was (%f, %f)" maxTries low up) 
            let lower = adjValueToMin low
            let upper = adjValueToMax up 
            match f lower, f upper with
            | x, y when (x*y = 0.)          -> lower, upper
            | x, y when (x*y < 0.)          -> lower, upper
            | x, y when (x*y > 0.)          -> rfindBounds (lower + factor * (lower - upper)) (upper + factor * (upper - lower)) tries 
            | x, y                          -> throw (sprintf "FindBounds: one of the values (%f, %f) cannot be used to evaluate the objective function" lower upper)        
        let low = adjValueToMin (guess - shift) 
        let high = adjValueToMax (guess + shift)        
        rfindBounds low high maxTries

    // Very simple bisection algo. (200) is a guess. It is high. The reason is that if a root doesn't exist, I don't mind the slight perf degradation of 200 iters. But I want to catch it if it exists.
    let bisection =
        let maxCount = 200
        let rec helper f a b count precision =
            if a = b then throw (sprintf "(a=b=%f) impossible to start bisection" a) 
            
            let fa = f a
            if abs fa < precision then a // a is the root
            else
                let fb = f b
                if abs fb < precision then b // b is the root
                else
                    let newCount = count + 1
                
                    if newCount > maxCount then throw (sprintf "No root found in %i iterations" maxCount)
                    if fa * fb > 0. then throw (sprintf "(%f,%f) don't bracket the root" a b)
                    
                    let midvalue = a + 0.5 * (b - a)
                    let fmid = f midvalue
                    
                    if abs fmid < precision then midvalue // the midvalue is the root
                    elif fa * fmid < 0. then helper f a midvalue newCount precision
                    elif fa * fmid > 0. then helper f midvalue b newCount precision
                    else throw "Bisection: It should never get here" 
        helper
        
    let newton =
        let maxCount = 20
        let rec helper f x count precision =
            let d f x = (f (x + precision) - f (x - precision))/(2. * precision)
            let fx = f x
            let Fx = d f x
            let newX = x - (fx / Fx)
            if abs (newX - x) < precision then Some( newX )
            elif count > maxCount then None
            else helper f newX (count + 1) precision
        helper

    // This is my main root finding algo. My strategy is to try a fast but not precise one (newton) first.
    // If the result is sensible (it exist and has the same sign as guess), then return it, else try bisection.
    // I'm sure more complex way to pick algos exist (i.e. Brent's method). But I favor simplicity here ...         
    let findRoot f guess =
        let precision = 0.0000001 // Excel precision on this, from docs
        let newtValue = newton f guess 0 precision
        if newtValue.IsSome && sign guess = sign newtValue.Value
        then newtValue.Value
        else
            let lower, upper = findBounds f guess -1.0 Double.MaxValue precision
            bisection f lower upper 0 precision
     
    let memoize f =
        let m = new Dictionary<_,_> ()
        fun x ->
                lock m (fun () ->
                    let foundIt, res = m.TryGetValue(x)
                    if foundIt then res
                    else
                        let r = f x
                        m.Add(x, r)
                        r
                )