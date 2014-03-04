// Depreciation calculations. AmorDegr and AmorLinc required a lot of work and trial and error. I wonder how many people are using them.
#light
namespace Excel.FinancialFunctions

open System
open Excel.FinancialFunctions.Common
open Excel.FinancialFunctions.DayCount // because of AmorDegr and AmorLinc

module internal Depreciation =

    // Main formulas
    let deprRate cost salvage life = Math.Round( 1. - (( salvage / cost) ** (1. / life)), 3)
    let deprForPeriod cost totDepr rate = (cost - totDepr) * rate
    let deprForFirstPeriod cost rate month = cost * rate * month / 12.
    let deprForLastPeriod cost totDepr rate month = (( cost - totDepr) * rate * ( 12. - month)) / 12.
    
    let db cost salvage life period month =
        let rate = deprRate cost salvage life
        let rec _db totDepr per =
            match int per with
            | 0                                         ->
                let depr = deprForFirstPeriod cost rate month
                if int period <= 1 then depr
                else _db depr (per + 1.)
            | x when x = int period - 1                 -> deprForPeriod cost totDepr rate
            | x when x = int life   - 1                 -> deprForLastPeriod cost totDepr rate month
            | _                                         ->
                let depr = deprForPeriod cost totDepr rate
                _db (totDepr + depr) (per + 1.)
        _db 0. 0.
    let sln cost salvage life = (cost - salvage) / life
    let syd cost salvage life per = ((cost - salvage) * (life - per + 1.) * 2.) / (life * (life + 1.))
    let totalDepr cost salvage life period factor straightLine =
        let rec _ddb totDepr per =
            let frac = rest period
            let ddbDeprFormula totDepr = min ((cost - totDepr) * (factor / life)) ((cost - salvage - totDepr))
            let slnDeprFormula totDepr aPeriod = sln (cost - totDepr) salvage (life - aPeriod)
            let ddbDepr, slnDepr = ddbDeprFormula totDepr, slnDeprFormula totDepr per
            let isSln = straightLine && ddbDepr < slnDepr
            let depr = if isSln then slnDepr else ddbDepr
            let newTotalDepr = totDepr + depr
            if int period = 0 then newTotalDepr * frac
            elif int per = int period - 1 then
                let ddbDeprNextPeriod = ddbDeprFormula newTotalDepr
                let slnDeprNextPeriod = slnDeprFormula newTotalDepr (per + 1.)
                let isSlnNextPeriod = straightLine && ddbDeprNextPeriod < slnDeprNextPeriod
                let deprNextPeriod =
                    if isSlnNextPeriod then
                        if int period = int life then 0.
                        else slnDeprNextPeriod
                    else
                        ddbDeprNextPeriod
                newTotalDepr + deprNextPeriod * frac
            else
                _ddb newTotalDepr (per + 1.)
        _ddb 0. 0.
    let deprBetweenPeriods cost salvage life startPeriod endPeriod factor straightLine =
        totalDepr cost salvage life endPeriod factor straightLine - totalDepr cost salvage life startPeriod factor straightLine        
    let ddb cost salvage life period factor = 
        if period >= 2.
        then deprBetweenPeriods cost salvage life (period - 1.) period factor false
        else totalDepr cost salvage life period factor false
                       
    let vdb cost salvage life startPeriod endPeriod factor bflag =
        if bflag = VdbSwitch.DontSwitchToStraightLine
        then deprBetweenPeriods cost salvage life startPeriod endPeriod factor false    
        else deprBetweenPeriods cost salvage life startPeriod endPeriod factor true
    
    let daysInYear date basis =
        if basis = DayCountBasis.ActualActual then
            if isLeapYear date then 366. else 365.
        else
            let dc = dayCount basis
            dc.DaysInYear date date
    let firstDeprLinc cost datePurch firstP salvage rate assLife basis =
        let fix29February (Date(y, m, d) as d1) =
            if (basis = DayCountBasis.ActualActual || basis = DayCountBasis.Actual365) && isLeapYear d1 && m = 2 && d >= 28
            then date y m 28 else d1 
        let dc = dayCount basis
        let daysInYr = daysInYear datePurch basis
        let datePurchased, firstPeriod = fix29February datePurch, fix29February firstP
        let firstLen = dc.DaysBetween datePurchased firstPeriod NumDenumPosition.Numerator
        let firstDeprTemp = firstLen / daysInYr * rate * cost
        let firstDepr = if firstDeprTemp = 0. then cost * rate else firstDeprTemp
        let assetLife = if firstDeprTemp = 0. then assLife else assLife + 1.
        let availDepr = cost - salvage
        if firstDepr > availDepr then availDepr, assetLife else firstDepr, assetLife
        
    let amorLinc cost datePurchased firstPeriod salvage period rate basis =
        let assetLifeTemp = ceiling (1. / rate)
        let rec findDepr countedPeriod depr availDepr =
            if countedPeriod > period then depr
            else
                let depr = if depr > availDepr then availDepr else depr
                let availDeprTemp = availDepr - depr
                let availDepr = if availDeprTemp < 0. then 0. else availDeprTemp
                findDepr (countedPeriod + 1.) depr availDepr
        if cost = salvage || period > assetLifeTemp then 0.
        else
            let firstDepr, _ = firstDeprLinc cost datePurchased firstPeriod salvage rate assetLifeTemp basis
            if period = 0. then firstDepr
            else findDepr 1. (rate * cost) (cost - salvage - firstDepr)
    let deprCoeff assetLife =
        let between x1 x2 = assetLife >= x1 && assetLife <= x2
        if between 3. 4. then 1.5
        elif between 5. 6. then 2.
        elif assetLife > 6. then 2.5
        else 1.
    let amorDegrc cost datePurchased firstPeriod salvage period rate basis excelComplaint =
        let assLife = ceiling (1. / rate)
        if cost = salvage || period > assLife then 0.
        else
            let deprCoeff = deprCoeff assLife
            let deprR = rate * deprCoeff
            let firstDeprLinc, assetLife = firstDeprLinc cost datePurchased firstPeriod salvage deprR assLife basis
            let firstDepr = round excelComplaint firstDeprLinc
            let rec findDepr countedPeriod depr deprRate remainCost =
                if countedPeriod > period then round excelComplaint depr
                else
                    let countedPeriod = countedPeriod + 1.
                    let calcT = assetLife - countedPeriod
                    let deprTemp = if areEqual calcT 2. then remainCost * 0.5 else deprRate * remainCost
                    let deprRate = if areEqual calcT 2. then 1. else deprRate
                    let depr =
                        if remainCost < salvage then
                            if remainCost - salvage < 0. then 0. else remainCost - salvage
                        else deprTemp
                    let remainCost = remainCost - depr
                    findDepr countedPeriod depr deprRate remainCost                    
            if period = 0. then firstDepr
            else findDepr 1. 0. deprR (cost - firstDepr)                                         
                         
    // Preconditions and special cases
    let calcDb cost salvage life period month =
        ( cost >= 0. )      |> elseThrow "Cost must be 0 or more"
        ( salvage >= 0. )   |> elseThrow "Salvage must be 0 or more"
        ( life > 0. )       |> elseThrow "Life must be 0 or more"
        ( month > 0. )      |> elseThrow "Month must be 0 or more"
        ( period <= life )  |> elseThrow "Period must be less than life"
        ( period > 0. )     |> elseThrow "Period must be more than 0"
        ( month <= 12. )    |> elseThrow "Month must be less or equal to 12"
        db cost salvage life period month
    let calcSln cost salvage life =
        ( cost >= 0. )      |> elseThrow "Cost must be 0 or more"
        ( salvage >= 0. )   |> elseThrow "Salvage must be 0 or more"
        ( life > 0. )       |> elseThrow "Life must be 0 or more"
        sln cost salvage life
    let calcSyd cost salvage life period =
        ( cost >= 0. )      |> elseThrow "Cost must be 0 or more"
        ( salvage >= 0. )   |> elseThrow "Salvage must be 0 or more"
        ( life > 0. )       |> elseThrow "Life must be 0 or more"
        ( period <= life )  |> elseThrow "Period must be less than life"
        ( period > 0. )     |> elseThrow "Period must be more than 0"
        syd cost salvage life period
    let calcDdb cost salvage life period factor =
        ( cost >= 0. )      |> elseThrow "Cost must be 0 or more"
        ( salvage >= 0. )   |> elseThrow "Salvage must be 0 or more"
        ( life > 0. )       |> elseThrow "Life must be 0 or more"
        ( factor > 0. )     |> elseThrow "Month must be 0 or more"
        ( period <= life )  |> elseThrow "Period must be less than life"
        ( period > 0. )     |> elseThrow "Period must be more than 0"
        if int period = 0 then min (cost * (factor / life)) ((cost - salvage)) 
        else ddb cost salvage life period factor
    let calcVdb cost salvage life startPeriod endPeriod factor bflag =
        ( cost >= 0. )              |> elseThrow "Cost must be 0 or more"
        ( salvage >= 0. )           |> elseThrow "Salvage must be 0 or more"
        ( life > 0. )               |> elseThrow "Life must be 0 or more"
        ( factor > 0. )             |> elseThrow "Month must be 0 or more"
        ( startPeriod <= life )     |> elseThrow "StartPeriod must be less than life"
        ( endPeriod <= life )       |> elseThrow "EndPeriod must be less than life"
        ( startPeriod <= endPeriod )|> elseThrow "StartPeriod must be less than endPeriod"
        ( endPeriod > 0. )          |> elseThrow "EndPeriod must be more than 0"
        ( bflag = VdbSwitch.DontSwitchToStraightLine || not(life = startPeriod && startPeriod = endPeriod) ) |> elseThrow "If bflag is set to SwitchToStraightLine, then life, startPeriod and endPeriod cannot all have the same value"
        vdb cost salvage life startPeriod endPeriod factor bflag 
    let calcAmorLinc cost datePurchased firstPeriod salvage period rate basis =
        ( cost >= 0. )                      |> elseThrow "Cost must be 0 or more"
        ( salvage >= 0. )                   |> elseThrow "Salvage must be 0 or more"
        ( salvage < cost )                  |> elseThrow "Salvage must be less than cost"
        ( period >= 0. )                    |> elseThrow "Period must be 0 or more"
        ( rate >= 0. )                      |> elseThrow "Rate must be 0 or more"
        (datePurchased < firstPeriod)       |> elseThrow "DatePurchased must be less than FirstPeriod"
        (basis <> DayCountBasis.Actual360 ) |> elseThrow "basis cannot be Actual360" 
        amorLinc cost datePurchased firstPeriod salvage period rate basis
    let calcAmorDegrc cost datePurchased firstPeriod salvage period rate basis excelComplaint =
        let assetLife = 1. / rate
        let between x1 x2 = assetLife >= x1 && assetLife <= x2
        ( not(between 0. 3.) )              |> elseThrow "Assset life cannot be between 0 and 3"
        ( not(between 4. 5.) )              |> elseThrow "Assset life cannot be between 4. and 5."
        ( cost >= 0. )                      |> elseThrow "Cost must be 0 or more"
        ( salvage >= 0. )                   |> elseThrow "Salvage must be 0 or more"
        ( salvage < cost )                  |> elseThrow "Salvage must be less than cost"
        ( period >= 0. )                    |> elseThrow "Period must be 0 or more"
        ( rate >= 0. )                      |> elseThrow "Rate must be 0 or more"
        (datePurchased < firstPeriod)       |> elseThrow "DatePurchased must be less than FirstPeriod"
        (basis <> DayCountBasis.Actual360 ) |> elseThrow "basis cannot be Actual360" 
        amorDegrc cost datePurchased firstPeriod salvage period rate basis excelComplaint
