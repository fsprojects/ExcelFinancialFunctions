// Messy, but excel compatible, treatment of day count conventions for bond mathematics.
// I tried to abstract out the commonality in one interface, but I have some special cases in the rest of the code
// If you want to support your own daycount, you should be ok just implementing the IDayCount interface 
#light
namespace Excel.FinancialFunctions

open System
open Excel.FinancialFunctions.Common

module internal DayCount =
    // Some of the Excel day count info comes from http://www.dwheeler.com/yearfrac/excel-ooxml-yearfrac.pdf
    
    type Method360Us =
    | ModifyStartDate
    | ModifyBothDates

    type NumDenumPosition =
    | Denumerator
    | Numerator

    // Date utility functions, not needed elsewhere, hence encapsulated in this module
    let freq2months freq = 12 / int freq
    let lastDayOfMonthBasis y m d basis = lastDayOfMonth y m d || (d = 30 && basis = DayCountBasis.UsPsa30_360)
    let changeMonth (Date(y, m, d) as orgDate) numMonths basis returnLastDay =
        let isLastDay = lastDayOfMonthBasis y m d basis
        let getLastDay y m = DateTime.DaysInMonth(y, m)
        let (Date(year, month, dayTemp) as newDate) = orgDate.AddMonths(numMonths)
        if returnLastDay then date year month (getLastDay year month) else newDate
    let noActionDates (d1:DateTime) (d2:DateTime) = 0.
    let datesAggregate1 startDate endDate numMonths basis f acc returnLastMonth =
        let rec iter frontDate trailingDate acc =
            let stop = if numMonths > 0 then frontDate >= endDate else frontDate <= endDate
            if stop then frontDate, trailingDate, acc
            else
                let trailingDate = frontDate
                let frontDate = changeMonth frontDate numMonths basis returnLastMonth 
                let acc = acc + f frontDate trailingDate
                iter frontDate trailingDate acc
        iter startDate endDate acc
    // Wasting a bit of time on the aggregation, but more concise code ...
    let findPcdNcd startDate endDate numMonths basis returnLastMonth =
        let pcd, ncd, _ = datesAggregate1 startDate endDate numMonths basis noActionDates 0. returnLastMonth
        pcd, ncd    
    let findCouponDates settl (Date(my, mm, md) as mat) freq basis =
        let endMonth = lastDayOfMonth my mm md 
        let numMonths = - freq2months freq
        findPcdNcd mat settl numMonths basis endMonth
    let findPreviousCouponDate settl mat freq basis =
        findCouponDates settl mat freq basis |> fst
    let findNextCouponDate settl mat freq basis =
        findCouponDates settl mat freq basis |> snd
    let numberOfCoupons settl (Date(my, mm, md) as mat) freq basis =
        let (Date(pcy, pcm, pcd) as pcdate) = findPreviousCouponDate settl mat freq basis
        let months = float ((my - pcy) * 12 + (mm - pcm))
        months * freq / 12.
    let lessOrEqualToAYearApart (Date(y1, m1, d1) as date1) (Date(y2, m2, d2) as date2) =
        y1 = y2 || (y2 = y1 + 1 && (m1 > m2 || (m1 = m2 && d1 >= d2)))
    let isFeb29BetweenConsecutiveYears (Date(y1, m1, d1) as date1) (Date(y2, m2, d2) as date2) =
        if y1 = y2 && isLeapYear date1 then if m1 <= 2 && m2 > 2 then true else false
        elif y1 = y2 then false
        elif y2 = y1 + 1 then
            if isLeapYear date1 then if m1 <= 2 then true else false
            elif isLeapYear date2 then if m2 > 2 then true else false
            else false
        else throw "isFeb29BetweenConsecutiveYears: function called with non consecutive years"
    let considerAsBisestile (Date(y1, m1, d1) as date1) (Date(y2, m2, d2) as date2) =
        (y1 = y2 && isLeapYear date1) || (m2 = 2 && d2 = 29) || isFeb29BetweenConsecutiveYears date1 date2
                
    let dateDiff360 sd sm sy ed em ey  =
        (ey - sy) * 360 + (em - sm) * 30 + (ed - sd)
    let dateDiff365 (Date(sy,sm,sd) as startDate) (Date(ey,em,ed) as endDate) =
        let mutable sd1, sm1, sy1, ed1, em1, ey1, startDate1, endDate1 = sd, sm, sy, ed, em, ey, startDate, endDate
        if sd1 > 28 && sm1 = 2 then sd1 <- 28
        if ed1 > 28 && em1 = 2 then ed1 <- 28
        let startd, endd = date sy1 sm1 sd1, date ey1 em1 ed1
        (ey1 - sy1) * 365 + days endd startd        
    let dateDiff360Us (Date(sy,sm,sd) as startDate) (Date(ey,em,ed) as endDate)  method360 =
        let mutable sd1, sm1, sy1, ed1, em1, ey1, startDate1, endDate1 = sd, sm, sy, ed, em, ey, startDate, endDate
        if lastDayOfFebruary endDate1 && (lastDayOfFebruary startDate1 || method360 = Method360Us.ModifyBothDates)
            then ed1 <- 30
        if ed1 = 31 && (sd1 >= 30 || method360 = Method360Us.ModifyBothDates) then ed1 <- 30
        if sd1 = 31 then sd1 <- 30
        if lastDayOfFebruary startDate1 then sd1 <- 30
        dateDiff360 sd1 sm1 sy1 ed1 em1 ey1
    let dateDiff360Eu (Date(sy,sm,sd) as startDate) (Date(ey,em,ed) as endDate) =
        let mutable sd1, sm1, sy1, ed1, em1, ey1, startDate1, endDate1 = sd, sm, sy, ed, em, ey, startDate, endDate
        sd1 <- if sd1 = 31 then 30 else sd1
        ed1 <- if ed1 = 31 then 30 else ed1
        dateDiff360 sd1 sm1 sy1 ed1 em1 ey1

    type IDayCount =
        abstract CoupDays: DateTime -> DateTime -> float -> float
        abstract CoupPCD: DateTime -> DateTime -> float -> DateTime
        abstract CoupNCD: DateTime -> DateTime -> float -> DateTime
        abstract CoupNum: DateTime -> DateTime -> float -> float
        abstract CoupDaysBS: DateTime -> DateTime -> float -> float
        abstract CoupDaysNC: DateTime -> DateTime -> float -> float
        abstract DaysBetween: DateTime -> DateTime -> NumDenumPosition -> float
        abstract DaysInYear: DateTime -> DateTime -> float
        abstract ChangeMonth: DateTime -> int -> bool -> DateTime
                          
    let UsPsa30_360 () =
        { new IDayCount with
            member dc.CoupDays settl mat freq =
                360. / freq
            member dc.CoupPCD settl mat freq =
                findPreviousCouponDate settl mat freq DayCountBasis.UsPsa30_360    
            member dc.CoupNCD settl mat freq =
                findNextCouponDate settl mat freq DayCountBasis.UsPsa30_360
            member dc.CoupNum settl mat freq =
                numberOfCoupons settl mat freq DayCountBasis.UsPsa30_360
            member dc.CoupDaysBS settl mat freq =
                float(dateDiff360Us (dc.CoupPCD settl mat freq) settl Method360Us.ModifyStartDate)
            member dc.CoupDaysNC settl mat freq =
                let pdc = findPreviousCouponDate settl mat freq DayCountBasis.UsPsa30_360
                let ndc = findNextCouponDate settl mat freq DayCountBasis.UsPsa30_360
                let totDaysInCoup = dateDiff360Us pdc ndc Method360Us.ModifyBothDates 
                let daysToSettl =  dateDiff360Us pdc settl Method360Us.ModifyStartDate
                float (totDaysInCoup - daysToSettl)
            member dc.DaysBetween issue settl position =
                float (dateDiff360Us issue settl Method360Us.ModifyStartDate)
            member dc.DaysInYear issue settl =
                    360.
            member dc.ChangeMonth date months returnLastDay =
                changeMonth date months DayCountBasis.UsPsa30_360 returnLastDay
                }    
    let Europ30_360 () =
        { new IDayCount with
            member dc.CoupDays settl mat freq =
                360. / freq
            member dc.CoupPCD settl mat freq =
                findPreviousCouponDate settl mat freq DayCountBasis.Europ30_360
            member dc.CoupNCD settl mat freq =
                findNextCouponDate settl mat freq DayCountBasis.Europ30_360
            member dc.CoupNum settl mat freq =
                numberOfCoupons settl mat freq DayCountBasis.Europ30_360
            member dc.CoupDaysBS settl mat freq =
                float(dateDiff360Eu (dc.CoupPCD settl mat freq) settl)
            member dc.CoupDaysNC settl mat freq =
                float(dateDiff360Eu settl (dc.CoupNCD settl mat freq))
            member dc.DaysBetween issue settl position =
                float (dateDiff360Eu issue settl)
            member dc.DaysInYear issue settl =
                360.
            member dc.ChangeMonth date months returnLastDay =
                changeMonth date months DayCountBasis.Europ30_360 returnLastDay
                }    
    let Actual360 () =
        { new IDayCount with
            member dc.CoupDays settl mat freq =
                360. / freq
            member dc.CoupPCD settl mat freq =
                findPreviousCouponDate settl mat freq DayCountBasis.Actual360
            member dc.CoupNCD settl mat freq =
                findNextCouponDate settl mat freq DayCountBasis.Actual360
            member dc.CoupNum settl mat freq =
                numberOfCoupons settl mat freq DayCountBasis.Actual360
            member dc.CoupDaysBS settl mat freq =
                float(days settl (dc.CoupPCD settl mat freq))
            member dc.CoupDaysNC settl mat freq =
                float (days (dc.CoupNCD settl mat freq) settl)
            member dc.DaysBetween issue settl position =
                if position = NumDenumPosition.Numerator
                then float (days settl issue)
                else float (dateDiff360Us issue settl Method360Us.ModifyStartDate)
            member dc.DaysInYear issue settl =
                360.
            member dc.ChangeMonth date months returnLastDay =
                changeMonth date months DayCountBasis.Actual360 returnLastDay               
                }    
    let Actual365 () =
        { new IDayCount with
            member dc.CoupDays settl mat freq =
                365. / freq
            member dc.CoupPCD settl mat freq =
                findPreviousCouponDate settl mat freq DayCountBasis.Actual365
            member dc.CoupNCD settl mat freq =
                findNextCouponDate settl mat freq DayCountBasis.Actual365
            member dc.CoupNum settl mat freq =
                numberOfCoupons settl mat freq DayCountBasis.Actual365
            member dc.CoupDaysBS settl mat freq =
                float(days settl (dc.CoupPCD settl mat freq))
            member dc.CoupDaysNC settl mat freq =
                float (days (dc.CoupNCD settl mat freq) settl)
            member dc.DaysBetween issue settl position =
                if position = NumDenumPosition.Numerator
                then float (days settl issue)
                else float (dateDiff365 issue settl)
            member dc.DaysInYear issue settl =
                365.
            member dc.ChangeMonth date months returnLastDay =
                changeMonth date months DayCountBasis.Actual365 returnLastDay              
                }

    let actualCoupDays settl mat freq =
        let pcd = findPreviousCouponDate settl mat freq DayCountBasis.ActualActual
        let ncd = findNextCouponDate settl mat freq DayCountBasis.ActualActual
        float (days ncd pcd)

    let ActualActual () =
        { new IDayCount with
            member dc.CoupDays settl mat freq =
                actualCoupDays settl mat freq
            member dc.CoupPCD settl mat freq =
                findPreviousCouponDate settl mat freq DayCountBasis.ActualActual
            member dc.CoupNCD settl mat freq =
                findNextCouponDate settl mat freq DayCountBasis.ActualActual
            member dc.CoupNum settl mat freq =
                numberOfCoupons settl mat freq DayCountBasis.ActualActual
            member dc.CoupDaysBS settl mat freq =
                float(days settl (dc.CoupPCD settl mat freq))
            member dc.CoupDaysNC settl mat freq =
                float (days (dc.CoupNCD settl mat freq) settl)
            member dc.DaysBetween startDate endDate position =
                 float (days endDate startDate)
            member dc.DaysInYear issue settl =
                if not(lessOrEqualToAYearApart issue settl) then
                    let totYears = (settl.Year - issue.Year) + 1
                    let totDays = days (date (settl.Year + 1) 1 1) (date issue.Year 1 1)
                    float totDays / float totYears
                elif considerAsBisestile issue settl then 366. else 365.                    
            member dc.ChangeMonth date months returnLastDay =
                changeMonth date months DayCountBasis.ActualActual returnLastDay    
                }
                        
    let dayCount = memoize (function
        | DayCountBasis.UsPsa30_360                 -> UsPsa30_360 ()
        | DayCountBasis.ActualActual                -> ActualActual ()
        | DayCountBasis.Actual360                   -> Actual360 ()
        | DayCountBasis.Actual365                   -> Actual365 ()
        | DayCountBasis.Europ30_360                 -> Europ30_360 ()
        | _                                         -> throw "dayCount: it should never get here")
    
    let calcCoupDays settlement maturity (frequency:Frequency) basis =
        maturity > settlement       |> elseThrow "settlement must be before maturity"
        let dc = dayCount basis
        dc.CoupDays settlement maturity (float (int frequency))
    let calcCoupPCD settlement maturity (frequency:Frequency) basis =
        maturity > settlement       |> elseThrow "settlement must be before maturity"
        let dc = dayCount basis
        dc.CoupPCD settlement maturity (float (int frequency))
    let calcCoupNCD settlement maturity (frequency:Frequency) basis =
        maturity > settlement       |> elseThrow "settlement must be before maturity"
        let dc = dayCount basis
        dc.CoupNCD settlement maturity (float (int frequency))
    let calcCoupNum settlement maturity (frequency:Frequency) basis =
        maturity > settlement       |> elseThrow "settlement must be before maturity"
        let dc = dayCount basis
        dc.CoupNum settlement maturity (float (int frequency))
    let calcCoupDaysBS settlement maturity (frequency:Frequency) basis =
        maturity > settlement       |> elseThrow "settlement must be before maturity"
        let dc = dayCount basis
        dc.CoupDaysBS settlement maturity (float (int frequency))
    let calcCoupDaysNC settlement maturity (frequency:Frequency) basis =
        maturity > settlement       |> elseThrow "settlement must be before maturity"
        let dc = dayCount basis
        dc.CoupDaysNC settlement maturity (float (int frequency))
    let calcYearFrac startDate endDate basis =
        startDate < endDate         |> elseThrow "startDate must be before endDate"
        let dc = dayCount basis
        dc.DaysBetween startDate endDate NumDenumPosition.Numerator / dc.DaysInYear startDate endDate