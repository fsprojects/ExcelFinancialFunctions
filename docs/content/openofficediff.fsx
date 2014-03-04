(*** hide ***)
#I "../../bin"
#r "ExcelFinancialFunctions.dll"

open System
open Excel.FinancialFunctions

(**
Difference between OpenOffice and the library
=============================================

The library was designed to be Excel-compliant (see [Compatibility](compatibility.html) section), therefore its behavior is different from OpenOffice/LibreOffice. 
Most of the differences are because of the day count conventions and root finding algorithm implementation details.
Some examples are provided below. 
*)

(**
ODDFYIELD, ODDFPRICE
--------------------

According to the [OO wiki](https://wiki.openoffice.org/wiki/Documentation/How_Tos/Calc:_ODDFYIELD_function), 
these functions currently return invalid results (it's #VALUE! even for valid inputs). 
*)

Financial.OddFYield(DateTime(2008, 12, 11), DateTime(2021, 4, 1), DateTime(2008, 10, 15), 
    DateTime(2009, 4, 1), 0.06, 100., 100., Frequency.Quarterly, DayCountBasis.ActualActual)
// [fsi:Excel: 0.059976999]
Financial.OddFPrice(DateTime(1999, 2, 28), DateTime(2010, 6, 30), DateTime(1998, 2, 28),
    DateTime(2009, 6, 30), 0.07, 0.03, 100., Frequency.Annual, DayCountBasis.Actual360)
// [fsi:Excel: 127.9031274]

(**
ODDLYIELD, OODLPRICE
--------------------

The functions return different results. ODDLYIELD example can be found [here](https://wiki.openoffice.org/wiki/Documentation/How_Tos/Calc:_ODDLYIELD_function).
*)

Financial.OddLPrice(DateTime(1999, 2, 28), DateTime(2000, 2, 28), DateTime(1998, 2, 28),
    0.07, 0.03, 130., Frequency.SemiAnnual, DayCountBasis.Actual360)
// [fsi:Excel: 132.8058252  LibreOffice: 132.8407748124]

Financial.OddLYield(DateTime(1990, 6, 1), DateTime(1995, 12, 31), DateTime(1990, 1, 1),
    0.002, 103., 100., Frequency.Quarterly, DayCountBasis.ActualActual)
// [fsi:Excel: -0.00327563  LibreOffice: -0.002925876]
// Returns the same value even though the frequency is different
Financial.OddLYield(DateTime(1990, 6, 1), DateTime(1995, 12, 31), DateTime(1990, 1, 1),
    0.002, 103., 100., Frequency.Annual, DayCountBasis.ActualActual)
// [fsi:Excel: -0.00327205  LibreOffice: -0.002925876] 

(**
ACCRINT, DISC, DURATION, PRICE, YIELD, INTRATE, TBILL* and others
---------------------------------------------------------

Most likely the differences can be explained with YEARFRAC/day count implementations. 
DURATION in OO is a completely different function, the analog of Excel one is called DURATION\_ADD.
*)
// in our tests, the numbers for European 30/360 basis were the same.
let accrint basis = 
    Financial.AccrInt(DateTime(1990, 3, 4), DateTime(1993, 3, 31), 
        DateTime(1992, 3, 4), 0.07, 10000., Frequency.SemiAnnual, basis)

accrint DayCountBasis.UsPsa30_360
// [fsi:Excel: 1401.944444  LibreOffice: 1400.000000]
accrint DayCountBasis.ActualActual
// [fsi:Excel: 1398.076923  LibreOffice: 1401.917808]
accrint DayCountBasis.Actual360
// [fsi:Excel: 1394.166667  LibreOffice: 1421.388889]
accrint DayCountBasis.Actual365
// [fsi:Excel: 1399.041096  LibreOffice: 1401.917808]

Financial.AccrIntM(DateTime(1990, 3, 4), DateTime(2010, 6, 5), 0.1, 12030.34, DayCountBasis.ActualActual)
// [fsi:Excel: 24367.7909   LibreOffice: 24383.68639]

Financial.Disc(DateTime(2003, 2, 14), DateTime(2004, 3, 31), 23., 100., DayCountBasis.ActualActual)
// [fsi:Excel: 0.684757     LibreOffice: 0.683820]

Financial.Duration(DateTime(1980, 2, 15), DateTime(2000, 2, 28), 100., 0.03, 
    Frequency.Annual, DayCountBasis.Actual360)
// [fsi:Excel: 8.949173     LibreOffice: 9.254729]

Financial.MDuration(DateTime(1980, 2, 15), DateTime(2000, 2, 28), 100., 0.03,
    Frequency.SemiAnnual, DayCountBasis.Actual360)
// [fsi:Excel: 8.860247     LibreOffice: 9.158550]

Financial.Price(DateTime(1980, 2, 15), DateTime(2000, 2, 28), 0.07, 0.1, 100., 
    Frequency.Annual, DayCountBasis.Actual360)
// [fsi:Excel: 74.442516    LibreOffice: 74.334983]    

Financial.PriceDisc(DateTime(1980, 2, 15), DateTime(2000, 2, 28), 0.01, 100., DayCountBasis.ActualActual)
// [fsi:Excel: 79.966367    LibreOffice: 80.005464] 

Financial.IntRate(DateTime(1980, 2, 15), DateTime(1980, 5, 4), 23., 130., DayCountBasis.UsPsa30_360)
// [fsi:Excel: 21.199780    LibreOffice: 21.471572]

Financial.TBillPrice(DateTime(1980, 2, 15), DateTime(1980, 3, 15), 2.)
// [fsi:Excel: 83.888889    LibreOffice: 82.777778]

Financial.Received(DateTime(1980, 2, 15), DateTime(2000, 2, 28), 200., 0.01, DayCountBasis.ActualActual)
// [fsi:Excel: 250.105148   LibreOffice: 249.982925]

// the only function which seems to behave differently in Excel Office for Mac
Financial.AmorDegrc(100., DateTime(1998, 2, 28), DateTime(2000, 2, 29), 
    10., 0.3, 0.15, DayCountBasis.Actual365, true)
// [fsi:Excel: 0    Excel for Mac: -2   LibreOffice: 75]

(**
AMORLINC
--------

As stated [here](https://wiki.openoffice.org/wiki/Documentation/How_Tos/Calc:_AMORLINC_function), 
when the date of purchase is the end of a period, Excel regards the initial period 0 as the first full period, 
whereas OO regards the initial period as of zero length and returns 0.  
However, there're other differences too.
*)

Financial.AmorLinc(100., DateTime(1998, 2, 28), DateTime(2000, 2, 29), 10., 0., 0.07, DayCountBasis.Actual365)
// [fsi:Excel: 14.000000   LibreOffice: 14.019178]
Financial.AmorLinc(100., DateTime(1998, 2, 28), DateTime(2009, 6, 30), 50., 1.7, 0.1, DayCountBasis.UsPsa30_360)
// [fsi:Excel: 0.0000000   LibreOffice: -63.33333]

(**
COUPDAYS, COUPDAYSBS, COUPDAYSNC
--------------------------------

In Excel the equality `coupDays = coupDaysBS + coupDaysNC` doesn't necessary hold when basis is other than Actual/Actual.
*)
let cdParam = DateTime(1980, 2, 15), DateTime(2000, 2, 28), Frequency.Annual, DayCountBasis.UsPsa30_360

Financial.CoupDays cdParam 
Financial.CoupDaysBS cdParam
Financial.CoupDaysNC cdParam
// [fsi:Excel: 360 <> 345 + 13  LibreOffice: 360 = 345 + 15]

(** 
CUMIPMT, CUMPRINC
-----------------

OO analogs are called CUMIPMT\_ADD and CUMPRINC\_ADD (they're expected to be [compatible with Excel](https://wiki.openoffice.org/wiki/Documentation/How_Tos/Calc:_CUMIPMT_ADD_function))
*)
Financial.CumIPmt(0.6, 10., 100., 1.3, 2., PaymentDue.EndOfPeriod)
// [fsi:Excel: -59.669577   LibreOffice: -119.669577]
Financial.CumPrinc(0.6, 10., 100., 1.3, 2., PaymentDue.EndOfPeriod)
// [fsi:Excel: -0.8811289   LibreOffice: -1.431834]

(**
DB, DDB
-------

Seems like DDB doesn't accept fractional periods.
*)
Financial.Db(100., 10., 1., 0.3, 1.)
// [fsi:Excel: 7.5  LibreOffice: 0.0]
Financial.Ddb(100., 10., 1., 0.3, 1.)
// [fsi:Excel: 90.0 LibreOffice: Err:502]

(** 
IRR, XIRR
---------

The implementation details of root finding algorithms might be the cause of differences. 
We didn't check all the tests, because OO doesn't accept arrays as parameters (e.g. {-100;100}). But some of them don't work anyway.
*)
Financial.XIrr([206101714.849377; -156650972.54265], [DateTime(2001, 2, 28); DateTime(2001, 3, 31)], -0.1)
// [fsi:Excel: -0.960452    LibreOffice: Err:502 (Invalid argument)]