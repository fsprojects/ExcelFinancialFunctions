(*** hide ***)
#I "../../bin"
#r "ExcelFinancialFunctions.dll"

open System
open Excel.FinancialFunctions

(**
Compatibility
=============

This library replicates Excel behavior. There're 199,252 tests verifying the results against Excel 2010 and their number can be raised significantly by adding new test values. Several tests check the function properties, e.g. that bond duration can't be greater than maturity.    
The current version matches Excel 2010, which is slightly different from 2003, see [Function Improvements in Excel 2010](http://blogs.office.com/b/microsoft-excel/archive/2009/09/10/function-improvements-in-excel-2010.aspx). 
Note, that console tests require Excel, whereas the unit tests can be run even on mono - their parameters and expected results are stored in files. 
  

However, there're still some differences comparing to Excel. 
_Read more about OpenOffice vs Excel [here](openofficediff.html)._


COUPDAYS
--------

The Excel algorithm doesn't respect equality `coupDays = coupDaysBS + coupDaysNC`. The library result differs from Excel by +/- one or two days when the date spans a leap year. ([office docs](http://office.microsoft.com/en-us/excel/HP052090311033.aspx))  
*)

let settlement = DateTime(2012, 1, 1)
let maturity   = DateTime(2016, 2, 29)

let param = settlement, maturity, Frequency.SemiAnnual, DayCountBasis.ActualActual

let days = Financial.CoupDays param
let bs = Financial.CoupDaysBS param
let nc = Financial.CoupDaysNC param
// Excel: 2
days - bs - nc
// [fsi:val days : float = 182.0]
// [fsi:val bs : float = 123.0]
// [fsi:val nc : float = 59.0]
// [fsi:val it : float = 0.0]


(** 
VDB
---

In the Excel version of this algorithm the depreciation in the period (0,1) is not the same as the sum of the depreciations in periods (0,0.5) (0.5,1)  
Notice that in Excel by using '1' (no_switch) instead of '0' as the last parameter everything works as expected.  
In truth, the last parameter should have no influence in the calculation given that in the first period there is no switch to sln depreciation.  
Overall, the algorithm is correct, even if it disagrees with Excel when startperiod is fractional. ([office docs](http://office.microsoft.com/en-us/excel/HP052093341033.aspx))  
*)

let vdb sp ep switch = 
    Financial.Vdb(100.0, 10.0, 13.0, sp, ep, 1.0, 
        if switch then VdbSwitch.SwitchToStraightLine else VdbSwitch.DontSwitchToStraightLine)

let p1 = vdb 0.0 0.5 false
let p2 = vdb 0.5 1.0 false
let total = vdb 0.0 1.0 false
// Excel: 0.1479
total - p1 - p2 
// [fsi:val p1 : float = 3.846153846]
// [fsi:val p2 : float = 3.846153846]
// [fsi:val total : float = 7.692307692]
// [fsi:val it : float = 0.0]

let p1sw = vdb 0.0 0.5 true
let p2sw = vdb 0.5 1.0 true
let totalsw = vdb 0.0 1.0 true
// Excel: 0.0000
totalsw - p1sw - p2sw 
// [fsi:val p1sw : float = 3.846153846]
// [fsi:val p2sw : float = 3.846153846]
// [fsi:val totalsw : float = 7.692307692]
// [fsi:val it : float = 0.0]


(**
AMORDEGRC
---------

ExcelCompliant is used because Excel stores 13 digits. AmorDegrc algorithm rounds numbers  
and returns different results unless the numbers get rounded to 13 digits before rounding them.  
I.E. 22.49999999999999 is considered 22.5 by Excel, but 22.4 by the .NET framework. ([office docs](http://office.microsoft.com/en-us/excel/HP052089841033.aspx))     
*)

let amorDegrc excelCompliant = 
    Financial.AmorDegrc(100.0, DateTime(2014,1,1), DateTime(2016,1,1), 
        50.0, 1.0, 0.3, DayCountBasis.ActualActual, excelCompliant)

amorDegrc true
// [fsi:val it : float = 23.0]
amorDegrc false
// [fsi:val it : float = 22.0]


(**
DDB
---

Excel Ddb has two interesting characteristics:  
1. It special cases ddb for fractional periods between 0 and 1 by considering them to be 1  
2. It is inconsistent with VDB(..., True) for fractional periods, even if VDB(..., True) is defined to be the same as ddb. The algorithm for VDB is theoretically correct.  
This function makes the same 1. adjustment.([office docs](http://office.microsoft.com/en-us/excel/HP052090511033.aspx))
*)


(**
RATE and ODDFYIELD
------------------

Excel uses a different root finding algo. Sometimes the library results are better, sometimes Excel's. ([office docs](http://office.microsoft.com/en-us/excel/HP052092321033.aspx))
*)

(**
XIRR and XNPV
-------------

XIRR and XNPV functions are related: the net present value, given the internal rate of return, should be zero.
However, XNPV works only for positive rates even though the XIRR results might be negative. 
The results can also be different because of the root finding functions. ([office docs](http://office.microsoft.com/en-us/excel/HP052093411033.aspx))
*)

let dates = [|DateTime(2000, 2, 29); DateTime(2000, 3, 31)|]
let values = [|206101714.849377; -156650972.54265|]
// Excel: -0.960452189
Financial.XIrr(values, dates, -0.1)
// [fsi:val it : float = -0.960452195]
// Excel: #NUM!
Financial.XNpv(-0.960452195, values, dates)
// [fsi:val it : float = -0.008917063475]
// Excel: #NUM!
Financial.XNpv(-0.960452189, values, dates)
// [fsi:val it : float = 2.646784514]

let values2 = [|15108163.3840923; -75382259.6628424|]
// Excel: #NUM!
Financial.XIrr(values2, dates, -0.1)
// [fsi:val it : float = 165601346.1]
// Excel: 165601345.6
Financial.XIrr(values2, dates, 0.1)
// [fsi:val it : float = 165601346.1]
Financial.XNpv(165601346.1, values2, dates)
// [fsi:val it : float = -0.000269997865]
Financial.XNpv(165601345.6, values2, dates)
// [fsi:val it : float = -0.004144238308]
