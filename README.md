ExcelFinancialFunctions
=======================
  
#### What is it?
This is a .NET library that provides the full set of financial functions from Excel. The main goal for the library is compatibility with Excel, by providing the same functions, with the same behaviour. Note though that this is not a wrapper over the Excel library; the functions have been re-implemented in managed code so that you do not need to have Excel installed to use this library.


#### Where I can find documentation on these functions?
Just open Excel and click on Formulas/Financial or go to this [link](http://office.microsoft.com/client/helppreview.aspx?AssetID=HP100791841033&ns=EXCEL&lcid=1033&CTT=3&Origin=HP100623561033)
There's also [API reference](http://fsprojects.github.io/ExcelFinancialFunctions/reference/index.html)

#### I don't think one of the function is right. Excel produces the wrong results! Why don't you do it right?
The goal is to replicate Excel results (right and wrong).  Feel free to contribute to the effort by coding what you think is the right solution


#### How do I use the library?
Just add ExcelFinancialFunctions.dll to the references in your project. The functions are provided as static methods on a Financial class in the Excel.FinancialFunctions namespace


#### Have you tested this thing?
Yes, there're 199,252 testcases running against it. You can easily raise that number significantly by adding new values to test in testdef.fs.  
_ExcelFinancialFunctions.ConsoleTests.sln_ contains the tests comparing the library results to Excel, so it should be installed on a test machine.  
_ExcelFinancialFunctions.Test.sln_ - the unit tests matching Excel 2010 (their parameters and results are read from files).  
You can find more information on compatibility [here](http://fsprojects.github.io/ExcelFinancialFunctions/compatibility.html).

#### Are there any functions that behave different from Excel?
Yes, there are two of them.

##### CoupDays
The Excel algorithm seems wrong in that it doesn't respect the following:

    coupDays = coupDaysBS + coupDaysNC.

This equality should stand. The result differs from Excel by +/- one or two days when the date spans a leap year.


##### VDB
In the excel version of this algorithm the depreciation in the period (0,1) is not the same as 
the sum of the depreciations in periods (0,0.5) (0.5,1).
    
    VDB(100,10,13,0,0.5,1,0) + VDB(100,10,13,0.5,1,1,0) <> VDB(100,10,13,0,1,1,0)

Notice that in Excel by using '1' (no_switch) instead of '0' as the last parameter everything works as expected.  The last parameter should have no influence in the calculation given that in the first period there is no switch to sln depreciation.


_(Note, the original version of the library is still available [here.](http://code.msdn.microsoft.com/office/Excel-Financial-functions-6afc7d42))_

#### Maintainer(s)

- [@luajalla](https://github.com/luajalla)

The default maintainer account for projects under "fsprojects" is [@fsgit](https://github.com/fsgit) - F# Community Project Incubation Space (repo management)
