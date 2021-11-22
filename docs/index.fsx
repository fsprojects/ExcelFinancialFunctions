(*** hide ***)
// This block of code is omitted in the generated HTML documentation. Use 
// it to define helpers that you do not want to show in the documentation.
#I "../src/ExcelFinancialFunctions/bin/Release/netstandard2.0"

(**

Excel Financial Functions
===================

This is a .NET library that provides the full set of financial functions from Excel. 
It can be used from both F# and C# as well as from other .NET languages.
The main goal for the library is compatibility with Excel, by providing the same functions, 
with the same behaviour. 

Note though that this is not a wrapper over the Excel library; the functions have been 
re-implemented in managed code so that you do not need to have Excel installed to use this library.

The package is available on <a href="https://nuget.org/packages/ExcelFinancialFunctions">NuGet</a>. [![NuGet Status](//img.shields.io/nuget/v/ExcelFinancialFunctions?style=flat)](https://www.nuget.org/packages/ExcelFinancialFunctions/)

You can also use `ExcelFinancialFunctions` in [dotnet interactive](https://github.com/dotnet/interactive) 
notebooks, in [Visual Studio Code](https://code.visualstudio.com/) 
or [Jupyter](https://jupyter.org/), or in F# scripts (`.fsx` files), 
by referencing the package as follows:

    #r "nuget: ExcelFinancialFunctions" // Use the latest version 

Example
-------

This example demonstrates using the YIELD function to calculate bond yield.

*)
#r "ExcelFinancialFunctions.dll"
open System
open Excel.FinancialFunctions

// returns 0.065 or 6.5%
Financial.Yield (DateTime(2008,2,15), DateTime(2016,11,15), 0.0575, 95.04287, 100.0, 
                 Frequency.SemiAnnual, DayCountBasis.UsPsa30_360)


(**

Samples & documentation
-----------------------

The library comes with comprehensible documentation. The tutorials and articles are
automatically generated from `*.fsx` files in [the docs folder][docs]. The API 
reference is automatically generated from Markdown comments in the library implementation.

* [API Reference](reference/index.html) contains automatically generated documentation for all types, modules
   and functions in the library. This includes the links to the Excel documentation.
* [Excel Compatibility](compatibility.html) section explains the possible differences with Excel's results. 
  
Contributing and copyright
--------------------------

The project is hosted on [GitHub][gh] where you can [report issues][issues], fork 
the project and submit pull requests. If you're adding new public API, please also 
consider adding [samples][content] that can be turned into a documentation. 

The library was originally developed by Luca Bolognese, the initial version can be
downloaded [here][msdn]. It is available under Apache License, for more information 
see the [License file][license] in the GitHub repository. 

  [content]: https://github.com/fsprojects/ExcelFinancialFunctions/tree/master/docs/content
  [gh]: https://github.com/fsprojects/ExcelFinancialFunctions
  [issues]: https://github.com/fsprojects/ExcelFinancialFunctions/issues
  [readme]: https://github.com/fsprojects/ExcelFinancialFunctions/blob/master/README.md
  [license]: https://github.com/fsprojects/ExcelFinancialFunctions/blob/master/LICENSE.txt
  [msdn]: http://code.msdn.microsoft.com/office/Excel-Financial-functions-6afc7d42
*)
