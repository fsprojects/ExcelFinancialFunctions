# Excel Financial Functions

This is a .NET Standard library that provides the full set of financial functions from Excel. The main goal for the library is compatibility with Excel, by providing the same functions, with the same behaviour. Note though that this is not a wrapper over the Excel library; the functions have been re-implemented in managed code so that you do not need to have Excel installed to use this library.

[![NuGet Badge](https://img.shields.io/nuget/v/ExcelFinancialFunctions.svg?style=flat)](https://www.nuget.org/packages/ExcelFinancialFunctions/)
[![Build and Test](https://github.com/fsprojects/ExcelFinancialFunctions/actions/workflows/dotnet.yml/badge.svg)](https://github.com/fsprojects/ExcelFinancialFunctions/actions/workflows/dotnet.yml)

## Goal: Match Excel 

We replicate the results Excel would produce in every situation,
even in cases where we might disagree with Excel\'s approach. Please have a look at the [Compatibility](http://fsprojects.github.io/ExcelFinancialFunctions/compatibility.html) page for more detail on this topic.

Microsoft\'s official documentation on [Excel Functions](https://support.microsoft.com/en-us/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb) is the best place to learn more about how the functions should work. The scope for this library is the full set of functions in the "Financial Functions" category.

### Thoroughly tested

As of last count, the library is validated against 199,252 test cases.

* [ExcelFinancialFunctions.Tests](./tests/ExcelFinancialFunctions.Tests): Unit tests checking against previously-determined truth values from Excel 2010. Inputs and expected outputs are read from data files.
* [ExcelFinancialFunctions.ConsoleTests](./tests/ExcelFinancialFunctions.ConsoleTests): Test cases comparing the library results directly to running Excel code via interop. These should be run on a Windows machine with Excel 2013 (or later) installed.  

### You can help!

Found a discrepency? [Open an Issue](https://github.com/fsprojects/ExcelFinancialFunctions/issues)! Or better yet, a [Pull Request](https://github.com/fsprojects/ExcelFinancialFunctions/pulls).

## Adding it to your project

Excel Financial Functions is a .NET Standard 2.0 library, which you can add to any project
based on a .NET implementation which [supports the standard](https://docs.microsoft.com/en-us/dotnet/standard/net-standard). This includes .NET Framework 4.6.1 or later, and .NET Core 2.0 or later.

Simply add it from NuGet in the usual way:

```
PS> dotnet add package ExcelFinancialFunctions

  Determining projects to restore...
info : Adding PackageReference for package 'ExcelFinancialFunctions' into project 
info :   GET https://api.nuget.org/v3/registration5-gz-semver2/excelfinancialfunctions/index.json
info :   OK https://api.nuget.org/v3/registration5-gz-semver2/excelfinancialfunctions/index.json 69ms
info : Restoring packages for project.csproj...
info : PackageReference for package 'ExcelFinancialFunctions' version '2.4.1' added to file 'project.csproj'.
info : Committing restore...
log  : Restored project.csproj (in 72 ms).
```

## Using it

Even though the libary is written in F#, you can use it from any .NET language, including C#. The functions are provided as static methods on a Financial class in the Excel.FinancialFunctions namespace.

``` c#
using Excel.FinancialFunctions;

Console.WriteLine( Financial.IPmt(rate: 0.005, per: 53, nper: 360, pv: 500000, fv: 0, typ: PaymentDue.EndOfPeriod) );
// Displays -796.3747578439793

Console.WriteLine( Financial.Pmt(rate: 0.005, nper: 360, pv: 500000, fv: 0, typ: PaymentDue.EndOfPeriod) );
// Displays -1687.7136560969248
```

Or from F#:

```F#
open Excel.FinancialFunctions

printfn "%f" <| Financial.IPmt (0.005, 53., 180., 200000., 0., PaymentDue.EndOfPeriod) 
// Displays -796.374758

printfn "%f" <| Financial.Pmt (0.005, 180., 200000., 0., PaymentDue.EndOfPeriod) 
// Displays -1687.713656
```
