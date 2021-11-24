# Excel Financial Functions

This is a .NET Standard library that provides the full set of financial functions from Excel. The main goal for the library is compatibility with Excel, by providing the same functions, with the same behaviour. Note though that this is not a wrapper over the Excel library; the functions have been re-implemented in managed code so that you do not need to have Excel installed to use this library.

[![Build+Test+Docs](https://github.com/fsprojects/ExcelFinancialFunctions/actions/workflows/push-master.yml/badge.svg)](https://github.com/fsprojects/ExcelFinancialFunctions/actions/workflows/push-master.yml)
[![NuGet Badge](https://img.shields.io/nuget/v/ExcelFinancialFunctions.svg?style=flat)](https://www.nuget.org/packages/ExcelFinancialFunctions/)

## Goal: Match Excel 

We replicate the results Excel would produce in every situation,
even in cases where we might disagree with Excel\'s approach. Please have a look at the [Compatibility](http://fsprojects.github.io/ExcelFinancialFunctions/compatibility.html) page for more detail on this topic.

Microsoft\'s official documentation on [Excel Functions](https://support.microsoft.com/en-us/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb) is the best place to learn more about how the functions should work. The scope for this library is the full set of functions in the "Financial Functions" category.

### Thoroughly tested

As of last count, the library is validated against 199,252 test cases.

* [ExcelFinancialFunctions.Tests](./tests/ExcelFinancialFunctions.Tests): Unit tests checking against previously-determined truth values from Excel 2010. Inputs and expected outputs are read from data files.
* [ExcelFinancialFunctions.ConsoleTests](./tests/ExcelFinancialFunctions.ConsoleTests): Test cases comparing the library results directly to running Excel code via interop. These should be run on a Windows machine with Excel 2013 (or later) installed.  

### Difference #1: CoupDays

There are two notable areas where we judged that Excel was sufficiently incorrect
such that we needed to deviate from the primary goal of matching Excel precisely.

The first is the coupDays algorithm. Excel doesn't respect the following:

```
coupDays = coupDaysBS + coupDaysNC.
```

This equality should stand. The result differs from Excel by +/- one or two days when the date spans a leap year.

### Difference #2: VDB

In the excel version of this algorithm the depreciation in the period (0,1) is not the same as 
the sum of the depreciations in periods (0,0.5) (0.5,1).

```
VDB(100,10,13,0,0.5,1,0) + VDB(100,10,13,0.5,1,1,0) <> VDB(100,10,13,0,1,1,0)
```    

Notice that in Excel by using '1' (no_switch) instead of '0' as the last parameter everything works as expected.  The last parameter should have no influence in the calculation given that in the first period there is no switch to sln depreciation.

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

## Code of Conduct

This repository is governed by the [Contributor Covenant Code of Conduct](https://www.contributor-covenant.org/).

We pledge to be overt in our openness, welcoming all people to contribute, and pledging in return to value them as whole human beings and to foster an atmosphere of kindness, cooperation, and understanding.

## Library license

The library is available under Apache 2.0. For more information see the [License file](./LICENSE.txt).

## Maintainers

Current maintainer is [James Coliz](https://github.com/jcoliz).

Original author is [Luca Bolognese](https://github.com/lucabol). Historical maintainers of this project are [Natallie Baikevich](https://github.com/luajalla) and [Chris Pell](https://github.com/jcoliz). And of course, where would we be without [Don Syme](https://github.com/dsyme)?

The default maintainer account for projects under "fsprojects" is [@fsprojectsgit](https://github.com/fsprojectsgit) - F# Community Project Incubation Space (repo management)
