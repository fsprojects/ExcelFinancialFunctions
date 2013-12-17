namespace System
open System.Reflection

[<assembly: AssemblyTitleAttribute("ExcelFinancialFunctions")>]
[<assembly: AssemblyProductAttribute("ExcelFinancialFunctions")>]
[<assembly: AssemblyDescriptionAttribute("A .NET library that provides the full set of financial functions from Excel.")>]
[<assembly: AssemblyVersionAttribute("2.0")>]
[<assembly: AssemblyFileVersionAttribute("2.0")>]
do ()

module internal AssemblyVersionInformation =
    let [<Literal>] Version = "2.0"
