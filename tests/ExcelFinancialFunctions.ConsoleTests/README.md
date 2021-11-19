# Console Tests

These tests require Excel 2013 or later installed on your host machine. They directly
compare the results for many operations against Excel directly using [Excel Interop](https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel). A better name for 
them might be "Interop Tests".

They also take a very long time to run, as the test matrices get quite complex. For example on my
12-core machine, they run for 15 minutes.

You can run just the "Fast" tests as a smoke test:

```
PS tests\ExcelFinancialFunctions.ConsoleTests> dotnet build

Microsoft (R) Build Engine version 16.11.0+0538acc04 for .NET
Copyright (C) Microsoft Corporation. All rights reserved.

  ExcelFinancialFunctions -> \src\ExcelFinancialFunctions\bin\Debug\netstandard2.0\ExcelFinancialFunctions.dll
  ExcelFinancialFunctions.ConsoleTests -> \tests\ExcelFinancialFunctions.ConsoleTests\bin\Debug\net48\ExcelFinancialFunctions.ConsoleTests.dll

Build succeeded.
    0 Warning(s)
    0 Error(s)

Time Elapsed 00:00:00.75

PS \tests\ExcelFinancialFunctions.ConsoleTests> vstest.console.exe bin\Debug\net48\ExcelFinancialFunctions.ConsoleTests.dll --TestCaseFilter:"Category=Fast"

Microsoft (R) Test Execution Command Line Tool Version 16.11.0
Copyright (c) Microsoft Corporation.  All rights reserved.

A total of 1 test files matched the specified pattern.
NUnit Adapter 4.1.0.0: Test execution started
Running selected tests in \tests\ExcelFinancialFunctions.ConsoleTests\bin\Debug\net48\ExcelFinancialFunctions.ConsoleTests.dll
   NUnit3TestExecutor discovered 69 of 69 NUnit test cases using Current Discovery mode, Non-Explicit run
  Passed RunMatrix("IRR") [912 ms]
...
Test Run Successful.
Total tests: 69
     Passed: 69
 Total time: 5.2782 Seconds
```