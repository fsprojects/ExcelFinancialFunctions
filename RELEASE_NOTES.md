## 3.2.0 - May 10 2022
* Removes needless constraint on 0-value inputs to FV & PMT functions. (PR #67)

## 3.1.0 - Dec 20 2021
* Adds PDURATION function. Returns the number of periods required by an investment to reach a specified value. (Resolves #62)
* Adds RRI function. Also used for CAGR. Returns an equivalent interest rate for the growth of an investment. (Resolves #60)
* Improves XIRR function by reducing the precision required before an answer is returned. (Fixes #27)
* Improves ACCRINT function by allowing first interest date on the settlement date. (Fixes #22)
* Adds PriceAllowNegativeYield function. This operates like the PRICE function except that it allows negative yield inputs. It is experimental. We'd love feedback on how this works for folks. (Fixes #13)

## 3.0.0 - Dec 7 2021
* Retarget library onto .NET Standard 2.0
* Adds explicit support for .NET Core 2.0 and higher including 5.0 and 6.0
* Removes support for full .NET Framework 4.6 and lower

## 2.4.1
* Relaxed FSharp.Core dependency

## 2.4
* Only build profile 259 portable profile of the library (net40 not needed)

## 2.3
* Portable version of the library

## 2.2.1
* Price and yield functions now accept rate = 0

## 2.2
* Rename the top-level namespace to `Excel.FinancialFunctions`

## 2.1
* Move to github

## 2.0
* Fixed order of parameter and naming to the Rate function

## 1.0
* Fixed call to throw in bisection
* Changed findBounds algo
* Added TestXirrBugs function
* Removed the NewValue functions everywhere


