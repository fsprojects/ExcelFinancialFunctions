using NUnit.Framework;
using Excel.FinancialFunctions;
using System;

namespace ExcelFinancialFunctions.ReleaseTests
{
    public class SpotTests
    {
        [Test(ExpectedResult = -0.67428578540657702)]
        public double YieldIssue8()
            => Financial.Yield(new DateTime(2015, 9, 21), new DateTime(2015, 10, 15), 0.04625, 105.124, 100.0, Frequency.SemiAnnual, DayCountBasis.UsPsa30_360);

        /*
        [<Test>]
        let YieldIssue8() =
            let param = DateTime(2015,9,21), DateTime(2015,10,15), 0.04625, 105.124, 100. , Frequency.SemiAnnual, DayCountBasis.UsPsa30_360
            Financial.Yield param
            |> shouldEqual (sprintf "YieldIssue8(%A)" param) -0.67428578540657702

        [<Test>]
        let spotYield() =
            let param = 
                DateTime(2008, 2, 15), DateTime(2016, 11, 15), 0.0575, 95.04287, 100.0,
                Frequency.SemiAnnual, DayCountBasis.UsPsa30_360
            Financial.Yield param
            |> shouldEqual (sprintf "spotYield(%A)" param) 0.065
    
        [<Test>]
        let spotXnpv() =
            let param = 0.14, [1.;3.;4.], [DateTime(1970, 3, 2); DateTime(1988, 2, 3); DateTime(1999, 3, 5)]
            Financial.XNpv param
            |> shouldEqual (sprintf "xnpv(%A)" param) 1.375214
         */
    }
}