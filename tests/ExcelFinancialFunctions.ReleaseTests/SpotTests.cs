using NUnit.Framework;
using Excel.FinancialFunctions;
using System;

namespace ExcelFinancialFunctions.ReleaseTests
{
    [DefaultFloatingPointTolerance(1e-6)]
    public class SpotTests
    {
        [Test(ExpectedResult = -0.67428578540657702)]
        public double YieldIssue8()
            => Financial.Yield(new DateTime(2015, 9, 21), new DateTime(2015, 10, 15), 0.04625, 105.124, 100.0, Frequency.SemiAnnual, DayCountBasis.UsPsa30_360);

        [Test(ExpectedResult = 0.065)]
        public double spotYield()
            => Financial.Yield(new DateTime(2008, 2, 15), new DateTime(2016, 11, 15), 0.0575, 95.04287, 100.0, Frequency.SemiAnnual, DayCountBasis.UsPsa30_360);

        [Test(ExpectedResult = 1.375214)]
        public double spotXnpv()
            => Financial.XNpv(0.14, new double[] { 1.0, 3.0, 4.0 }, new DateTime[] { new DateTime(1970, 3, 2), new DateTime(1988, 2, 3), new DateTime(1999, 3, 5) });
    }
}