using NUnit.Framework;
using Excel.FinancialFunctions;
using System;

namespace ExcelFinancialFunctions.ReleaseTests
{
    [DefaultFloatingPointTolerance(1e-6)]
    public class SpotTests
    {
        [Test(ExpectedResult = -796.374758)]
        public double Readme1()
            => Financial.IPmt(rate: 0.005, per: 53, nper: 180, pv: 200000, fv: 0, typ: PaymentDue.EndOfPeriod);

        [Test(ExpectedResult = -1687.713656)]
        public double Readme2()
            => Financial.Pmt(rate: 0.005, nper: 180, pv: 200000, fv: 0, typ: PaymentDue.EndOfPeriod);

        [Test(ExpectedResult = -0.67428578540657702)]
        public double YieldIssue8()
            => Financial.Yield(new DateTime(2015, 9, 21), new DateTime(2015, 10, 15), 0.04625, 105.124, 100.0, Frequency.SemiAnnual, DayCountBasis.UsPsa30_360);

        [Test(ExpectedResult = 0.065)]
        public double spotYield()
            => Financial.Yield(new DateTime(2008, 2, 15), new DateTime(2016, 11, 15), 0.0575, 95.04287, 100.0, Frequency.SemiAnnual, DayCountBasis.UsPsa30_360);

        [Test(ExpectedResult = 1.375214)]
        public double spotXnpv()
            => Financial.XNpv(0.14, new double[] { 1.0, 3.0, 4.0 }, new DateTime[] { new DateTime(1970, 3, 2), new DateTime(1988, 2, 3), new DateTime(1999, 3, 5) });

        [Test(ExpectedResult = 90.0)]
        public double CoupDays()
            => Financial.CoupDays(new DateTime(1984, 3, 4), new DateTime(1990, 4, 5), Frequency.Quarterly, DayCountBasis.UsPsa30_360);

        [Test(ExpectedResult = 59.0)]
        public double CoupDaysBS()
            => Financial.CoupDaysBS(new DateTime(1984,3,4), new DateTime(1990,4,5),Frequency.Quarterly,DayCountBasis.UsPsa30_360 );

        [Test(ExpectedResult = 31.0)]
        public double CoupDaysNC()
            => Financial.CoupDaysNC(new DateTime(1984, 3, 4), new DateTime(1990, 4, 5), Frequency.Quarterly, DayCountBasis.UsPsa30_360);

        [Test(ExpectedResult = 25.0)]
        public double CoupNum()
            => Financial.CoupNum(new DateTime(1984, 3, 4), new DateTime(1990, 4, 5), Frequency.Quarterly, DayCountBasis.UsPsa30_360);

        [Test(ExpectedResult = 1.78125)]
        public double DollarDe()
            => Financial.DollarDe(1.125,16.0);

        [Test(ExpectedResult = 1.02)]
        public double DollarFr()
            => Financial.DollarFr(1.125, 16.0);

        [Test(ExpectedResult = 0.05354266737)]
        public double Effect()
            => Financial.Effect(0.0525, 4.0);

        [Test(ExpectedResult = 121.5236352)]
        public double FvSchedule()
            => Financial.FvSchedule(100.0, new double[] { 0.13, 0.14, -0.2, 0.34, -0.12 });

        [Test(ExpectedResult = 0.260952337)]
        public double Irr()
            => Financial.Irr(new double[] { -123.0, 12.0, 15.0, 50.0, 200.0 }, 0.14);

        [Test(ExpectedResult = -10.5)]
        public double ISPmt()
            => Financial.ISPmt(0.15, 3.0, 10.0, 100.0);

        [Test(ExpectedResult = 0.2409336873)]
        public double Mirr()
            => Financial.Mirr(new double[] { -123.0, 12.0, 15.0, 50.0, 200.0 }, 0.14, 0.12);

    }
}