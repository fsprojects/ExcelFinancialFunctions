// All the enums exposed in the external API, I need to define them first becasue they are used in the internal part
#light
namespace System.Numeric

/// Indicates when payments are due (end/beginning of period)
type PaymentDue =
| EndOfPeriod = 0
| BeginningOfPeriod = 1

/// The type of Day Count Basis: US 30/360, Actual/actual, Actual/360, Actual/365 or European 30/360
type DayCountBasis =
| UsPsa30_360               = 0
| ActualActual              = 1
| Actual360                 = 2
| Actual365                 = 3
| Europ30_360               = 4

/// The number of coupon payments per year
type Frequency =
| Annual        = 1
| SemiAnnual    = 2
| Quarterly     = 4

/// Indicates whether accrued interest is computed from issue date (by default) or first interest to settlement
type AccrIntCalcMethod =
| FromFirstToSettlement = 0
| FromIssueToSettlement = 1

/// Specifies whether to switch to straight-line depreciation when depreciation is greater than the declining balance calculation
type VdbSwitch =
| DontSwitchToStraightLine  = 1
| SwitchToStraightLine      = 0
