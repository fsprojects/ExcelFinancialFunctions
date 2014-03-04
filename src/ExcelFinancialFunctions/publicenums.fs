// All the enums exposed in the external API, I need to define them first becasue they are used in the internal part
#light
namespace Excel.FinancialFunctions

/// Indicates when payments are due (end/beginning of period)
type PaymentDue =
| EndOfPeriod = 0
| BeginningOfPeriod = 1

/// The type of Day Count Basis
type DayCountBasis =
/// US 30/360
| UsPsa30_360               = 0
/// Actual/Actual
| ActualActual              = 1
/// Actual/360
| Actual360                 = 2
/// Actual/365
| Actual365                 = 3
/// European 30/360
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
