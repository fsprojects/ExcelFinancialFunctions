// All the enums exposed in the external API, I need to define them first becasue they are used in the internal part
#light
namespace System.Numeric

type PaymentDue =
| EndOfPeriod = 0
| BeginningOfPeriod = 1

type DayCountBasis =
| UsPsa30_360               = 0
| ActualActual              = 1
| Actual360                 = 2
| Actual365                 = 3
| Europ30_360               = 4

type Frequency =
| Annual        = 1
| SemiAnnual    = 2
| Quarterly     = 4

type AccrIntCalcMethod =
| FromFirstToSettlement = 0
| FromIssueToSettlement = 1

type VdbSwitch =
| DontSwitchToStraightLine  = 1
| SwitchToStraightLine      = 0
