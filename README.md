# CVA Estimation for a Five-Year Interest-Rate Swap

This repository contains a macro-enabled Excel implementation of a bilateral counterparty-credit-risk workflow for a five-year fixed-for-floating interest-rate swap (IRS). The model calibrates an initial term structure, values the swap, simulates short rates under a one-factor Vasicek process, revalues the swap across 10,000 scenarios, constructs future-exposure profiles, and aggregates the profiles into CVA-style measures.

> **Model-risk notice:** this is a prototype workbook, not a production pricing library. The repository documentation distinguishes the intended quantitative methodology from the workbook's exact implementation. Review the [known implementation issues](#known-implementation-and-model-risk-issues) before recalculating or relying on the results.

## Repository contents

- [`IRS_CVA2.xlsm`](IRS_CVA2.xlsm) — the Excel model, including VBA functions, simulation procedures, cached scenario matrices, charts, and an in-workbook `README` navigation sheet.
- [`README.md`](README.md) — this mathematical and operational reference.

## Table of contents

- [Scope and assumptions](#scope-and-assumptions)
- [How to navigate the workbook](#how-to-navigate-the-workbook)
- [Calculation architecture](#calculation-architecture)
- [Equation symbology](#equation-symbology)
- [Mathematical framework](#mathematical-framework)
- [Technical guide to every sheet](#technical-guide-to-every-sheet)
- [VBA procedures and recalculation order](#vba-procedures-and-recalculation-order)
- [Known implementation and model-risk issues](#known-implementation-and-model-risk-issues)
- [Validation checklist](#validation-checklist)

## Scope and assumptions

The workbook models a single IRS under the following configuration:

| Item | Workbook setting |
|---|---:|
| Notional | Normalized to 1 |
| Swap tenor | 5 years |
| Fixed-leg frequency | Quarterly |
| Floating-leg frequency | Quarterly |
| Simulation step | $\Delta t=0.25$ years |
| Simulation dates | $t_j=0,0.25,\ldots,5$ |
| Number of time points | 21, including time zero |
| Monte Carlo scenarios | $N=10{,}000$ |
| Interest-rate model | One-factor Vasicek |
| Curve representation | Nelson–Siegel–Svensson plus Vasicek ZCB fit |
| Credit inputs | Recovery rates and CDS spreads for parties A and B |
| Collateral/netting | Not modelled |
| Wrong-way risk | Not modelled; rate and credit inputs are not jointly simulated |

The sign convention is:

$$
\mathrm{FE}_A(t)=\max(V_{\mathrm{swap}}(t),0), \qquad
\mathrm{FE}_B(t)=\max(-V_{\mathrm{swap}}(t),0).
$$

A positive swap value is therefore exposure for party A; a negative value is represented as a positive exposure for party B.

## How to navigate the workbook

Open [`IRS_CVA2.xlsm`](IRS_CVA2.xlsm) in desktop Microsoft Excel. The first worksheet is `README`.

1. Start on `README` and use the **Open →** links to move directly to any model sheet.
2. Review inputs and curve calibration in `TERM structures`.
3. Review the inception swap price and par fixed rate in `IRS`.
4. Check scenario settings and short-rate paths in `MonteCarlo`.
5. Follow the scenario calculation through `Term_Structures`, `Annuity`, `PV_Swap`, and `Future Exposure`.
6. Return to `Main` for expected exposure, maximum exposure, percentiles, survival profiles, and CVA totals.

The workbook contains two similarly named sheets with different purposes:

- `TERM structures` — initial market-curve calibration and credit/survival construction.
- `Term_Structures` — the pathwise simulated discount-factor matrix.

The file contains VBA. Excel can display cached values with macros disabled, but VBA user-defined functions and simulation outputs will not reliably refresh. Enable macros only after reviewing the VBA and the implementation warnings below.

## Calculation architecture

```text
TERM structures -> IRS -> MonteCarlo -> Term_Structures
                                              |
                                              v
                                           Annuity
                                              |
                                              v
                                           PV_Swap
                                              |
                                              v
                                       Future Exposure ----------------> Main
TERM structures -- credit/survival ------------------------------------> Main
MonteCarlo -- expected rate and MMA -----------------------------------> Main
```

Each row in the large simulation sheets represents one Monte Carlo scenario. Each time column represents one quarterly date. `Main` reduces these matrices across scenarios into exposure and valuation-adjustment profiles.

## Equation symbology

Rates and spreads are represented as decimals in the equations unless explicitly stated otherwise. For example, 2% is written as $0.02$. Times and maturities are measured in years, present values are expressed per unit of notional, and discount factors and probabilities are dimensionless. A subscript identifies a payment date or scenario; an argument in parentheses identifies the time at which a quantity is evaluated.

### Indices, dates, and general operators

| Symbol | Meaning | Unit or domain |
|---|---|---|
| $t$ | Current valuation or observation time | Years |
| $T$ | Cash-flow or bond maturity date | Years |
| $\tau=T-t$ | Remaining time from valuation to maturity | Years |
| $t_j$ | The $j$-th Monte Carlo observation date | Years |
| $T_i$ | The $i$-th contractual swap payment date | Years |
| $\Delta t$ | Simulation time step; quarterly in the workbook | $0.25$ years |
| $i$ | Market-instrument or swap-payment index | Positive integer |
| $j$ | Simulation-time index | $j=0,\ldots,J$ |
| $n$ | Monte Carlo scenario index | $n=1,\ldots,N$ |
| $N$ | Number of Monte Carlo scenarios | $10{,}000$ in the workbook |
| $J$ | Number of simulation intervals | $20$ quarterly intervals over five years |
| $X$ | Counterparty label | $X\in\{A,B\}$ |
| $\sum$ | Sum over the stated index | Operator |
| $\max(x,0)$ | Positive part of $x$ | Same unit as $x$ |
| $\exp(x)=e^x$ | Exponential function | Operator |

### Nelson–Siegel–Svensson curve

| Symbol | Meaning | Unit or domain |
|---|---|---|
| $m$ | Market-curve maturity supplied to `NSS_Spot` | Years |
| $y(m)$ | NSS-fitted spot-rate value at maturity $m$ | Rate |
| $y_i^{\mathrm{market}}$ | Observed market rate for instrument $i$ | Rate |
| $y_i^{\mathrm{NSS}}$ | NSS model rate corresponding to instrument $i$ | Rate |
| $\beta_0$ | Long-run level factor | Rate |
| $\beta_1$ | Short-end slope factor | Rate |
| $\beta_2$ | First curvature factor | Rate |
| $\beta_3$ | Second curvature factor | Rate |
| $\tau_1,\tau_2$ | NSS decay parameters controlling where the factor loadings peak | Years; positive |
| $L(m,\tau)$ | NSS level/slope loading, $(1-e^{-m/\tau})/(m/\tau)$ | Dimensionless |
| $\mathrm{SSE}_{\mathrm{NSS}}$ | Sum of squared market-versus-NSS rate errors | Rate squared |

Here, $m$ means curve maturity. In the Monte Carlo section, $m_{n,j}$ instead denotes a conditional mean and is always written with both scenario and time subscripts.

### Vasicek short-rate model and zero-coupon bonds

| Symbol | Meaning | Unit or domain |
|---|---|---|
| $r_t$ | Continuously compounded instantaneous short rate at time $t$ | Rate |
| $r_0$ | Initial short rate used to start the model | Rate |
| $dr_t$ | Infinitesimal change in the short rate | Rate |
| $\alpha$ | Speed at which $r_t$ mean-reverts toward $\mu$ | Per year; normally $>0$ |
| $\mu$ | Long-run mean level of the short rate | Rate |
| $\sigma$ | Instantaneous short-rate volatility | Rate per square-root year |
| $W_t$ | Standard Brownian motion | Stochastic process |
| $dW_t$ | Brownian increment with variance $dt$ | Square-root years |
| $dt$ | Infinitesimal time increment | Years |
| $B(\tau)$ | Vasicek duration loading multiplying the current short rate | Years |
| $A(\tau)$ | Vasicek affine intercept in the bond-price exponent | Dimensionless |
| $P(t,T)$ | Price at $t$ of a zero-coupon bond paying one unit at $T$ | Discount factor |
| $P_i^{\mathrm{market}}$ | Market-implied discount factor for maturity $i$ | Discount factor |
| $P_i^{\mathrm{Vasicek}}$ | Vasicek discount factor for maturity $i$ | Discount factor |
| $\mathrm{SSE}_{\mathrm{Vasicek}}$ | Sum of squared market-versus-model discount-factor errors | Dimensionless |

The affine notation $A(\tau)$ in the bond formula is distinct from party $A$ and from the swap annuity $A_N$; its argument or subscript identifies the intended meaning.

### Monte Carlo simulation

| Symbol | Meaning | Unit or domain |
|---|---|---|
| $r_{n,j}$ | Simulated short rate in scenario $n$ at date $t_j$ | Rate |
| $m_{n,j}$ | Conditional mean of $r_{n,j}$ given $r_{n,j-1}$ | Rate |
| $U_{n,j}$ | Independent uniform random draw | $U(0,1)$ |
| $\Phi^{-1}(\cdot)$ | Inverse standard-normal cumulative distribution function | Operator |
| $Z_{n,j}$ | Standard-normal shock, $\Phi^{-1}(U_{n,j})$ | $N(0,1)$ |
| $s_{\mathrm{workbook}}$ | Shock standard deviation implemented by the VBA routine | Rate |
| $s_{\mathrm{standard}}$ | Shock standard deviation in the standard exact Vasicek transition | Rate |
| $\bar r(t_j)$ | Cross-scenario mean short rate at date $t_j$ | Rate |
| $I(t_j)$ | Trapezoidal approximation to $\int_0^{t_j}\bar r(u)\,du$ | Dimensionless |
| $u$ | Dummy time variable inside the rate integral | Years |
| $M(t_j)$ | Workbook money-market-account value, $e^{I(t_j)}$ | Dimensionless |

### Interest-rate swap valuation

| Symbol | Meaning | Unit or domain |
|---|---|---|
| $P_i=P(0,T_i)$ | Time-zero discount factor for payment date $T_i$ | Discount factor |
| $P_0$ | Discount factor at inception; normally equal to one | Discount factor |
| $P_N$ | Discount factor at the final payment date | Discount factor |
| $F_i$ | Unannualized forward return over payment period $i$ | Rate for the period |
| $\delta_i$ | Year fraction for payment period $i$ | Years |
| $\delta$ | Constant quarterly accrual factor used in the workbook | $0.25$ years |
| $\mathrm{PV}_{\mathrm{float}}$ | Present value of the floating-rate leg per unit notional | Value per unit notional |
| $A_N$ | Time-zero fixed-leg annuity through the final payment date | Years |
| $K$ | Par fixed swap rate determined at inception | Rate per year |
| $V_{\mathrm{swap}}(0)$ | Net receive-floating/pay-fixed swap value at inception | Value per unit notional |
| $P_{n,j}$ | Scenario-$n$ discount factor for the maturity represented by column $j$ | Discount factor |
| $A_{n,j}$ | Cumulative fixed-leg annuity in scenario $n$ through column $j$ | Years |
| $\mathrm{remaining}(j)$ | Reverse-column mapping to the swap's remaining maturity at exposure date $t_j$ | Column index |
| $V_{n,j}$ | Scenario swap value at exposure date $t_j$ | Value per unit notional |

All valuation equations assume unit notional. For a contractual notional $\mathcal N$, multiply swap values and exposures by $\mathcal N$.

### Exposure measures

| Symbol | Meaning | Unit or domain |
|---|---|---|
| $\mathrm{FE}_{A,n}(t_j)$ | Positive exposure to party $A$, $\max(V_{n,j},0)$ | Value per unit notional |
| $\mathrm{FE}_{B,n}(t_j)$ | Positive exposure in the opposite direction, $\max(-V_{n,j},0)$ | Value per unit notional |
| $\mathrm{EE}_A(t_j)$ | Scenario-average positive exposure to $A$ | Value per unit notional |
| $\mathrm{EE}_B(t_j)$ | Scenario-average positive exposure to $B$ | Value per unit notional |
| $\mathrm{MAXFE}_X(t_j)$ | Maximum scenario exposure for party $X$ at $t_j$ | Value per unit notional |
| $\mathrm{PFE}_{X,q}(t_j)$ | Empirical $q$-quantile of exposure for party $X$ at $t_j$ | Value per unit notional |
| $q$ | Selected confidence level for PFE | $0.90$, $0.95$, or $0.99$ |

The workbook rows labelled `PFE(A)` and `PFE(B)` report $\mathrm{MAXFE}$, not a percentile. The separately labelled 90%, 95%, and 99% rows contain the empirical percentile calculations.

### Credit risk and valuation adjustment

| Symbol | Meaning | Unit or domain |
|---|---|---|
| $s_X$ | CDS spread supplied for party $X$ | Rate per year |
| $R_X$ | Assumed recovery fraction for party $X$ after default | Between 0 and 1 |
| $\mathrm{LGD}_X=1-R_X$ | Loss-given-default fraction for party $X$ | Between 0 and 1 |
| $Q_X(t)$ | Workbook-implied survival measure for party $X$ through time $t$ | Probability-like value |
| $Q_X(t_{j-1})-Q_X(t_j)$ | Workbook marginal default measure over interval $(t_{j-1},t_j]$ | Probability-like value |
| $\mathrm{CVA}^{A}_j$ | Period-$j$ adjustment associated with positive exposure to $A$ and default of $B$ | Value per unit notional |
| $\mathrm{CVA}^{B}_j$ | Period-$j$ opposite-direction, DVA-like adjustment associated with default of $A$ | Value per unit notional |
| $\mathrm{CVA}^{A},\mathrm{CVA}^{B}$ | Sums of the corresponding period adjustments | Value per unit notional |
| $\mathrm{DVA}$ | Debit valuation adjustment: own-default benefit from one party's viewpoint | Value per unit notional |
| $\mathrm{BCVA}$ | Bilateral adjustment, commonly represented as CVA minus DVA under a stated sign convention | Value per unit notional |
| $\mathcal N$ | Optional contractual notional used to scale unit-notional results | Currency amount |

Superscripts $A$ and $B$ identify the exposure/default direction used by the workbook; they are labels, not mathematical powers. Because CVA/DVA sign conventions vary, the economic viewpoint should always be stated when reporting a bilateral result.

## Mathematical framework

### 1. Nelson–Siegel–Svensson curve

For maturity $m$, the VBA function `NSS_Spot` uses

$$
L(m,\tau)=\frac{1-e^{-m/\tau}}{m/\tau},
$$

and

$$
y(m)=\beta_0
+\beta_1L(m,\tau_1)
+\beta_2\left[L(m,\tau_1)-e^{-m/\tau_1}\right]
+\beta_3\left[L(m,\tau_2)-e^{-m/\tau_2}\right].
$$

`TERM structures!F3:F35` divides this result by 100. The fit error against the market observations in column C is

$$
\mathrm{SSE}_{\mathrm{NSS}}
=\sum_i\left(y_i^{\mathrm{market}}-y_i^{\mathrm{NSS}}\right)^2,
$$

reported in `O8`.

The VBA includes separate `tau1 = 0` and `tau2 = 0` branches. Those branches contain positive-exponent terms, $e^{+m/\tau}$, and should be reviewed before zero-decay parameters are permitted.

### 2. Vasicek short-rate model

The conceptual one-factor model is

$$
dr_t=\alpha(\mu-r_t)\,dt+\sigma\,dW_t.
$$

The VBA zero-coupon function has the signature

```text
Vasicek_ZCB(tau, r, mu, alpha, sigma)
```

and calculates

$$
B(\tau)=\frac{1-e^{-\alpha\tau}}{\alpha},
$$

$$
A(\tau)=
(B(\tau)-\tau)\frac{\alpha^2\mu-\sigma^2/2}{\alpha^2}
-\frac{\sigma^2B(\tau)^2}{4\alpha},
$$

$$
P(t,T)=\exp\left(A(\tau)-B(\tau)r_t\right).
$$

The initial Vasicek calibration objective is represented by

$$
\mathrm{SSE}_{\mathrm{Vasicek}}
=\sum_i\left(P_i^{\mathrm{Vasicek}}-P_i^{\mathrm{market}}\right)^2.
$$

The workbook stores this objective in `TERM structures!M6`; hidden Solver names identify `M2:M5` as adjustable cells.

### 3. Monte Carlo transition

For scenario $n$ and date $t_j$, the VBA computes the conditional mean

$$
m_{n,j}=e^{-\alpha\Delta t}r_{n,j-1}
+\left(1-e^{-\alpha\Delta t}\right)\mu,
$$

then samples

$$
r_{n,j}=m_{n,j}+s_{\mathrm{workbook}}Z_{n,j},
\qquad Z_{n,j}=\Phi^{-1}(U_{n,j}),\quad U_{n,j}\sim U(0,1).
$$

The shock scale in `MonteCarlo!B8` is implemented as

$$
s_{\mathrm{workbook}}
=\sqrt{\frac{\sigma^2\alpha}{2}
\left(1-e^{-2\alpha\Delta t}\right)}.
$$

For comparison, the standard exact Vasicek transition uses

$$
s_{\mathrm{standard}}
=\sqrt{\frac{\sigma^2}{2\alpha}
\left(1-e^{-2\alpha\Delta t}\right)}.
$$

This difference is material and is listed as a model-validation issue.

The scenario mean and its time integral are

$$
\bar r(t_j)=\frac{1}{N}\sum_{n=1}^{N}r_{n,j},
$$

$$
I(t_j)=I(t_{j-1})
+\frac{\Delta t}{2}\left[\bar r(t_{j-1})+\bar r(t_j)\right],
$$

with workbook money-market-account value

$$
M(t_j)=e^{I(t_j)}.
$$

### 4. IRS valuation

Let $T_i$ be the quarterly payment dates and $P_i=P(0,T_i)$. The workbook's periodic forward return is

$$
F_i=\frac{P_{i-1}}{P_i}-1.
$$

The floating-leg present value is

$$
\mathrm{PV}_{\mathrm{float}}
=\sum_{i=1}^{N}P_iF_i
=P_0-P_N
\approx 1-P_N.
$$

The fixed-leg annuity and par fixed rate are

$$
A_N=\sum_{i=1}^{N}\delta_iP_i,
$$

$$
K=\frac{\mathrm{PV}_{\mathrm{float}}}{A_N}.
$$

The inception swap value is

$$
V_{\mathrm{swap}}(0)
=\mathrm{PV}_{\mathrm{float}}-KA_N
\approx 0.
$$

### 5. Scenario annuity and swap value

The scenario annuity is accumulated as

$$
A_{n,1}=\delta P_{n,1},
$$

$$
A_{n,j}=A_{n,j-1}+\delta P_{n,j}.
$$

The `PV_Swap_sim` VBA routine reads the discount-factor and annuity arrays in reverse maturity order for the outstanding swap and uses

$$
V_{n,j}=\left(1-P_{n,\mathrm{remaining}(j)}\right)
-K A_{n,\mathrm{remaining}(j)}.
$$

The endpoints are set equal to `IRS!F5`, the inception residual value.

### 6. Exposure statistics

For each quarterly date,

$$
\mathrm{EE}_A(t_j)
=\frac{1}{N}\sum_{n=1}^{N}\max(V_{n,j},0),
$$

$$
\mathrm{EE}_B(t_j)
=\frac{1}{N}\sum_{n=1}^{N}\max(-V_{n,j},0).
$$

The rows labelled `PFE(A)` and `PFE(B)` actually calculate the scenario maximum:

$$
\mathrm{MAXFE}_X(t_j)=\max_n \mathrm{FE}_{X,n}(t_j).
$$

Separate rows calculate empirical 90%, 95%, and 99% quantiles using Excel's `PERCENTILE` function.

### 7. Credit and CVA equations used on `Main`

For party $X$, `Main` maps the CDS spread $s_X$ and recovery $R_X$ to

$$
Q_X(t)=\frac{e^{-s_Xt}-R_X}{1-R_X}.
$$

This is the workbook's exact formula; it is not the usual constant-hazard approximation $Q(t)=e^{-s t/(1-R)}$.

The period contributions implemented in rows 9 and 10 are

$$
\mathrm{CVA}^{A}_j
=(1-R_B)Q_B(t_j)
\left[Q_B(t_{j-1})-Q_B(t_j)\right]
\frac{\mathrm{EE}_A(t_j)}{M(t_j)},
$$

$$
\mathrm{CVA}^{B}_j
=(1-R_A)Q_A(t_j)
\left[Q_A(t_{j-1})-Q_A(t_j)\right]
\frac{\mathrm{EE}_B(t_j)}{M(t_j)}.
$$

Totals are

$$
\mathrm{CVA}^{A}=\sum_j\mathrm{CVA}^{A}_j,
\qquad
\mathrm{CVA}^{B}=\sum_j\mathrm{CVA}^{B}_j.
$$

From one party's viewpoint, the opposite-direction quantity is DVA-like. The workbook does not explicitly form a single bilateral value such as $\mathrm{BCVA}=\mathrm{CVA}-\mathrm{DVA}$.

## Technical guide to every sheet

### `README`

This is the first worksheet and the workbook's navigation page. It contains:

- a quick-start sequence;
- internal links to all calculation sheets;
- one-line descriptions of each sheet's role;
- the end-to-end dependency flow; and
- macro-disabled and recalculation warnings.

It does not participate in calculations.

### `Main`

`Main` is the reporting and CVA aggregation sheet.

#### Left block: `A:F`

- `A2:A22`: quarterly times from 0 to 5 years.
- `B2:B22`: expected short rate $\bar r(t)$, written by VBA after averaging Monte Carlo paths.
- `C2:C22`: trapezoidal integral $I(t)$, written by VBA.
- `D2:D22`: $M(t)=e^{I(t)}$, labelled `MMA`.
- `E:F`: survival values imported from `TERM structures!AA:AB`.

#### Credit inputs: `H:I`

- Recovery assumptions: `I1` for A and `I2=1-I1` for B.
- CDS spreads: `I3=70/10000` and `I4=50/10000` in the current file.
- `I6` and `I7`: total CVA(A) and CVA(B), obtained by summing rows 9 and 10.

#### Exposure/CVA block: `K:AF`

- Row 1: dates $0,0.25,\ldots,5$.
- Rows 2–3: maximum positive and negative-direction exposures.
- Rows 4–5: expected positive exposures for A and B.
- Rows 6–7: CDS/recovery-implied $Q_A(t)$ and $Q_B(t)$.
- Row 8: transposed money-market-account values.
- Rows 9–10: period CVA contributions.
- Rows 11–13: 90%, 95%, and 99% positive-exposure quantiles for A.

The sheet also contains charts comparing exposure profiles and selected simulated swap paths.

### `TERM structures`

This sheet calibrates the initial market term structure and constructs credit/survival curves.

#### Market and NSS section: `A:H`

- `A:B`: tenor labels and time to maturity.
- `C`: market `par yield AAA` observations.
- `D`: a second par-yield series.
- `E=D-C`: credit-spread difference.
- `F`: NSS fitted yield from `NSS_Spot(...)/100`.
- `G`: spot-rate series used to build discount factors.
- `H=(C-F)^2`: pointwise NSS fit error.

#### Vasicek calibration section: `I:O`

- `I`: Vasicek ZCB prices.
- `J=(I-W)^2`: pointwise ZCB pricing error.
- `M2:M5`: cells labelled `r0`, `alpha`, `mu`, and `sigma`.
- `M6=SUM(J3:J35)`: Vasicek SSE.
- `O2:O7`: NSS parameters $\beta_0,\ldots,\beta_3,\tau_1,\tau_2$.
- `O8=SUM(H3:H35)`: NSS SSE.

The chart on this sheet compares the market discount factors in `W3:W35` with Vasicek values in `I3:I35`.

#### Risky term-structure section: `Q:AB`

Let $P_i=e^{-z_iT_i}$ be the risk-free discount factor (`W`), $\delta_i$ the accrual factor (`U`), and $X_i$ the risky par yield (`X`). The workbook accumulates

$$
V_i=V_{i-1}+\delta_iP_i,
$$

$$
Y_i=Y_{i-1}+\delta_iP_iQ_i^A,
$$

$$
Z_i=Z_{i-1}+\left(Q_{i-1}^A-Q_i^A\right)P_i.
$$

For later rows, the survival recursions are implemented as

$$
Q_i^A=
\frac{1-X_iY_{i-1}-R_A\left(Z_{i-1}+Q_{i-1}^A P_i\right)}
{P_iR_B},
$$

$$
Q_i^B=
\frac{1-X_iY_{i-1}-R_B\left(Z_{i-1}+Q_{i-1}^A P_i\right)}
{P_iR_A}.
$$

`AA:AB` are imported into `Main!E:F`. The first maturity uses separate rounded initialization formulas. Because the B recursion also uses A's intermediate annuity/default terms, this section should be independently validated.

### `IRS`

This sheet prices the swap at inception.

#### Floating leg: `A:F`

- `A3:A23`: quarterly maturities from 0 to 5 years.
- `B3:B23`: Vasicek ZCB values.
- `C4:C23`: periodic forward returns, $P_{i-1}/P_i-1$.
- `F3=1-B23`: floating-leg value by the ZCB identity.
- `F4=SUMPRODUCT(B4:B23,C4:C23)`: floating-leg value from discounted forward payments.
- `F5=F3-N24`: inception residual swap value.

#### Fixed leg: `I:P`

- `I:J`: quarterly maturities and Vasicek discount factors.
- `K`: quarterly accrual factors, $0.25$.
- `L`: cumulative annuity.
- `M`: fixed rate repeated over payment dates.
- `N`: discounted fixed coupon amounts.
- `P4=F4/L23`: par fixed rate.
- `N24=SUM(N4:N23)`: fixed-leg present value.

The Vasicek inputs appear in `S3:S6`.

### `MonteCarlo`

This sheet stores the simulation controls and short-rate paths.

- `B1:B9`: swap tenor, frequencies, initial swap values, scenario count, time step, shock scale, and time-step count.
- `G2:G5`: `r0`, `alpha`, `mu`, and `sigma`.
- `B14:V14`: the 21-point grid from 0 to 5 years.
- `A15:A10014`: scenario identifiers.
- `B15:V10014`: 10,000 Vasicek short-rate paths written by VBA.

`MC_sim` also writes the averaged short-rate path, trapezoidal integral, and money-market account into `Main!B:D`.

### `Term_Structures`

This sheet contains pathwise discount factors and a small deterministic summary.

- `A:D`: expected-rate, integral, and money-market-account summary.
- `F2:F6`: Vasicek inputs and scenario count.
- `P2:AI2`: maturities from 0.25 to 5 years.
- `O3:O10002`: scenario identifiers.
- `P3:AI10002`: 10,000 by 20 matrix of simulated Vasicek discount factors.

For scenario $n$ and time column $j$, the combined macro writes

$$
P_{n,j}=\mathrm{Vasicek\_ZCB}
\left(t_j,r_{n,j},\alpha_{\mathrm{sheet}},\mu_{\mathrm{sheet}},\sigma\right).
$$

Because the UDF expects `(tau, r, mu, alpha, sigma)`, this call maps the sheet's `alpha` into the function's `mu` argument and the sheet's `mu` into the function's `alpha` argument.

Cells `AL2:AS2` extend the displayed maturity headings to seven years, but the principal simulation matrix stops at `AI` (five years).

### `Annuity`

`Annuity` stores one cumulative fixed-leg annuity path per scenario.

- `B4:V4`: dates from 0 to 5 years.
- `A5:A10004`: scenario identifiers.
- `B5:B10004`: zero initial annuity.
- `C5:V10004`: cumulative quarterly annuity values.

The intended recurrence is $A_{n,j}=A_{n,j-1}+0.25P_{n,j}$. These values are read in reverse order by `PV_Swap` to represent the remaining fixed-leg annuity.

### `PV_Swap`

`PV_Swap` contains the simulated mark-to-market matrix.

- `B4:V4`: dates from 0 to 5 years.
- `A5:A10004`: scenario identifiers.
- `B5:V10004`: scenario swap values.

The VBA sets the first and final columns to the inception residual `IRS!F5`. For the 19 intermediate dates it applies

$$
V=(1-P_{\mathrm{remaining}})-K A_{\mathrm{remaining}},
$$

using reverse maturity indices from the discount-factor and annuity matrices.

### `Future Exposure`

This sheet splits each signed swap value into two non-negative exposure matrices.

- `B:V`: `FE(A)` from $\max(V,0)$.
- `Y:AS`: `FE(B)` from $\max(-V,0)$.
- Rows `5:10004`: 10,000 scenarios.
- Row 4: dates from 0 to 5 years.

`Main` uses these matrices for averages, maxima, and percentiles.

## VBA procedures and recalculation order

The VBA project contains two UDF modules and six simulation modules:

| Procedure | Role |
|---|---|
| `NSS_Spot` | NSS fitted yield |
| `Vasicek_ZCB`, `Vasicek_A`, `Vasicek_B` | Vasicek zero-coupon price |
| `MC_sim` | Short-rate simulation plus `Main!B:D` summary |
| `Term_structure_sim` | Scenario discount factors |
| `Annuity_sim` | Scenario annuities |
| `PV_Swap_sim` | Scenario swap values |
| `FE_sim` | Positive/negative exposure split |
| `CVA_sim` | Attempted end-to-end orchestration |

The intended execution order is:

```text
MC_sim
-> Term_structure_sim
-> Annuity_sim
-> PV_Swap_sim
-> FE_sim
-> Excel recalculation of Main
```

Do not assume that running `CVA_sim` currently reproduces this sequence correctly; see the next section.

## Known implementation and model-risk issues

These findings are documented so that stored workbook results are not mistaken for a fully reproducible production calculation.

1. **Macro-disabled recalculation:** 106 worksheet formulas call VBA UDFs: `TERM structures!F3:F35`, `TERM structures!I3:I35`, `IRS!B4:B23`, and `IRS!J4:J23`. Without VBA, these cells may retain cached values or return `#NAME?` after recalculation.
2. **Stored scenario outputs:** most of the 10,000-row matrices are VBA-written constants, not live formulas. Parameter changes do not propagate until the appropriate macros run successfully.
3. **Missing `IRS_CVA` sheet:** `Term_structure_sim`, `Annuity_sim`, and `PV_Swap_sim` reference a worksheet named `IRS_CVA`, which does not exist. The likely intended sheet is `Term_Structures`.
4. **`TimeStep` versus `TimeSteps`:** `CVA_sim` declares `TimeSteps` but later loops on undeclared `TimeStep`. Under the current module settings this prevents the annuity, PV, and exposure loops from executing as intended.
5. **Vasicek argument mapping:** the UDF signature is `(tau, r, mu, alpha, sigma)`, but worksheet labels and macro calls pass `alpha` before `mu`. Cached ZCB outputs reflect this semantic reversal.
6. **Monte Carlo variance:** `MonteCarlo!B8` multiplies the variance term by $\alpha$; the standard exact Vasicek transition divides by $\alpha$.
7. **Calculation state:** standalone `MC_sim` disables calculation and screen updating but does not restore them before exit. This can leave Excel in a non-recalculating state.
8. **PFE naming:** rows labelled `PFE(A)` and `PFE(B)` use `MAX`, not a chosen confidence quantile. The percentile rows are the actual 90%, 95%, and 99% quantiles.
9. **Credit mapping:** the `Main` survival formula and the extra $Q(t_j)$ factor in period CVA differ from the most common reduced-form discrete CVA formulation. Validate the intended recovery/default convention.
10. **Risky-curve recursion:** the B survival recursion uses A's intermediate risky-annuity/default terms and should be independently checked.
11. **No production CCR features:** collateral, netting, margin period of risk, funding, wrong-way risk, stochastic credit spreads, and close-out conventions are outside the current scope.

## Validation checklist

Before using a refreshed result:

1. Confirm the semantic mapping of `alpha` and `mu` in every Vasicek UDF call.
2. Decide whether the Monte Carlo variance should use $\alpha$ or $1/\alpha$.
3. Replace or correct all `IRS_CVA` worksheet references.
4. Correct the `TimeStep`/`TimeSteps` variable and add `Option Explicit` to every macro module.
5. Restore `Application.ScreenUpdating` and worksheet calculation state in a guaranteed cleanup block.
6. Re-run all simulation stages in dependency order.
7. Confirm $V_{\mathrm{swap}}(0)\approx0$ and $V_{\mathrm{swap}}(T)\approx0$.
8. Compare simulated means and variances with analytical Vasicek moments.
9. Reconcile selected pathwise ZCB, annuity, PV, and exposure cells independently.
10. Benchmark CVA against a standard $\mathrm{LGD}\times\Delta\mathrm{PD}\times\mathrm{DF}\times\mathrm{EE}$ implementation.
11. Record the random seed or introduce reproducible random-number generation for controlled tests.

## Disclaimer

This repository is for educational, research, and portfolio-demonstration purposes. It is not investment advice and is not validated for trading, accounting, regulatory capital, or financial reporting.
