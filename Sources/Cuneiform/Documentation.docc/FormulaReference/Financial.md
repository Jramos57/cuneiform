# Financial Functions

Excel-compatible financial functions for loans, investments, depreciation, and securities analysis.

## Overview

Cuneiform provides comprehensive financial functions for calculating loan payments, analyzing investments, computing depreciation schedules, and evaluating securities. These functions are essential for financial modeling, accounting, and investment analysis.

All financial functions are compatible with Microsoft Excel's implementations and use standard financial formulas for present value, future value, internal rate of return, and other calculations.

## Quick Reference

### Loans and Annuities
- ``PMT`` - Calculate periodic payment for a loan
- ``IPMT`` - Calculate interest portion of a payment
- ``PPMT`` - Calculate principal portion of a payment
- ``NPER`` - Calculate number of payment periods
- ``RATE`` - Calculate interest rate per period
- ``PV`` - Calculate present value
- ``FV`` - Calculate future value
- ``CUMIPMT`` - Calculate cumulative interest paid
- ``CUMPRINC`` - Calculate cumulative principal paid

### Investment Analysis
- ``NPV`` - Calculate net present value
- ``XNPV`` - Calculate NPV for irregular cash flows
- ``IRR`` - Calculate internal rate of return
- ``XIRR`` - Calculate IRR for irregular cash flows
- ``MIRR`` - Calculate modified internal rate of return
- ``PDURATION`` - Calculate periods to reach investment goal
- ``RRI`` - Calculate equivalent interest rate

### Depreciation
- ``SLN`` - Straight-line depreciation
- ``DB`` - Declining balance depreciation
- ``DDB`` - Double declining balance depreciation
- ``SYD`` - Sum-of-years digits depreciation
- ``VDB`` - Variable declining balance depreciation

### Securities
- ``PRICE`` - Calculate bond price
- ``YIELD`` - Calculate bond yield
- ``PRICEDISC`` - Calculate price of discounted security
- ``YIELDDISC`` - Calculate yield of discounted security
- ``PRICEMAT`` - Calculate price of security paying interest at maturity
- ``YIELDMAT`` - Calculate yield of security paying interest at maturity
- ``ACCRINT`` - Calculate accrued interest
- ``ACCRINTM`` - Calculate accrued interest at maturity

### Treasury Bills
- ``TBILLEQ`` - Calculate bond-equivalent yield
- ``TBILLPRICE`` - Calculate T-bill price
- ``TBILLYIELD`` - Calculate T-bill yield

### Interest Rate Conversions
- ``EFFECT`` - Calculate effective annual interest rate
- ``NOMINAL`` - Calculate nominal annual interest rate
- ``DOLLARFR`` - Convert decimal to fractional dollar
- ``DOLLARDE`` - Convert fractional dollar to decimal

### Bond Analysis
- ``DURATION`` - Calculate Macaulay duration
- ``MDURATION`` - Calculate modified duration
- ``DISC`` - Calculate discount rate
- ``INTRATE`` - Calculate interest rate for fully invested security
- ``RECEIVED`` - Calculate amount received at maturity
- ``COUPDAYBS`` - Days from coupon period start to settlement
- ``COUPDAYS`` - Days in coupon period
- ``COUPDAYSNC`` - Days from settlement to next coupon
- ``COUPNCD`` - Next coupon date
- ``COUPPCD`` - Previous coupon date
- ``COUPNUM`` - Number of coupons between settlement and maturity

---

## Loans and Annuities

### PMT

Calculates the periodic payment for a loan based on constant payments and a constant interest rate.

**Syntax:** `PMT(rate, nper, pv, [fv], [type])`

**Parameters:**
- `rate`: Interest rate per period
- `nper`: Total number of payment periods
- `pv`: Present value (loan amount)
- `fv` *(optional)*: Future value or cash balance after last payment (default 0)
- `type` *(optional)*: Payment timing - 0 for end of period (default), 1 for beginning

**Returns:** Double - The periodic payment amount (typically negative)

**Examples:**
```swift
// Monthly payment on $200,000 loan at 5% annual rate for 30 years
let payment = evaluator.evaluate("=PMT(0.05/12, 30*12, 200000)")  // -1073.64

// Payment at beginning of period
let payment2 = evaluator.evaluate("=PMT(0.06/12, 10*12, 100000, 0, 1)")  // -1101.09
```

**Excel Documentation:** [PMT function](https://support.microsoft.com/en-us/office/pmt-function-0214da64-9a63-4996-bc20-214433fa6441)

**Implementation Status:** ✅ Full implementation

---

### IPMT

Calculates the interest portion of a payment for a specific period of a loan.

**Syntax:** `IPMT(rate, per, nper, pv, [fv], [type])`

**Parameters:**
- `rate`: Interest rate per period
- `per`: Period for which to calculate interest (1 to nper)
- `nper`: Total number of payment periods
- `pv`: Present value (loan amount)
- `fv` *(optional)*: Future value (default 0)
- `type` *(optional)*: Payment timing - 0 for end (default), 1 for beginning

**Returns:** Double - Interest payment for the specified period

**Examples:**
```swift
// Interest portion of 12th payment on $200,000 loan at 5% for 30 years
let interest = evaluator.evaluate("=IPMT(0.05/12, 12, 30*12, 200000)")  // -832.40

// Interest on first payment
let firstInterest = evaluator.evaluate("=IPMT(0.05/12, 1, 30*12, 200000)")  // -833.33
```

**Excel Documentation:** [IPMT function](https://support.microsoft.com/en-us/office/ipmt-function-5cce0ad6-8402-4a41-8d29-61a0b054cb6f)

**Implementation Status:** ✅ Full implementation

---

### PPMT

Calculates the principal portion of a payment for a specific period of a loan.

**Syntax:** `PPMT(rate, per, nper, pv, [fv], [type])`

**Parameters:**
- `rate`: Interest rate per period
- `per`: Period for which to calculate principal (1 to nper)
- `nper`: Total number of payment periods
- `pv`: Present value (loan amount)
- `fv` *(optional)*: Future value (default 0)
- `type` *(optional)*: Payment timing - 0 for end (default), 1 for beginning

**Returns:** Double - Principal payment for the specified period

**Examples:**
```swift
// Principal portion of 12th payment on $200,000 loan at 5% for 30 years
let principal = evaluator.evaluate("=PPMT(0.05/12, 12, 30*12, 200000)")  // -241.24

// Principal on first payment
let firstPrincipal = evaluator.evaluate("=PPMT(0.05/12, 1, 30*12, 200000)")  // -240.31
```

**Excel Documentation:** [PPMT function](https://support.microsoft.com/en-us/office/ppmt-function-c370d9e3-7749-4ca4-beea-b06c6ac95e1b)

**Implementation Status:** ✅ Full implementation

---

### NPER

Calculates the number of periods for an investment or loan based on periodic, constant payments and a constant interest rate.

**Syntax:** `NPER(rate, pmt, pv, [fv], [type])`

**Parameters:**
- `rate`: Interest rate per period
- `pmt`: Payment made each period (constant)
- `pv`: Present value (loan amount or investment)
- `fv` *(optional)*: Future value (default 0)
- `type` *(optional)*: Payment timing - 0 for end (default), 1 for beginning

**Returns:** Double - Number of periods

**Examples:**
```swift
// Periods to pay off $100,000 loan with $1,000 monthly payments at 6% annual
let periods = evaluator.evaluate("=NPER(0.06/12, -1000, 100000)")  // 139.58

// Years to reach $1M with $500/month at 8% return
let years = evaluator.evaluate("=NPER(0.08/12, -500, 0, 1000000)/12")  // 41.04
```

**Excel Documentation:** [NPER function](https://support.microsoft.com/en-us/office/nper-function-240535b5-6653-4d2d-bfcf-b6a38151d815)

**Implementation Status:** ✅ Full implementation

---

### RATE

Calculates the interest rate per period of an annuity or loan.

**Syntax:** `RATE(nper, pmt, pv, [fv], [type], [guess])`

**Parameters:**
- `nper`: Total number of payment periods
- `pmt`: Payment made each period
- `pv`: Present value (loan amount or investment)
- `fv` *(optional)*: Future value (default 0)
- `type` *(optional)*: Payment timing - 0 for end (default), 1 for beginning
- `guess` *(optional)*: Initial guess for rate (default 0.1)

**Returns:** Double - Interest rate per period

**Examples:**
```swift
// Interest rate for $200,000 loan with $1,073.64 monthly payment for 30 years
let rate = evaluator.evaluate("=RATE(360, -1073.64, 200000)")  // 0.00417 (0.417% monthly)

// Annual rate
let annualRate = evaluator.evaluate("=RATE(360, -1073.64, 200000)*12")  // 0.05 (5%)
```

**Excel Documentation:** [RATE function](https://support.microsoft.com/en-us/office/rate-function-9f665657-4a7e-4bb7-a030-83fc59e748ce)

**Implementation Status:** ✅ Full implementation (uses Newton-Raphson iteration)

---

### PV

Calculates the present value of an investment based on periodic, constant payments and a constant interest rate.

**Syntax:** `PV(rate, nper, pmt, [fv], [type])`

**Parameters:**
- `rate`: Interest rate per period
- `nper`: Total number of payment periods
- `pmt`: Payment made each period
- `fv` *(optional)*: Future value (default 0)
- `type` *(optional)*: Payment timing - 0 for end (default), 1 for beginning

**Returns:** Double - Present value

**Examples:**
```swift
// Present value of annuity paying $1,000/month for 10 years at 6%
let pv = evaluator.evaluate("=PV(0.06/12, 120, 1000)")  // -90,073.45

// Loan amount for given payment
let loanAmount = evaluator.evaluate("=PV(0.05/12, 360, -1073.64)")  // 200,000
```

**Excel Documentation:** [PV function](https://support.microsoft.com/en-us/office/pv-function-23879d31-0e02-4321-be01-da16e8168cbd)

**Implementation Status:** ✅ Full implementation

---

### FV

Calculates the future value of an investment based on periodic, constant payments and a constant interest rate.

**Syntax:** `FV(rate, nper, pmt, [pv], [type])`

**Parameters:**
- `rate`: Interest rate per period
- `nper`: Total number of payment periods
- `pmt`: Payment made each period
- `pv` *(optional)*: Present value (default 0)
- `type` *(optional)*: Payment timing - 0 for end (default), 1 for beginning

**Returns:** Double - Future value

**Examples:**
```swift
// Future value of $500/month investment for 10 years at 8% annual return
let fv = evaluator.evaluate("=FV(0.08/12, 120, -500)")  // 91,473.44

// Future value including initial investment
let fvWithPV = evaluator.evaluate("=FV(0.08/12, 120, -500, -10000)")  // 113,506.59
```

**Excel Documentation:** [FV function](https://support.microsoft.com/en-us/office/fv-function-2eef9f44-a084-4c61-bdd8-4fe4bb1b71b3)

**Implementation Status:** ✅ Full implementation

---

### CUMIPMT

Calculates cumulative interest paid on a loan between two periods.

**Syntax:** `CUMIPMT(rate, nper, pv, start_period, end_period, type)`

**Parameters:**
- `rate`: Interest rate per period
- `nper`: Total number of payment periods
- `pv`: Present value (loan amount)
- `start_period`: First period in calculation (1-based)
- `end_period`: Last period in calculation
- `type`: Payment timing - 0 for end of period, 1 for beginning

**Returns:** Double - Cumulative interest paid

**Examples:**
```swift
// Total interest paid in first year of $200,000 loan at 5% for 30 years
let interest = evaluator.evaluate("=CUMIPMT(0.05/12, 360, 200000, 1, 12, 0)")  // -9,916.77

// Interest in years 2-5
let interestYears2to5 = evaluator.evaluate("=CUMIPMT(0.05/12, 360, 200000, 13, 60, 0)")
```

**Excel Documentation:** [CUMIPMT function](https://support.microsoft.com/en-us/office/cumipmt-function-61067bb0-9016-427d-b95b-1a752af0e606)

**Implementation Status:** ✅ Full implementation

---

### CUMPRINC

Calculates cumulative principal paid on a loan between two periods.

**Syntax:** `CUMPRINC(rate, nper, pv, start_period, end_period, type)`

**Parameters:**
- `rate`: Interest rate per period
- `nper`: Total number of payment periods
- `pv`: Present value (loan amount)
- `start_period`: First period in calculation (1-based)
- `end_period`: Last period in calculation
- `type`: Payment timing - 0 for end of period, 1 for beginning

**Returns:** Double - Cumulative principal paid

**Examples:**
```swift
// Total principal paid in first year of $200,000 loan at 5% for 30 years
let principal = evaluator.evaluate("=CUMPRINC(0.05/12, 360, 200000, 1, 12, 0)")  // -2,966.91

// Principal reduction over 5 years
let principal5yr = evaluator.evaluate("=CUMPRINC(0.05/12, 360, 200000, 1, 60, 0)")
```

**Excel Documentation:** [CUMPRINC function](https://support.microsoft.com/en-us/office/cumprinc-function-94a4516d-bd65-41a1-bc16-053a6af4c04d)

**Implementation Status:** ✅ Full implementation

---

## Investment Analysis

### NPV

Calculates the net present value of an investment based on a discount rate and a series of future periodic cash flows.

**Syntax:** `NPV(rate, value1, [value2], ...)`

**Parameters:**
- `rate`: Discount rate per period
- `value1, value2, ...`: 1 to 254 arguments representing periodic cash flows

**Returns:** Double - Net present value

**Examples:**
```swift
// NPV of investment with initial cost in cell (not included in NPV)
let npv = evaluator.evaluate("=NPV(0.10, -50000, 10000, 15000, 20000, 25000)")  // 11,826.59

// Full NPV including initial investment
let fullNPV = evaluator.evaluate("=-100000 + NPV(0.08, 20000, 30000, 40000, 50000)")
```

**Excel Documentation:** [NPV function](https://support.microsoft.com/en-us/office/npv-function-8672cb67-2576-4d07-b67b-ac28acf2a568)

**Implementation Status:** ✅ Full implementation

---

### XNPV

Calculates the net present value for irregular cash flows with specific dates.

**Syntax:** `XNPV(rate, values, dates)`

**Parameters:**
- `rate`: Annual discount rate
- `values`: Series of cash flows (array or range)
- `dates`: Corresponding dates for each cash flow (as Excel serial dates)

**Returns:** Double - Net present value for irregular cash flows

**Examples:**
```swift
// NPV with specific dates for each cash flow
let xnpv = evaluator.evaluate("=XNPV(0.09, {-100000,20000,40000,25000}, {44562,44653,44744,44927})")

// Using date ranges
let xnpvRange = evaluator.evaluate("=XNPV(0.09, A1:A4, B1:B4)")
```

**Excel Documentation:** [XNPV function](https://support.microsoft.com/en-us/office/xnpv-function-1b42bbf6-370f-4532-a0eb-d67c16b664b7)

**Implementation Status:** ✅ Full implementation

---

### IRR

Calculates the internal rate of return for a series of periodic cash flows.

**Syntax:** `IRR(values, [guess])`

**Parameters:**
- `values`: Array or range containing cash flows (must include at least one positive and one negative)
- `guess` *(optional)*: Initial guess for rate (default 0.1)

**Returns:** Double - Internal rate of return

**Examples:**
```swift
// IRR for investment with initial cost and returns
let irr = evaluator.evaluate("=IRR({-100000, 20000, 30000, 40000, 50000})")  // 0.189 (18.9%)

// Using cell range
let irrRange = evaluator.evaluate("=IRR(A1:A5)")
```

**Excel Documentation:** [IRR function](https://support.microsoft.com/en-us/office/irr-function-64925eaa-9988-495b-b290-3ad0c163c1bc)

**Implementation Status:** ✅ Full implementation (uses Newton-Raphson iteration)

---

### XIRR

Calculates the internal rate of return for irregular cash flows with specific dates.

**Syntax:** `XIRR(values, dates, [guess])`

**Parameters:**
- `values`: Series of cash flows (array or range)
- `dates`: Corresponding dates for each cash flow (as Excel serial dates)
- `guess` *(optional)*: Initial guess for rate (default 0.1)

**Returns:** Double - Internal rate of return for irregular cash flows

**Examples:**
```swift
// IRR with specific dates
let xirr = evaluator.evaluate("=XIRR({-100000,20000,40000,60000}, {44562,44653,44835,45017})")

// Using date ranges
let xirrRange = evaluator.evaluate("=XIRR(A1:A4, B1:B4, 0.1)")
```

**Excel Documentation:** [XIRR function](https://support.microsoft.com/en-us/office/xirr-function-de1242ec-6477-445b-b11b-a303ad9adc9d)

**Implementation Status:** ✅ Full implementation (uses Newton-Raphson iteration)

---

### MIRR

Calculates the modified internal rate of return for a series of periodic cash flows.

**Syntax:** `MIRR(values, finance_rate, reinvest_rate)`

**Parameters:**
- `values`: Array or range containing cash flows
- `finance_rate`: Interest rate paid on money borrowed (cost of financing)
- `reinvest_rate`: Interest rate received on reinvestment of positive cash flows

**Returns:** Double - Modified internal rate of return

**Examples:**
```swift
// MIRR with different finance and reinvestment rates
let mirr = evaluator.evaluate("=MIRR({-100000, 20000, 30000, 40000, 50000}, 0.10, 0.12)")  // 0.138

// Conservative assumption: same rate for both
let mirrSame = evaluator.evaluate("=MIRR(A1:A5, 0.08, 0.08)")
```

**Excel Documentation:** [MIRR function](https://support.microsoft.com/en-us/office/mirr-function-b020f038-7492-4fb4-93c1-35c345b53524)

**Implementation Status:** ✅ Full implementation

---

### PDURATION

Calculates the number of periods required for an investment to reach a specified value.

**Syntax:** `PDURATION(rate, pv, fv)`

**Parameters:**
- `rate`: Interest rate per period
- `pv`: Present value (initial investment)
- `fv`: Future value (target amount)

**Returns:** Double - Number of periods

**Examples:**
```swift
// Periods to grow $10,000 to $20,000 at 8% annual return
let periods = evaluator.evaluate("=PDURATION(0.08, 10000, 20000)")  // 9.01 years

// Doubling time
let doublingTime = evaluator.evaluate("=PDURATION(0.07, 1, 2)")  // 10.24 years
```

**Excel Documentation:** [PDURATION function](https://support.microsoft.com/en-us/office/pduration-function-44f33460-5be5-4c90-b857-22308892adaf)

**Implementation Status:** ✅ Full implementation

---

### RRI

Calculates the equivalent interest rate for the growth of an investment.

**Syntax:** `RRI(nper, pv, fv)`

**Parameters:**
- `nper`: Number of periods
- `pv`: Present value (initial investment)
- `fv`: Future value (final amount)

**Returns:** Double - Equivalent interest rate per period

**Examples:**
```swift
// Rate needed for $10,000 to grow to $20,000 in 10 years
let rate = evaluator.evaluate("=RRI(10, 10000, 20000)")  // 0.0718 (7.18%)

// Annual growth rate
let growthRate = evaluator.evaluate("=RRI(5, 100000, 150000)")  // 0.0845 (8.45%)
```

**Excel Documentation:** [RRI function](https://support.microsoft.com/en-us/office/rri-function-6f5822d8-7ef1-4233-944c-79e8172930f4)

**Implementation Status:** ✅ Full implementation

---

## Depreciation

### SLN

Calculates straight-line depreciation for one period.

**Syntax:** `SLN(cost, salvage, life)`

**Parameters:**
- `cost`: Initial cost of the asset
- `salvage`: Value at end of depreciation (salvage value)
- `life`: Number of periods over which asset is depreciated

**Returns:** Double - Depreciation per period

**Examples:**
```swift
// Annual depreciation of $30,000 asset with $5,000 salvage over 10 years
let depreciation = evaluator.evaluate("=SLN(30000, 5000, 10)")  // 2,500

// Equipment depreciation
let equipmentDepr = evaluator.evaluate("=SLN(100000, 10000, 5)")  // 18,000
```

**Excel Documentation:** [SLN function](https://support.microsoft.com/en-us/office/sln-function-cdb666e5-c1c6-40a7-806a-e695edc2f1c8)

**Implementation Status:** ✅ Full implementation

---

### DB

Calculates depreciation using the fixed-declining balance method.

**Syntax:** `DB(cost, salvage, life, period, [month])`

**Parameters:**
- `cost`: Initial cost of the asset
- `salvage`: Value at end of depreciation
- `life`: Number of periods over which asset is depreciated
- `period`: Period for which to calculate depreciation
- `month` *(optional)*: Number of months in first year (default 12)

**Returns:** Double - Depreciation for specified period

**Examples:**
```swift
// First year depreciation of $50,000 asset with $10,000 salvage over 5 years
let depreciation = evaluator.evaluate("=DB(50000, 10000, 5, 1)")  // 13,388

// Partial first year (7 months)
let partialYear = evaluator.evaluate("=DB(50000, 10000, 5, 1, 7)")  // 7,810
```

**Excel Documentation:** [DB function](https://support.microsoft.com/en-us/office/db-function-354e7d28-5f93-4ff1-8a52-eb4ee549d9d7)

**Implementation Status:** ✅ Full implementation

---

### DDB

Calculates depreciation using the double-declining balance method or another specified method.

**Syntax:** `DDB(cost, salvage, life, period, [factor])`

**Parameters:**
- `cost`: Initial cost of the asset
- `salvage`: Value at end of depreciation
- `life`: Number of periods over which asset is depreciated
- `period`: Period for which to calculate depreciation
- `factor` *(optional)*: Rate at which balance declines (default 2)

**Returns:** Double - Depreciation for specified period

**Examples:**
```swift
// Double-declining depreciation for period 1
let depreciation = evaluator.evaluate("=DDB(50000, 10000, 5, 1)")  // 20,000

// 150% declining balance
let depreciation150 = evaluator.evaluate("=DDB(50000, 10000, 5, 1, 1.5)")  // 15,000
```

**Excel Documentation:** [DDB function](https://support.microsoft.com/en-us/office/ddb-function-519a7a37-8772-4c96-85c0-ed2c209717a5)

**Implementation Status:** ✅ Full implementation

---

### SYD

Calculates depreciation using the sum-of-years' digits method.

**Syntax:** `SYD(cost, salvage, life, period)`

**Parameters:**
- `cost`: Initial cost of the asset
- `salvage`: Value at end of depreciation
- `life`: Number of periods over which asset is depreciated
- `period`: Period for which to calculate depreciation

**Returns:** Double - Depreciation for specified period

**Examples:**
```swift
// First year SYD depreciation
let depreciation = evaluator.evaluate("=SYD(50000, 10000, 5, 1)")  // 13,333

// Second year
let year2 = evaluator.evaluate("=SYD(50000, 10000, 5, 2)")  // 10,667
```

**Excel Documentation:** [SYD function](https://support.microsoft.com/en-us/office/syd-function-069f8106-b60b-4ca2-98e0-2a0f206bdb27)

**Implementation Status:** ✅ Full implementation

---

### VDB

Calculates depreciation using the variable declining balance method.

**Syntax:** `VDB(cost, salvage, life, start_period, end_period, [factor], [no_switch])`

**Parameters:**
- `cost`: Initial cost of the asset
- `salvage`: Value at end of depreciation
- `life`: Number of periods over which asset is depreciated
- `start_period`: Starting period for calculation
- `end_period`: Ending period for calculation
- `factor` *(optional)*: Rate at which balance declines (default 2)
- `no_switch` *(optional)*: Switch to straight-line (default FALSE)

**Returns:** Double - Depreciation for period range

**Examples:**
```swift
// Depreciation from period 0 to 1
let depreciation = evaluator.evaluate("=VDB(50000, 10000, 5, 0, 1)")  // 20,000

// Depreciation from period 2 to 3
let period2to3 = evaluator.evaluate("=VDB(50000, 10000, 5, 2, 3)")  // 9,600
```

**Excel Documentation:** [VDB function](https://support.microsoft.com/en-us/office/vdb-function-dde4e207-f3fa-488d-91d2-66d55e861d73)

**Implementation Status:** ✅ Full implementation (simplified)

---

## Securities

### PRICE

Calculates the price per $100 face value of a security that pays periodic interest.

**Syntax:** `PRICE(settlement, maturity, rate, yld, redemption, frequency, [basis])`

**Parameters:**
- `settlement`: Security's settlement date (as Excel serial date)
- `maturity`: Security's maturity date
- `rate`: Annual coupon rate
- `yld`: Annual yield
- `redemption`: Redemption value per $100 face value
- `frequency`: Number of coupon payments per year (1, 2, or 4)
- `basis` *(optional)*: Day count basis (default 0)

**Returns:** Double - Security price per $100 face value

**Examples:**
```swift
// Price of bond with 6% coupon, 8% yield, semi-annual payments
let price = evaluator.evaluate("=PRICE(DATE(2024,2,15), DATE(2029,2,15), 0.06, 0.08, 100, 2)")

// Using cell references
let priceFromCells = evaluator.evaluate("=PRICE(A1, B1, 0.05, 0.06, 100, 2)")
```

**Excel Documentation:** [PRICE function](https://support.microsoft.com/en-us/office/price-function-3ea9deac-8dfa-436f-a7c8-17ea02c21b0a)

**Implementation Status:** ✅ Full implementation (simplified bond pricing)

---

### YIELD

Calculates the yield of a security that pays periodic interest.

**Syntax:** `YIELD(settlement, maturity, rate, pr, redemption, frequency, [basis])`

**Parameters:**
- `settlement`: Security's settlement date (as Excel serial date)
- `maturity`: Security's maturity date
- `rate`: Annual coupon rate
- `pr`: Security's price per $100 face value
- `redemption`: Redemption value per $100 face value
- `frequency`: Number of coupon payments per year (1, 2, or 4)
- `basis` *(optional)*: Day count basis (default 0)

**Returns:** Double - Security yield

**Examples:**
```swift
// Yield of bond priced at $95 with 6% coupon, semi-annual payments
let yield = evaluator.evaluate("=YIELD(DATE(2024,2,15), DATE(2029,2,15), 0.06, 95, 100, 2)")

// Yield calculation from cells
let yieldFromCells = evaluator.evaluate("=YIELD(A1, B1, 0.05, 98, 100, 2)")
```

**Excel Documentation:** [YIELD function](https://support.microsoft.com/en-us/office/yield-function-f5f5ca43-c4bd-434f-8bd2-ed3c9727a4fe)

**Implementation Status:** ✅ Full implementation (simplified yield approximation)

---

### PRICEDISC

Calculates the price per $100 face value of a discounted security.

**Syntax:** `PRICEDISC(settlement, maturity, discount, redemption, basis)`

**Parameters:**
- `settlement`: Security's settlement date (as Excel serial date)
- `maturity`: Security's maturity date
- `discount`: Security's discount rate
- `redemption`: Redemption value per $100 face value
- `basis`: Day count basis (0-4)

**Returns:** Double - Price per $100 face value

**Examples:**
```swift
// Price of discounted security with 5% discount rate
let price = evaluator.evaluate("=PRICEDISC(DATE(2024,2,1), DATE(2024,8,1), 0.05, 100, 2)")

// T-bill pricing
let tbillPrice = evaluator.evaluate("=PRICEDISC(44952, 45135, 0.0525, 100, 2)")
```

**Excel Documentation:** [PRICEDISC function](https://support.microsoft.com/en-us/office/pricedisc-function-d06ad7c1-380e-4be7-9fd9-75e3079acfd3)

**Implementation Status:** ✅ Full implementation

---

### YIELDDISC

Calculates the annual yield of a discounted security.

**Syntax:** `YIELDDISC(settlement, maturity, pr, redemption, basis)`

**Parameters:**
- `settlement`: Security's settlement date (as Excel serial date)
- `maturity`: Security's maturity date
- `pr`: Security's price per $100 face value
- `redemption`: Redemption value per $100 face value
- `basis`: Day count basis (0-4)

**Returns:** Double - Annual yield of discounted security

**Examples:**
```swift
// Yield of discounted security priced at $97.50
let yield = evaluator.evaluate("=YIELDDISC(DATE(2024,2,1), DATE(2024,8,1), 97.50, 100, 2)")

// T-bill yield
let tbillYield = evaluator.evaluate("=YIELDDISC(44952, 45135, 98.45, 100, 2)")
```

**Excel Documentation:** [YIELDDISC function](https://support.microsoft.com/en-us/office/yielddisc-function-a9dbdbae-7dae-46de-afea-f020d5c95a04)

**Implementation Status:** ✅ Full implementation

---

### PRICEMAT

Calculates the price per $100 face value of a security that pays interest at maturity.

**Syntax:** `PRICEMAT(settlement, maturity, issue, rate, yld, [basis])`

**Parameters:**
- `settlement`: Security's settlement date (as Excel serial date)
- `maturity`: Security's maturity date
- `issue`: Security's issue date
- `rate`: Interest rate at issue
- `yld`: Annual yield
- `basis` *(optional)*: Day count basis (default 0)

**Returns:** Double - Price per $100 face value

**Excel Documentation:** [PRICEMAT function](https://support.microsoft.com/en-us/office/pricemat-function-52c3b4da-bc7e-476a-989f-a95f675cae77)

**Implementation Status:** 🔄 Stub (returns #CALC! error)

---

### YIELDMAT

Calculates the annual yield of a security that pays interest at maturity.

**Syntax:** `YIELDMAT(settlement, maturity, issue, rate, pr, [basis])`

**Parameters:**
- `settlement`: Security's settlement date (as Excel serial date)
- `maturity`: Security's maturity date
- `issue`: Security's issue date
- `rate`: Interest rate at issue
- `pr`: Security's price per $100 face value
- `basis` *(optional)*: Day count basis (default 0)

**Returns:** Double - Annual yield

**Excel Documentation:** [YIELDMAT function](https://support.microsoft.com/en-us/office/yieldmat-function-ba7d1809-0d33-4bcb-96c7-6c56ec62ef6f)

**Implementation Status:** 🔄 Stub (returns #CALC! error)

---

### ACCRINT

Calculates the accrued interest for a security that pays periodic interest.

**Syntax:** `ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method])`

**Parameters:**
- `issue`: Security's issue date (as Excel serial date)
- `first_interest`: First interest date
- `settlement`: Security's settlement date
- `rate`: Annual coupon rate
- `par`: Par value (face value)
- `frequency`: Number of coupon payments per year (1, 2, or 4)
- `basis` *(optional)*: Day count basis (default 0)
- `calc_method` *(optional)*: Calculation method (default TRUE)

**Returns:** Double - Accrued interest

**Examples:**
```swift
// Accrued interest on $1,000 bond with 6% coupon
let interest = evaluator.evaluate("=ACCRINT(DATE(2024,1,1), DATE(2024,7,1), DATE(2024,4,1), 0.06, 1000, 2)")

// Using cell references
let accrInt = evaluator.evaluate("=ACCRINT(A1, B1, C1, 0.05, 1000, 2)")
```

**Excel Documentation:** [ACCRINT function](https://support.microsoft.com/en-us/office/accrint-function-fe45d089-6722-4fb3-9379-e1f911d8dc74)

**Implementation Status:** ✅ Full implementation (simplified)

---

### ACCRINTM

Calculates the accrued interest for a security that pays interest at maturity.

**Syntax:** `ACCRINTM(issue, settlement, rate, par, [basis])`

**Parameters:**
- `issue`: Security's issue date (as Excel serial date)
- `settlement`: Security's maturity date
- `rate`: Annual coupon rate
- `par`: Par value (face value)
- `basis` *(optional)*: Day count basis (default 0)

**Returns:** Double - Accrued interest at maturity

**Excel Documentation:** [ACCRINTM function](https://support.microsoft.com/en-us/office/accrintm-function-f62f01f9-5754-4cc4-805b-0e70199328a7)

**Implementation Status:** 🔄 Stub (returns #CALC! error)

---

## Treasury Bills

### TBILLEQ

Calculates the bond-equivalent yield for a Treasury bill.

**Syntax:** `TBILLEQ(settlement, maturity, discount)`

**Parameters:**
- `settlement`: T-bill's settlement date (as Excel serial date)
- `maturity`: T-bill's maturity date (must be within 1 year of settlement)
- `discount`: T-bill's discount rate

**Returns:** Double - Bond-equivalent yield

**Examples:**
```swift
// Bond-equivalent yield for 90-day T-bill with 5% discount
let yield = evaluator.evaluate("=TBILLEQ(DATE(2024,1,1), DATE(2024,4,1), 0.05)")  // 0.0512

// Using days from today
let yieldNow = evaluator.evaluate("=TBILLEQ(TODAY(), TODAY()+90, 0.0525)")
```

**Excel Documentation:** [TBILLEQ function](https://support.microsoft.com/en-us/office/tbilleq-function-2ab72d90-9b4d-4efe-9fc2-0f81f2c19c8c)

**Implementation Status:** ✅ Full implementation

---

### TBILLPRICE

Calculates the price per $100 face value for a Treasury bill.

**Syntax:** `TBILLPRICE(settlement, maturity, discount)`

**Parameters:**
- `settlement`: T-bill's settlement date (as Excel serial date)
- `maturity`: T-bill's maturity date (must be within 1 year of settlement)
- `discount`: T-bill's discount rate

**Returns:** Double - Price per $100 face value

**Examples:**
```swift
// Price of 90-day T-bill with 5% discount rate
let price = evaluator.evaluate("=TBILLPRICE(DATE(2024,1,1), DATE(2024,4,1), 0.05)")  // 98.75

// Current T-bill price
let currentPrice = evaluator.evaluate("=TBILLPRICE(TODAY(), TODAY()+182, 0.0485)")
```

**Excel Documentation:** [TBILLPRICE function](https://support.microsoft.com/en-us/office/tbillprice-function-eacca992-c29d-425a-9eb8-0513fe6035a2)

**Implementation Status:** ✅ Full implementation

---

### TBILLYIELD

Calculates the yield for a Treasury bill.

**Syntax:** `TBILLYIELD(settlement, maturity, pr)`

**Parameters:**
- `settlement`: T-bill's settlement date (as Excel serial date)
- `maturity`: T-bill's maturity date (must be within 1 year of settlement)
- `pr`: T-bill's price per $100 face value

**Returns:** Double - T-bill yield

**Examples:**
```swift
// Yield of 90-day T-bill priced at $98.75
let yield = evaluator.evaluate("=TBILLYIELD(DATE(2024,1,1), DATE(2024,4,1), 98.75)")  // 0.0506

// Yield from current market price
let currentYield = evaluator.evaluate("=TBILLYIELD(TODAY(), TODAY()+182, 97.50)")
```

**Excel Documentation:** [TBILLYIELD function](https://support.microsoft.com/en-us/office/tbillyield-function-6d381232-f4b0-4cd5-8e97-45b9c03468ba)

**Implementation Status:** ✅ Full implementation

---

## Interest Rate Conversions

### EFFECT

Calculates the effective annual interest rate given the nominal rate and compounding periods.

**Syntax:** `EFFECT(nominal_rate, npery)`

**Parameters:**
- `nominal_rate`: Nominal annual interest rate
- `npery`: Number of compounding periods per year

**Returns:** Double - Effective annual interest rate

**Examples:**
```swift
// Effective rate for 6% nominal compounded monthly
let effectiveRate = evaluator.evaluate("=EFFECT(0.06, 12)")  // 0.0617 (6.17%)

// Quarterly compounding
let quarterly = evaluator.evaluate("=EFFECT(0.08, 4)")  // 0.0824 (8.24%)
```

**Excel Documentation:** [EFFECT function](https://support.microsoft.com/en-us/office/effect-function-910d4e4c-79e2-4009-95e6-507e04f11bc4)

**Implementation Status:** ✅ Full implementation

---

### NOMINAL

Calculates the nominal annual interest rate given the effective rate and compounding periods.

**Syntax:** `NOMINAL(effect_rate, npery)`

**Parameters:**
- `effect_rate`: Effective annual interest rate
- `npery`: Number of compounding periods per year

**Returns:** Double - Nominal annual interest rate

**Examples:**
```swift
// Nominal rate for 6.17% effective rate with monthly compounding
let nominalRate = evaluator.evaluate("=NOMINAL(0.0617, 12)")  // 0.06 (6%)

// From effective to nominal
let nominal = evaluator.evaluate("=NOMINAL(0.0824, 4)")  // 0.08 (8%)
```

**Excel Documentation:** [NOMINAL function](https://support.microsoft.com/en-us/office/nominal-function-7f1ae29b-6b92-435e-b950-ad8b190ddd2b)

**Implementation Status:** ✅ Full implementation

---

### DOLLARFR

Converts a decimal dollar price to a fractional dollar price.

**Syntax:** `DOLLARFR(decimal_dollar, fraction)`

**Parameters:**
- `decimal_dollar`: Decimal number to convert
- `fraction`: Integer to use as denominator of fraction

**Returns:** Double - Price expressed as fraction

**Examples:**
```swift
// Convert 1.25 to sixteenths (1.04 represents 1 4/16)
let fractional = evaluator.evaluate("=DOLLARFR(1.25, 16)")  // 1.04

// Convert to eighths
let eighths = evaluator.evaluate("=DOLLARFR(1.125, 8)")  // 1.01
```

**Excel Documentation:** [DOLLARFR function](https://support.microsoft.com/en-us/office/dollarfr-function-0835d163-3023-4a33-9824-3042c5d4f495)

**Implementation Status:** ✅ Full implementation

---

### DOLLARDE

Converts a fractional dollar price to a decimal dollar price.

**Syntax:** `DOLLARDE(fractional_dollar, fraction)`

**Parameters:**
- `fractional_dollar`: Fractional number to convert
- `fraction`: Integer denominator of fraction

**Returns:** Double - Price expressed as decimal

**Examples:**
```swift
// Convert 1.04 sixteenths to decimal (1 4/16 = 1.25)
let decimal = evaluator.evaluate("=DOLLARDE(1.04, 16)")  // 1.25

// Convert eighths to decimal
let eighthsDecimal = evaluator.evaluate("=DOLLARDE(1.01, 8)")  // 1.125
```

**Excel Documentation:** [DOLLARDE function](https://support.microsoft.com/en-us/office/dollarde-function-db85aab0-1677-428a-9dfd-a38476693427)

**Implementation Status:** ✅ Full implementation

---

## Bond Analysis

### DURATION

Calculates the Macaulay duration of a security with periodic interest payments.

**Syntax:** `DURATION(settlement, maturity, coupon, yld, frequency, [basis])`

**Parameters:**
- `settlement`: Security's settlement date (as Excel serial date)
- `maturity`: Security's maturity date
- `coupon`: Annual coupon rate
- `yld`: Annual yield
- `frequency`: Number of coupon payments per year (1, 2, or 4)
- `basis` *(optional)*: Day count basis (default 0)

**Returns:** Double - Macaulay duration

**Excel Documentation:** [DURATION function](https://support.microsoft.com/en-us/office/duration-function-b254ea57-eadc-4c32-8b27-87c56d9b3549)

**Implementation Status:** 🔄 Stub (returns #CALC! error - requires complex bond duration calculation)

---

### MDURATION

Calculates the modified Macaulay duration for a security with an assumed par value of $100.

**Syntax:** `MDURATION(settlement, maturity, coupon, yld, frequency, [basis])`

**Parameters:**
- `settlement`: Security's settlement date (as Excel serial date)
- `maturity`: Security's maturity date
- `coupon`: Annual coupon rate
- `yld`: Annual yield
- `frequency`: Number of coupon payments per year (1, 2, or 4)
- `basis` *(optional)*: Day count basis (default 0)

**Returns:** Double - Modified duration

**Excel Documentation:** [MDURATION function](https://support.microsoft.com/en-us/office/mduration-function-b3786a69-4f20-469a-94ad-33e5b90a763c)

**Implementation Status:** 🔄 Stub (returns #CALC! error - requires complex bond duration calculation)

---

### DISC

Calculates the discount rate for a security.

**Syntax:** `DISC(settlement, maturity, pr, redemption, basis)`

**Parameters:**
- `settlement`: Security's settlement date (as Excel serial date)
- `maturity`: Security's maturity date
- `pr`: Security's price per $100 face value
- `redemption`: Redemption value per $100 face value
- `basis`: Day count basis (0-4)

**Returns:** Double - Discount rate

**Examples:**
```swift
// Discount rate for security priced at $97.50
let discount = evaluator.evaluate("=DISC(DATE(2024,2,1), DATE(2024,8,1), 97.50, 100, 2)")

// Using cell references
let discRate = evaluator.evaluate("=DISC(A1, B1, 98, 100, 2)")
```

**Excel Documentation:** [DISC function](https://support.microsoft.com/en-us/office/disc-function-71fce9f3-3f05-4acf-a5a3-eac6ef4daa53)

**Implementation Status:** ✅ Full implementation

---

### INTRATE

Calculates the interest rate for a fully invested security.

**Syntax:** `INTRATE(settlement, maturity, investment, redemption, basis)`

**Parameters:**
- `settlement`: Security's settlement date (as Excel serial date)
- `maturity`: Security's maturity date
- `investment`: Amount invested
- `redemption`: Amount received at maturity
- `basis`: Day count basis (0-4)

**Returns:** Double - Interest rate

**Examples:**
```swift
// Interest rate for $10,000 investment returning $10,500
let rate = evaluator.evaluate("=INTRATE(DATE(2024,1,1), DATE(2024,7,1), 10000, 10500, 2)")

// Using actual/360 basis
let rateActual360 = evaluator.evaluate("=INTRATE(44927, 45109, 100000, 102500, 2)")
```

**Excel Documentation:** [INTRATE function](https://support.microsoft.com/en-us/office/intrate-function-5cb34dde-a221-4cb1-8d8f-7ff8f86c4f99)

**Implementation Status:** ✅ Full implementation

---

### RECEIVED

Calculates the amount received at maturity for a fully invested security.

**Syntax:** `RECEIVED(settlement, maturity, investment, discount, basis)`

**Parameters:**
- `settlement`: Security's settlement date (as Excel serial date)
- `maturity`: Security's maturity date
- `investment`: Amount invested
- `discount`: Security's discount rate
- `basis`: Day count basis (0-4)

**Returns:** Double - Amount received at maturity

**Examples:**
```swift
// Amount received for $10,000 investment at 5% discount
let received = evaluator.evaluate("=RECEIVED(DATE(2024,1,1), DATE(2024,7,1), 10000, 0.05, 2)")

// T-bill maturity value
let maturityValue = evaluator.evaluate("=RECEIVED(44927, 45109, 98000, 0.0525, 2)")
```

**Excel Documentation:** [RECEIVED function](https://support.microsoft.com/en-us/office/received-function-7a3f8b93-6611-4f81-8576-828312c9b5e5)

**Implementation Status:** ✅ Full implementation

---

### COUPDAYBS

Returns the number of days from the beginning of the coupon period to the settlement date.

**Syntax:** `COUPDAYBS(settlement, maturity, frequency, [basis])`

**Parameters:**
- `settlement`: Security's settlement date (as Excel serial date)
- `maturity`: Security's maturity date
- `frequency`: Number of coupon payments per year (1, 2, or 4)
- `basis` *(optional)*: Day count basis (default 0)

**Returns:** Double - Number of days

**Excel Documentation:** [COUPDAYBS function](https://support.microsoft.com/en-us/office/coupdaybs-function-eb9a8dfb-2fb2-4c61-8e5d-690b320cf872)

**Implementation Status:** 🔄 Stub (returns #CALC! error - requires coupon calculation)

---

### COUPDAYS

Returns the number of days in the coupon period that contains the settlement date.

**Syntax:** `COUPDAYS(settlement, maturity, frequency, [basis])`

**Parameters:**
- `settlement`: Security's settlement date (as Excel serial date)
- `maturity`: Security's maturity date
- `frequency`: Number of coupon payments per year (1, 2, or 4)
- `basis` *(optional)*: Day count basis (default 0)

**Returns:** Double - Number of days in coupon period

**Excel Documentation:** [COUPDAYS function](https://support.microsoft.com/en-us/office/coupdays-function-cc64380b-315b-4e7b-950c-b30b0a76f671)

**Implementation Status:** 🔄 Stub (returns #CALC! error - requires coupon calculation)

---

### COUPDAYSNC

Returns the number of days from the settlement date to the next coupon date.

**Syntax:** `COUPDAYSNC(settlement, maturity, frequency, [basis])`

**Parameters:**
- `settlement`: Security's settlement date (as Excel serial date)
- `maturity`: Security's maturity date
- `frequency`: Number of coupon payments per year (1, 2, or 4)
- `basis` *(optional)*: Day count basis (default 0)

**Returns:** Double - Number of days to next coupon

**Excel Documentation:** [COUPDAYSNC function](https://support.microsoft.com/en-us/office/coupdaysnc-function-5ab3f0b2-029f-4a8b-bb65-47d525eea547)

**Implementation Status:** 🔄 Stub (returns #CALC! error - requires coupon calculation)

---

### COUPNCD

Returns the next coupon date after the settlement date.

**Syntax:** `COUPNCD(settlement, maturity, frequency, [basis])`

**Parameters:**
- `settlement`: Security's settlement date (as Excel serial date)
- `maturity`: Security's maturity date
- `frequency`: Number of coupon payments per year (1, 2, or 4)
- `basis` *(optional)*: Day count basis (default 0)

**Returns:** Double - Next coupon date (as Excel serial date)

**Excel Documentation:** [COUPNCD function](https://support.microsoft.com/en-us/office/coupncd-function-fd962fef-506b-4d9d-8590-16df5393691f)

**Implementation Status:** 🔄 Stub (returns #CALC! error - requires coupon calculation)

---

### COUPPCD

Returns the previous coupon date before the settlement date.

**Syntax:** `COUPPCD(settlement, maturity, frequency, [basis])`

**Parameters:**
- `settlement`: Security's settlement date (as Excel serial date)
- `maturity`: Security's maturity date
- `frequency`: Number of coupon payments per year (1, 2, or 4)
- `basis` *(optional)*: Day count basis (default 0)

**Returns:** Double - Previous coupon date (as Excel serial date)

**Excel Documentation:** [COUPPCD function](https://support.microsoft.com/en-us/office/couppcd-function-2eb50473-6ee9-4052-a206-77a9a385d5b3)

**Implementation Status:** 🔄 Stub (returns #CALC! error - requires coupon calculation)

---

### COUPNUM

Returns the number of coupons payable between the settlement date and maturity date.

**Syntax:** `COUPNUM(settlement, maturity, frequency, [basis])`

**Parameters:**
- `settlement`: Security's settlement date (as Excel serial date)
- `maturity`: Security's maturity date
- `frequency`: Number of coupon payments per year (1, 2, or 4)
- `basis` *(optional)*: Day count basis (default 0)

**Returns:** Double - Number of coupons

**Excel Documentation:** [COUPNUM function](https://support.microsoft.com/en-us/office/coupnum-function-a90af57b-de53-4969-9c99-dd6139db2522)

**Implementation Status:** 🔄 Stub (returns #CALC! error - requires coupon calculation)

---

## See Also

- <doc:Mathematical> - Mathematical and trigonometric functions
- <doc:Statistical> - Statistical analysis functions
- <doc:DateTime> - Date and time functions for financial calculations
- <doc:Lookup> - Lookup functions for financial data retrieval

## Topics

### Implementation Reference

- ``FormulaEvaluator/evaluatePMT(_:)``
- ``FormulaEvaluator/evaluateNPV(_:)``
- ``FormulaEvaluator/evaluateIRR(_:)``
- ``FormulaEvaluator/evaluateSLN(_:)``
