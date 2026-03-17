# MCSimAddin - Monte Carlo Simulation Add-in for Excel

A comprehensive Monte Carlo simulation add-in for Microsoft Excel that replicates the functionality of @RISK from Palisade (now Lumivero). Built entirely in VBA, this add-in provides probability distribution functions, simulation engine, statistical analysis, charting, sensitivity analysis, and a custom ribbon UI.

## Features

### Distribution Functions (25+ distributions)

#### Continuous Distributions
| Function | Parameters | Description |
|----------|-----------|-------------|
| `MCNormal(mean, stdDev)` | Mean, Standard Deviation | Normal/Gaussian distribution |
| `MCUniform(min, max)` | Minimum, Maximum | Continuous uniform distribution |
| `MCTriang(min, mostLikely, max)` | Min, Mode, Max | Triangular distribution |
| `MCLognorm(mean, stdDev)` | Arithmetic Mean, Std Dev | Lognormal distribution |
| `MCExpon(beta)` | Beta (mean) | Exponential distribution |
| `MCBeta(alpha1, alpha2)` | Alpha1, Alpha2 | Beta distribution on [0,1] |
| `MCBetaGeneral(a1, a2, min, max)` | Alpha1, Alpha2, Min, Max | Beta on custom range |
| `MCGamma(alpha, beta)` | Shape, Scale | Gamma distribution |
| `MCWeibull(shape, scale)` | Shape, Scale | Weibull distribution |
| `MCPERT(min, mostLikely, max)` | Min, Mode, Max | PERT (modified beta) |
| `MCPareto(shape, scale)` | Shape (theta), Scale (xm) | Pareto distribution |
| `MCLogistic(mean, scale)` | Mean, Scale | Logistic distribution |
| `MCRayleigh(scale)` | Scale | Rayleigh distribution |
| `MCExtValue(alpha, beta)` | Location, Scale | Gumbel Maximum distribution |
| `MCExtValueMin(alpha, beta)` | Location, Scale | Gumbel Minimum distribution |
| `MCChiSq(df)` | Degrees of Freedom | Chi-squared distribution |
| `MCStudentT(df)` | Degrees of Freedom | Student's t-distribution |

#### Discrete Distributions
| Function | Parameters | Description |
|----------|-----------|-------------|
| `MCBinomial(n, p)` | Trials, Probability | Binomial distribution |
| `MCPoisson(lambda)` | Lambda (mean) | Poisson distribution |
| `MCNegBinom(s, p)` | Successes, Probability | Negative binomial |
| `MCDiscrete(values, probs)` | Value array, Probability array | Custom discrete |
| `MCDUniform(values)` | Value array | Discrete uniform |
| `MCIntUniform(min, max)` | Min integer, Max integer | Integer uniform |

#### Empirical/Special Distributions
| Function | Parameters | Description |
|----------|-----------|-------------|
| `MCCumul(xValues, probValues)` | X values, Cumulative probs | Cumulative ascending |
| `MCCumulD(xValues, probValues)` | X values, Cumulative probs | Cumulative descending |
| `MCHistogrm(min, max, binProbs)` | Min, Max, Bin probabilities | Histogram distribution |
| `MCCompound(freqCell, sevCell)` | Frequency cell, Severity cell | Compound distribution |
| `MCSimtable(values)` | Value array | Scenario table |

### Output & Statistics Functions

| Function | Description |
|----------|-------------|
| `MCOutput("name")` | Mark cell as simulation output |
| `MCAddOutput(cellRef, "name")` | Register output by reference |
| `MCMean(cellRef)` | Mean of simulation results |
| `MCStdDev(cellRef)` | Standard deviation |
| `MCVariance(cellRef)` | Variance |
| `MCMin(cellRef)` | Minimum value |
| `MCMax(cellRef)` | Maximum value |
| `MCMedian(cellRef)` | Median (50th percentile) |
| `MCMode(cellRef)` | Mode estimate |
| `MCSkewness(cellRef)` | Skewness |
| `MCKurtosis(cellRef)` | Excess kurtosis |
| `MCPercentile(cellRef, pct)` | Ascending percentile (0-1) |
| `MCPercentileD(cellRef, pct)` | Descending percentile |
| `MCTarget(cellRef, value)` | P(output <= value) |
| `MCTargetD(cellRef, value)` | P(output > value) |
| `MCData(cellRef, iter, sim)` | Value from specific iteration |

### Simulation Engine
- **Monte Carlo** and **Latin Hypercube** sampling methods
- Mersenne Twister pseudo-random number generator
- Configurable iterations (up to 1,000,000)
- Multiple simulation runs
- Fixed seed for reproducibility
- Convergence monitoring with configurable tolerance
- Real-time progress bar in status bar
- Pause/Resume/Stop controls

### Sensitivity Analysis
- **Regression Coefficients** (standardized)
- **Rank Correlation** (Spearman)
- **Contribution to Variance** (sequential sum of squares)
- **Change in Output Statistic**

### Charts & Visualization
- **Histogram** (relative frequency)
- **Cumulative Distribution** (ascending/descending S-curves)
- **Overlay Chart** (histogram + cumulative on dual axes)
- **Tornado Diagram** (sensitivity ranking)
- **Spider Plot** (multi-variable sensitivity)
- **Scatter Plot** (output vs output)
- **Statistics Summary Table** with formatting

### Additional Features
- **Distribution Fitting** - Fit data to best distribution (KS test)
- **Correlation Support** - Spearman rank correlations between inputs (Iman-Conover method)
- **Model Window** - View all inputs, outputs, and correlations
- **Data Export** - Export raw iteration data to worksheet
- **Custom Ribbon Tab** - Full UI with Define, Simulation, and Results groups

## Installation

### Prerequisites
- Microsoft Excel 2010 or later (Windows)
- VBA macros must be enabled
- Trust access to VBA project object model (for build scripts)

### Method 1: Build Script (Recommended)

1. **Enable VBA Trust Settings**:
   - Excel > File > Options > Trust Center > Trust Center Settings
   - Macro Settings > Enable all macros
   - Check "Trust access to the VBA project object model"

2. **Run the build script** (one of):
   ```
   # PowerShell
   .\BuildAddin.ps1

   # VBScript
   cscript BuildAddin.vbs
   ```

3. **Install the Ribbon** (optional, for custom tab):
   - Download [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor)
   - Open `MCSimAddin.xlam` in the editor
   - Insert `src/CustomUI/customUI14.xml` as `customUI14`
   - Save

4. **Activate the add-in**:
   - Excel > File > Options > Add-ins
   - Manage: Excel Add-ins > Go...
   - Browse... > Select `MCSimAddin.xlam`
   - Check it and click OK

### Method 2: Manual VBA Import

1. Open Excel, press **Alt+F11**
2. Create a new workbook
3. For each `.bas` file in `src/Modules/`: Right-click project > Import File
4. For each `.frm` file in `src/Forms/`: Right-click project > Import File
5. Copy `ThisWorkbook.cls` code into the ThisWorkbook module
6. Save As > Excel Add-in (.xlam)

### Method 3: Quick Install Macro

1. Open Excel, press **Alt+F11**
2. Insert > Module
3. Paste the contents of `ManualInstall.bas`
4. Edit the `SOURCE_PATH` constant to your actual path
5. Run `InstallMCSimAddin`

## Quick Start

### 1. Add Distribution Inputs
In any cell, enter a distribution function:
```
=MCNormal(100, 15)        ' Normal with mean=100, std=15
=MCTriang(50, 75, 120)    ' Triangular min=50, mode=75, max=120
=MCPERT(10, 20, 40)       ' PERT min=10, mode=20, max=40
=MCUniform(0, 100)        ' Uniform between 0 and 100
```

### 2. Mark Outputs
Add `+MCOutput("name")` to any formula cell you want to track:
```
=SUM(A1:A10)+MCOutput("Total Cost")
=B5*C5+MCOutput("Revenue")
```

### 3. Add Statistics Functions
Place these anywhere to see results after simulation:
```
=MCMean(F19)              ' Mean of cell F19's distribution
=MCPercentile(F19, 0.95)  ' 95th percentile
=MCTarget(F19, 150000)    ' P(F19 <= 150000)
```

### 4. Run Simulation
- Click **Run** in the MC Simulation ribbon tab
- Or run `RunSimulation` from VBA

### 5. View Results
- Click **Browse Results** for statistics and charts
- Click **Charts** for visualizations
- Click **Sensitivity** for tornado/spider analysis

## Sample Model
Import `SampleModel.bas` and run `CreateSampleModel` to generate a complete project cost estimation example with all features demonstrated.

## Architecture

```
MonteCarloAddin/
├── src/
│   ├── Modules/
│   │   ├── modGlobals.bas          # Global types, enums, variables
│   │   ├── modRandomGenerators.bas  # Mersenne Twister, Box-Muller, special functions
│   │   ├── modDistributions.bas     # All 25+ distribution worksheet functions
│   │   ├── modOutputFunctions.bas   # Output and statistic functions
│   │   ├── modSimEngine.bas         # Core simulation engine, LHS, convergence
│   │   ├── modSensitivity.bas       # Sensitivity analysis methods
│   │   ├── modCharting.bas          # Chart generation (histogram, tornado, etc.)
│   │   ├── modDistFitting.bas       # Distribution fitting (KS test)
│   │   └── modRibbonCallbacks.bas   # Ribbon UI callback procedures
│   ├── Forms/
│   │   ├── frmSettings.frm          # Simulation settings dialog
│   │   ├── frmResults.frm           # Results browser with stats & charts
│   │   ├── frmAddInput.frm          # Add distribution input dialog
│   │   ├── frmCorrelation.frm       # Define correlations dialog
│   │   ├── frmModelWindow.frm       # Model overview (inputs/outputs/correlations)
│   │   ├── frmSensitivity.frm       # Sensitivity analysis dialog
│   │   └── frmOutputSelect.frm      # Output selector dialog
│   ├── CustomUI/
│   │   └── customUI14.xml           # Ribbon XML definition
│   └── ThisWorkbook.cls             # Workbook event handlers
├── BuildAddin.ps1                   # PowerShell build script
├── BuildAddin.vbs                   # VBScript build script
├── ManualInstall.bas                # Manual install macro
├── SampleModel.bas                  # Sample model generator
└── README.md                        # This file
```

## Function Reference vs @RISK

| @RISK Function | MCSimAddin Equivalent |
|---------------|----------------------|
| RiskNormal | MCNormal |
| RiskUniform | MCUniform |
| RiskTriang | MCTriang |
| RiskLognorm | MCLognorm |
| RiskExpon | MCExpon |
| RiskBeta | MCBeta |
| RiskBetaGeneral | MCBetaGeneral |
| RiskGamma | MCGamma |
| RiskWeibull | MCWeibull |
| RiskPERT | MCPERT |
| RiskPareto | MCPareto |
| RiskLogistic | MCLogistic |
| RiskRayleigh | MCRayleigh |
| RiskExtValue | MCExtValue |
| RiskExtValueMin | MCExtValueMin |
| RiskChiSq | MCChiSq |
| RiskStudentT | MCStudentT |
| RiskBinomial | MCBinomial |
| RiskPoisson | MCPoisson |
| RiskNegbinom | MCNegBinom |
| RiskDiscrete | MCDiscrete |
| RiskDUniform | MCDUniform |
| RiskIntUniform | MCIntUniform |
| RiskCumul | MCCumul |
| RiskCumulD | MCCumulD |
| RiskHistogrm | MCHistogrm |
| RiskCompound | MCCompound |
| RiskSimtable | MCSimtable |
| RiskOutput | MCOutput |
| RiskMean | MCMean |
| RiskStdDev | MCStdDev |
| RiskMin | MCMin |
| RiskMax | MCMax |
| RiskVariance | MCVariance |
| RiskSkewness | MCSkewness |
| RiskKurtosis | MCKurtosis |
| RiskPercentile | MCPercentile |
| RiskPercentileD | MCPercentileD |
| RiskTarget | MCTarget |
| RiskTargetD | MCTargetD |
| RiskMode | MCMode |
| RiskData | MCData |
| RiskMedian | MCMedian |

## License
This project is provided as-is for educational and professional use.
