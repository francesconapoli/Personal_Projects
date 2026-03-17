Attribute VB_Name = "modGlobals"
'===============================================================================
' Module: modGlobals
' Purpose: Global constants, types, and variables for MCSimAddin
'===============================================================================
Option Explicit

' ---- Version ----
Public Const ADDIN_NAME As String = "MCSimAddin"
Public Const ADDIN_VERSION As String = "1.0.0"

' ---- Simulation State Enum ----
Public Enum SimulationState
    ssIdle = 0
    ssRunning = 1
    ssPaused = 2
    ssStopped = 3
    ssComplete = 4
End Enum

' ---- Sampling Method Enum ----
Public Enum SamplingMethod
    smMonteCarlo = 0
    smLatinHypercube = 1
End Enum

' ---- Distribution Type Enum ----
Public Enum DistributionType
    dtNone = 0
    ' Continuous
    dtNormal = 1
    dtUniform = 2
    dtTriangular = 3
    dtLognormal = 4
    dtExponential = 5
    dtBeta = 6
    dtBetaGeneral = 7
    dtGamma = 8
    dtWeibull = 9
    dtPERT = 10
    dtPareto = 11
    dtLogistic = 12
    dtRayleigh = 13
    dtExtValue = 14
    dtExtValueMin = 15
    dtChiSquared = 16
    dtStudentT = 17
    ' Discrete
    dtBinomial = 30
    dtPoisson = 31
    dtNegBinomial = 32
    dtDiscrete = 33
    dtDUniform = 34
    dtIntUniform = 35
    ' Empirical
    dtCumul = 50
    dtCumulD = 51
    dtGeneral = 52
    dtHistogrm = 53
    dtCompound = 54
End Enum

' ---- Sensitivity Method Enum ----
Public Enum SensitivityMethod
    snRegression = 0
    snRankCorrelation = 1
    snContributionToVariance = 2
    snChangeInStatistic = 3
End Enum

' ---- Convergence Statistic Enum ----
Public Enum ConvergenceStat
    csMean = 0
    csStdDev = 1
    csPercentile5 = 2
    csPercentile95 = 3
End Enum

' ---- Input Definition UDT ----
Public Type SimInput
    CellAddress As String
    SheetName As String
    DistType As DistributionType
    Params() As Double
    OriginalValue As Variant
    InputName As String
    CorrelationGroup As Long
End Type

' ---- Output Definition UDT ----
Public Type SimOutput
    CellAddress As String
    SheetName As String
    OutputName As String
    IterationData() As Double
    StatMean As Double
    StatStdDev As Double
    StatMin As Double
    StatMax As Double
    StatSkewness As Double
    StatKurtosis As Double
    StatMedian As Double
    StatMode As Double
    StatVariance As Double
    Percentiles(0 To 100) As Double
End Type

' ---- Simulation Settings UDT ----
Public Type SimSettings
    Iterations As Long
    Simulations As Long
    SampMethod As SamplingMethod
    RandomSeed As Long
    UseFixedSeed As Boolean
    ConvergenceEnabled As Boolean
    ConvergenceTolerance As Double
    ConvergenceConfidence As Double
    ConvergenceTestEvery As Long
    ConvergenceStatistic As ConvergenceStat
    UpdateFrequency As Long
    CollectDistributionSamples As Boolean
End Type

' ---- Correlation Entry UDT ----
Public Type CorrelationEntry
    Input1Index As Long
    Input2Index As Long
    Coefficient As Double
End Type

' ---- Global Variables ----
Public g_SimState As SimulationState
Public g_Settings As SimSettings
Public g_Inputs() As SimInput
Public g_Outputs() As SimOutput
Public g_InputCount As Long
Public g_OutputCount As Long
Public g_Correlations() As CorrelationEntry
Public g_CorrelationCount As Long
Public g_CurrentIteration As Long
Public g_CurrentSimulation As Long
Public g_AllSimData() As Variant
Public g_ConvergenceReached As Boolean
Public g_SimStartTime As Double
Public g_SimEndTime As Double

' ---- Latin Hypercube permutation arrays ----
Public g_LHSPermutations() As Long

' ---- Application state backup ----
Public g_OrigScreenUpdating As Boolean
Public g_OrigCalculation As XlCalculation
Public g_OrigEvents As Boolean
Public g_OrigStatusBar As Variant

'===============================================================================
' Sub: InitializeDefaults
' Purpose: Set default simulation settings
'===============================================================================
Public Sub InitializeDefaults()
    With g_Settings
        .Iterations = 10000
        .Simulations = 1
        .SampMethod = smLatinHypercube
        .RandomSeed = 12345
        .UseFixedSeed = False
        .ConvergenceEnabled = False
        .ConvergenceTolerance = 0.03
        .ConvergenceConfidence = 0.95
        .ConvergenceTestEvery = 100
        .ConvergenceStatistic = csMean
        .UpdateFrequency = 100
        .CollectDistributionSamples = True
    End With

    g_SimState = ssIdle
    g_InputCount = 0
    g_OutputCount = 0
    g_CorrelationCount = 0
    g_CurrentIteration = 0
    g_CurrentSimulation = 0
    g_ConvergenceReached = False
End Sub

'===============================================================================
' Sub: ResetSimulation
' Purpose: Clear all simulation data
'===============================================================================
Public Sub ResetSimulation()
    g_InputCount = 0
    g_OutputCount = 0
    g_CorrelationCount = 0
    Erase g_Inputs
    Erase g_Outputs
    Erase g_Correlations
    g_CurrentIteration = 0
    g_CurrentSimulation = 0
    g_ConvergenceReached = False
    g_SimState = ssIdle
End Sub
