Attribute VB_Name = "mdlConstants"
'
'    Primary Author Daniel Villa, dlvilla@sandia.gov, 505-340-9162
'
'    Copyright 2018 National Technology & Engineering Solutions of Sandia, LLC (NTESS).
'    Under the terms of Contract DE-NA0003525 with NTESS, the U.S. Government retains
'    certain rights in this software
'
'    Redistribution and use in source and binary forms, with or without modification, are permitted
'    provided that the following conditions are met:
'
'    1. Redistributions of source code must retain the above copyright notice, this list of
'       conditions and the following disclaimer.
'
'    2. Redistributions in binary form must reproduce the above copyright notice,
'       this list of conditions and the following disclaimer in the documentation and/or other
'       materials provided with the distribution.
'
'    3. Neither the name of the copyright holder nor the names of its contributors may be used
'       to endorse or promote products derived from this software without specific prior written
'       permission.
'
'    THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES,
'    INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
'    DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
'    SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
'    SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
'    WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE
'    USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
'    -------------------- End of Copyright
'
'    The above copyright is the "New BSD License" or BSD-3-Clause for open source software obtained on 2/26/2018 from:
'    https://opensource.org/licenses/BSD-3-Clause
'    It is formally recognized by the Open Source Initiative.



' All constants used by the code

' Troubleshooting
Public Const glbTroubleshoot = False ' !@#$ Never run the model with this True unless you are troubleshooting.

' Initial guess parameters - these may need to be adjusted if convergence cannot be achieved!
Public Const glbInitialGuessAtFractionofMDMassFlow = 0.00001
Public Const glbInitialGuessCondensateThickFraction = 0.05  'This may need adjustment if
Public Const glbInitialGuessOfSaturationPressFrac = 0.5
Public Const glbInitialFractionToColdOrHotForInterfaceTemperature = 0.1

Public Const glbMaxHeatFlowFractionForInitialCondition = 0.5   ' This constant may eventually be calculated. It is VERY IMPORTANT WHETHER THE SOLUTION WILL CONVERGE!!!
                        ' FOR CO-FLOW a value of 0.5 works well. Values greater than 0.5
                        ' for very long counterflow setups, 0.9 is needed!
Public Const glbNewtonReductionFactor = 1    ' Set this value <1 to reduce the step increment used by Newton's method
                                             ' this will greatly increase the simulation time but may lead to a
                                             ' valid solution for some cases.

' UPPER AND LOWER LIMITS ON INPUT AND VARIABLES!
Public Const glbEndOnError = True
Public Const glbIncludeMsgBox = True

Public Const glbNumberDivUpperLimit = 400
Public Const glbNumberDivLowerLimit = 1

Public Const glbThermConductUpperLimit = 10000
Public Const glbThermConductLowerLimit = 0.0001

Public Const glbThickUpperLimit = 100   'mm
Public Const glbThickLowerLimit = 0.0001 'mm

Public Const glbPlateDimensionsUpperLimit = 10 'meters
Public Const glbPlateDimensionsLowerLimit = 0.01 'meters

Public Const glbNumberOfLayersUpperLimit = 1000
Public Const glbNumberOfLayersLowerLimit = 1

Public Const glbAmbientTemperatureUpperLimit = 100 '100C is really hot!
Public Const glbAmbientTemperatureLowerLimit = -40 '-40C is really cold! I cannot imagine application outside this range.

Public Const glbDirectMassTransferCoeffUpperLimit = 0.00001
Public Const glbDirectMassTransferCoeffLowerLimit = 0.0000000001

Public Const glbPorosityUpperLimit = 0.99
Public Const glbPorosityLowerLimit = 0

'no upper limit since the upper limit is the thickness of the material.
Public Const glbMeanPoreRadLowerLimit = 0

Public Const glbMaxReynolds = 2000
Public Const glbMaxTurbulentReynolds = 10000000# 'This is a complete guess,

Public Const glbWaterTempInUpperLimit = 100
Public Const glbWaterTempInLowerLimit = 0
Public Const glbWaterMassFlowUpperLimit = 100
Public Const glbWaterMassFlowLowerLimit = 0.000001
Public Const glbWaterVolumeFlowLowerLimit = 0.0000009
Public Const glbWaterVolumeFlowUpperLimit = 4800
Public Const glbWaterPresInUpperLimit = 10
Public Const glbWaterPresInLowerLimit = 0
Public Const glbWaterPHUpperLimit = 14
Public Const glbWaterPHLowerLimit = 0#
Public Const glbWaterSalinityUpperLimit = 300
Public Const glbWaterSalinityLowerLimit = 0
Public Const glbWaterConductanceUpperLimit = 100000
Public Const glbWaterConductanceLowerLimit = 1
Public Const glbAirPressureLowerLimit = 0.000000000001 ' I do not think anyone is ever going to achieve this high of a vacuum for MD application
' We will make the upper limit a function of the hot side water stream pressure.


Public Const glbFilamentIntersectAngleLowerLimit = 0
Public Const glbFilamentIntersectAngleUpperLimit = 90

'Units conversions
Public Const glbMiliMeterToMeter = 0.001
Public Const glbCelciusToKelvinOffset = 273.15
Public Const glbMiligramPerLiterToKilogramPerMeter3 = 0.001
Public Const glbGramPerKilogramToFraction = 0.001
Public Const glbBarToPascal = 100000
Public Const glbMicroSiemenToSiemen = 0.000001
Public Const glbmeters3ToLiters = 1000
Public Const glbLPMToMeter3PerSecond = 1.66666666667E-05
Public Const glbSecondsInHour = 3600
Public Const glbSecondPerMinute = 60
Public Const glbcm2ToMeter2 = 1 / 100 ^ 2
Public Const glbDegreesToRadians = 3.14159265358979 / 180

' Error Thresholds
Public Const glbAcceptableEnergyError = 0.01 'Watts
Public Const glbSatErr = 0.05

' Physical Constants
' standard gravitational constant (no altitude dependence)
Public Const glbGravity = 9.80665
Public Const glbGasConstant = 8.314 'J/mol-K
Public Const glbAirMolecularWeight = 0.02897 'kg/mol
Public Const glbWaterMolecularWeight = 0.01802 'kg/mol
Public Const glbSodiumChlorideMolecularWeight = 0.05844277 'kg/mol
Public Const glbBoltzmann = 1.380658E-23 'J/K Boltzmann's constant
Public Const glbWaterVaporCollisionDiameter = 0.0000000002641 'Collision diameter for water vapor (meters)

' Mathematical Parameters
Public Const glbMaxNewtonIterations = 35
Public Const glbMaxReductionInStepSizeDueToConstraint = 0.1
Public Const glbNewtonMaxIncrementToMaxOrMin = 0.5
Public Const glbDerivativeIncrement = 0.0001
Public Const glbNewtonConvergenceCriterion = 0.0001
Public Const glbAbsConvergCriterionEnergy = 0.0001 ' Converge regardless of the relative error.
Public Const glbAbsConvergCriterionMassFlow = 0.000000001
Public Const glbAbsConvergCriterionSalinity = 0.000001
Public Const glbAbsConvergCriterionThickness = 0.0000000001
Public Const glbAbsConvergCriterionTemperature = 0.00001
Public Const glbAbsConvergCriterionPressure = 0.01

Public Const glbZeroThresholdForCubicSolution = 1E-20 ' Any values smaller than this in the coefficients will be considered to be zero.
Public Const glbPi = 3.14159265358979

' OTHER
Public Const glbINVALID_VALUE = -999
Public Const glbNumberOfLinesInExcelToSearch = 25
Public Const glbNISTReferenceTemperature = 273.15
Public Const glbFractionOfMaxPressureToStart = 0.99
'SET THIS TO FALSE TO SAVE TIME!!!!
Public Const glbEvaluateEquationsAtEnd = True
Public Const glbUseTurbulentNusseltForForcedConvection = True
' THIS IS AN OLD FEATURE THAT CONSIDERS EACH LAYER TO BE A MEMBRANE WITH A FULL COLD AND HOT STREAM.
' WHEN THIS IS TRUE, EACH LAYER IS A HOT AND COLD STREAM BUT MD Configurations (AGMD, DCMD) TOUCH BOTH SIDES
' OF ALL HOT AND COLD STREAMS!!! KEEP FALSE!!!!
Public Const glbUseIndependentLayers = False
Public Const glbDebugFileName = "MultiConfigurationMembraneDistillation_Output.txt"


' Values that need to be initialized

Public glbMembraneDistillationTypes() As String
Public glbSensitivityStudies() As String
Public glbSensitivityStudyHighValue() As Double
Public glbSensitivityStudyNumDiv() As Long
Public glbSensitivityStudyLowValue() As Double
Public glbSensitivityStudyAlterRange() As String

Public Sub GlobalArrays()

    ReDim glbMembraneDistillationTypes(1 To 2)
    glbMembraneDistillationTypes(1) = "Direct Contact"
    glbMembraneDistillationTypes(2) = "Air Gap"

    ReDim glbSensitivityStudies(1 To 2)
    ReDim glbSensitivityStudyHighValue(1 To 2)
    ReDim glbSensitivityStudyNumDiv(1 To 2)
    ReDim glbSensitivityStudyLowValue(1 To 2)
    ReDim glbSensitivityStudyAlterRange(1 To 2)
    
    glbSensitivityStudies(1) = "Horizontal Length"
    glbSensitivityStudies(2) = "Vertical Length"
    
    glbSensitivityStudyHighValue(1) = 5.5
    glbSensitivityStudyHighValue(2) = 5.5
    
    glbSensitivityStudyLowValue(1) = 0.5
    glbSensitivityStudyLowValue(2) = 0.5
    
    glbSensitivityStudyNumDiv(1) = 11
    glbSensitivityStudyNumDiv(2) = 11
    
    glbSensitivityStudyAlterRange(1) = "Horizontal_Length"
    glbSensitivityStudyAlterRange(2) = "Vertical_Length"
    
End Sub

Public Sub GlobalEquationSetConvergenceCriterion(SysEq As clsSystemEquations, ConvergCrit() As Double)

Dim i As Long
Dim NumBC As Long
Dim NumCVEq As Long
Dim iMod As Long

NumBC = SysEq.Connectivity.NumberOfBoundaryConditions
NumCVEq = SysEq.NumberEquations - NumBC

ReDim ConvergCrit(1 To SysEq.NumberEquations)

'Public Const glbAbsConvergCriterionEnergy = 0.0001 ' Converge regardless of the relative error.
'Public Const glbAbsConvergCriterionMassFlow = 0.000000001
'Public Const glbAbsConvergCriterionSalinity = 0.000001
'Public Const glbAbsConvergCriterionThickness = 0.0000000001
'Public Const glbAbsConvergCriterionTemperature = 0.00001

' The order of "SysEq.Equations" can change the order of the output ConvergCrit!
For i = 1 To NumBC
   iMod = i Mod 3
   If iMod = 1 Then
      ConvergCrit(i) = glbAbsConvergCriterionTemperature
   ElseIf iMod = 2 Then
      ConvergCrit(i) = glbAbsConvergCriterionMassFlow
   ElseIf iMod = 0 Then
      ConvergCrit(i) = glbAbsConvergCriterionSalinity
   End If
Next

For i = NumBC + 1 To NumBC + NumCVEq
   iMod = i Mod 6
   If iMod = 1 Or iMod = 2 Then
      ConvergCrit(i) = glbAbsConvergCriterionEnergy
   ElseIf iMod = 3 Or iMod = 4 Then
      ConvergCrit(i) = glbAbsConvergCriterionMassFlow
   ElseIf iMod = 5 Or iMod = 0 Then
      ConvergCrit(i) = glbAbsConvergCriterionSalinity
   End If
Next
        
End Sub
