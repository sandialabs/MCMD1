VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'        Primary Author Daniel Villa, dlvilla@sandia.gov, 505-340-9162
'
'        Copyright (year first published) Sandia Corporation. Under the terms of Contract DE-AC04-94AL85000,
'        there is a non-exclusive license for use of this work by or on behalf of the U.S. Government.
'        Export of this data may require a license from the United States Government.
'
'                                                       NOTICE:
'
'        For five (5) years from 02/09/2015, the United States Government is granted for itself and others
'        acting on its behalf a paid-up, nonexclusive, irrevocable worldwide license in this data to reproduce,
'        prepare derivative works, and perform publicly and display publicly, by or on behalf of the Government.
'        There is provision for the possible extension of the term of this license. Subsequent to that period or
'        any extension granted, the United States Government is granted for itself and others acting on its behalf
'        a paid-up, nonexclusive, irrevocable worldwide license in this data to reproduce, prepare derivative works,
'        distribute copies to the public, perform publicly and display publicly, and to permit others to do so. The
'        specific term of the license can be identified by inquiry made to Sandia Corporation or DOE.
     
 '       NEITHER THE UNITED STATES GOVERNMENT, NOR THE UNITED STATES DEPARTMENT OF ENERGY, NOR SANDIA CORPORATION,
 '       NOR ANY OF THEIR EMPLOYEES, MAKES ANY WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LEGAL RESPONSIBILITY
 '       FOR THE ACCURACY, COMPLETENESS, OR USEFULNESS OF ANY INFORMATION, APPARATUS, PRODUCT, OR PROCESS DISCLOSED,
 '       OR REPRESENTS THAT ITS USE WOULD NOT INFRINGE PRIVATELY OWNED RIGHTS.

Option Explicit
Option Base 1

Private Function TotalMDMassFlow(SysEq As clsSystemEquations, LiterPerHourFlow As Double) As Double


   Dim i As Long
   Dim MDMassFlow As Double
   Dim MDVolumetricFlow As Double
   Dim mM As Double
   Dim rho As Double
   Dim CVP As clsControlVolumePair
   
   
   MDMassFlow = 0
   MDVolumetricFlow = 0
   
   For i = 1 To SysEq.Connectivity.NumberOfControlVolumePairs
      Set CVP = SysEq.ControlVolumePair(i)
      
      mM = SysEq.ControlVolumePair(i).MembraneDistillationMassFlow
      
      MDVolumetricFlow = MDVolumetricFlow + mM / CVP.MembraneAverageDensity
      MDMassFlow = MDMassFlow + mM

   Next i
   
   TotalMDMassFlow = MDMassFlow
   ' Convert volumetric flow to liters per hour
   LiterPerHourFlow = MDVolumetricFlow * mdlConstants.glbmeters3ToLiters * mdlConstants.glbSecondsInHour

End Function

Private Function GainOutputRatio(SysEq As clsSystemEquations, Numerator As Double, Denom As Double, _
                                 Optional HotSalinityOut As Double, Optional ColdSalinityOut, Optional HotMassFlowOut As Double, Optional ColdMassFlowOut As Double) As Double

Dim i As Long
Dim LatHeat As Double
Dim j As Long
Dim Cp_in As Double
Dim Cp_out As Double
Dim T_in As Double   'Hot side input temperature for a hot flow path
Dim T_out As Double  'Hot side output temperature for a Hot flow path
Dim S_in As Double   'Hot side input salinity for a hot flow path
Dim S_out As Double  'Hot side output salinity for a hot flow path
Dim Tref As Double
Dim m_in As Double
Dim m_out As Double
    
Dim Con As clsConnectivity
Dim CVP As clsControlVolumePair
Dim HotPaths() As Long
Dim ColdPaths() As Long
Dim CVHotPaths() As Long
Dim CVColdPaths() As Long
Dim M_out_hot As Double
Dim M_out_cold As Double

Dim Lavg As Double
Dim mM As Double

    Set Con = SysEq.Connectivity
    
    Tref = mdlConstants.glbNISTReferenceTemperature
    
    Con.GetFlowPaths HotPaths, ColdPaths, CVHotPaths, CVColdPaths
    
    Numerator = 0
    For i = 1 To SysEq.Connectivity.NumberOfControlVolumePairs
        Set CVP = SysEq.ControlVolumePair(i)
        Lavg = mdlProperties.LatentHeatOfPureWater(CVP.MembraneAverageTemperature)
        mM = CVP.MassFlow
        Numerator = Numerator + Lavg * mM
    Next
    
    Denom = 0
    HotSalinityOut = 0
    ColdSalinityOut = 0
    HotMassFlowOut = 0
    ColdMassFlowOut = 0
    ' This is not for the gain output ratio but to calculate metrics for the cold
    ' side that have no direct relevance to the GOR.
    For i = 1 To UBound(ColdPaths, 2)
       m_in = SysEq.MassFlow(ColdPaths(1, i))
       m_out = SysEq.MassFlow(ColdPaths(UBound(ColdPaths, 1), i))
       
       S_in = SysEq.Salinity(ColdPaths(1, i))
       S_out = SysEq.Salinity(ColdPaths(UBound(ColdPaths, 1), i))
    
       ColdMassFlowOut = ColdMassFlowOut + m_out
       ColdSalinityOut = ColdSalinityOut + S_out * m_out
    Next i
    ColdSalinityOut = ColdSalinityOut / ColdMassFlowOut
    
    ' This calculates the GOR and some additional metrics for the hot side.
    For i = 1 To UBound(HotPaths, 2)
       T_in = SysEq.Temperature(HotPaths(1, i))
       T_out = SysEq.Temperature(HotPaths(UBound(HotPaths, 1), i))
       
       m_in = SysEq.MassFlow(HotPaths(1, i))
       m_out = SysEq.MassFlow(HotPaths(UBound(HotPaths, 1), i))
       
       S_in = SysEq.Salinity(HotPaths(1, i))
       S_out = SysEq.Salinity(HotPaths(UBound(HotPaths, 1), i))
    
       HotMassFlowOut = HotMassFlowOut + m_out
       HotSalinityOut = HotSalinityOut + S_out * m_out
          
       Cp_in = mdlProperties.SeaWaterSpecificHeat(T_in, S_in)
       Cp_out = mdlProperties.SeaWaterSpecificHeat(T_out, S_out)
       Denom = Denom + (Cp_in * (T_in - Tref) * m_in - Cp_out * (T_out - Tref) * m_out)
    Next i
    HotSalinityOut = HotSalinityOut / HotMassFlowOut
    
    GainOutputRatio = Numerator / Denom

End Function

Private Sub DetermineLayerMultiplicationFactor(SysEq As clsSystemEquations, RunConfig() As Long, MultFact() As Long)
' This sub must be consistent with mdlMain.FigureOutRunSequence
Dim NumRun As Long
Dim Inp As clsInput

Set Inp = SysEq.Inputs

NumRun = UBound(RunConfig)

ReDim MultFact(1 To NumRun)

If NumRun = 1 Then
   MultFact(1) = Inp.NumberOfLayers
ElseIf NumRun = 2 Then ' THIS ONLY HAPPENS WHEN USING INDEPENDENT LAYERS
   If mdlConstants.glbUseIndependentLayers Then
      MultFact(1) = Inp.NumberOfLayers - 1
      MultFact(2) = 1
   Else
      MsgBox "This should never happen!"
   End If
ElseIf NumRun = 3 Then
   If mdlConstants.glbUseIndependentLayers Then
      MultFact(1) = Inp.NumberOfLayers - 2
   Else
      MultFact(1) = 2 * Inp.NumberOfLayers - 3
   End If
   MultFact(2) = 1
   MultFact(3) = 1
End If

End Sub

Public Sub CalculateOutput(SysEqns() As clsSystemEquations, RunConfig() As Long)
   Dim FlowLperH As Double
   Dim HotAvg As Double
   Dim ColdAvg As Double
   Dim CWS As clsWaterStream
   Dim HWS As clsWaterStream
   Dim NumRun As Long
   Dim Numerator As Double
   Dim Denom As Double
   Dim TempNumerator As Double
   Dim TempDenom As Double
   Dim i As Long
   Dim GOR As Double
   Dim MultFact() As Long
   Dim MassFlow As Double
   Dim TempFlowLperH As Double
   Dim HeatFlow As Double
   Dim TempHeatLoss As Double
   Dim HeatLoss As Double
   Dim MassHeat As Double
   Dim TempMassHeat As Double
   Dim DebugString As String
   Dim HotSalinityOut As Double
   Dim ColdSalinityOut As Double
   Dim TempHotSalinityOut As Double
   Dim TempColdSalinityOut As Double
   Dim HotMassFlowOut As Double
   Dim ColdMassFlowOut As Double
   Dim TempHotMassFlowOut As Double
   Dim TempColdMassFlowOut As Double
   'RunConfig - 0 = no losses , 1 = hot losses, 2 = cold losses, 3 = hot and cold losses.
   NumRun = UBound(RunConfig)

   ' The multiplication factor is different than the factor of 1/2 applied to flows (if symmetric run) at the beginning
   ' of the analysis, it is used to multiply numbers if many layers exist.
   DetermineLayerMultiplicationFactor SysEqns(1), RunConfig, MultFact

   Set CWS = SysEqns(1).Inputs.WaterStreams(SysEqns(1).Inputs.ColdWaterStream)
   Set HWS = SysEqns(1).Inputs.WaterStreams(SysEqns(1).Inputs.HotWaterStream)
   
   
   HotSalinityOut = 0
   ColdSalinityOut = 0
   HotMassFlowOut = 0
   ColdMassFlowOut = 0
   ' Gain output ratio
   For i = 1 To NumRun
       GOR = GainOutputRatio(SysEqns(i), TempNumerator, TempDenom, TempHotSalinityOut, TempColdSalinityOut, TempHotMassFlowOut, ColdMassFlowOut)
       HotMassFlowOut = HotMassFlowOut + MultFact(i) * TempHotMassFlowOut
       ColdMassFlowOut = ColdMassFlowOut + MultFact(i) * TempColdMassFlowOut
       
       'Use weighted average to determine mixed salinity out for all layers
       HotSalinityOut = HotSalinityOut + HotMassFlowOut * TempHotSalinityOut
       ColdSalinityOut = ColdSalinityOut + ColdMassFlowOut * TempColdSalinityOut
       
       
       Numerator = Numerator + MultFact(i) * TempNumerator
       Denom = Denom + MultFact(i) * TempDenom
   Next i
   GOR = Numerator / Denom
   HotSalinityOut = HotSalinityOut / HotMassFlowOut
   ColdSalinityOut = ColdSalinityOut / ColdMassFlowOut
   
   Range("GainOutputRatioRange").Value = GOR
   Range("HotSalinityOutRange").Value = HotSalinityOut
   Range("ColdSalinityOutRange").Value = ColdSalinityOut
   Range("HotMassFlowOutRange").Value = HotMassFlowOut
   Range("ColdMassFlowOutRange").Value = ColdMassFlowOut
   
   ' Total MD Mass Flow
   MassFlow = 0
   FlowLperH = 0
   HeatFlow = 0
   HeatLoss = 0
   MassHeat = 0
   For i = 1 To NumRun
      MassFlow = MassFlow + MultFact(i) * TotalMDMassFlow(SysEqns(i), TempFlowLperH)
      FlowLperH = FlowLperH + MultFact(i) * TempFlowLperH
      HeatFlow = HeatFlow + MultFact(i) * TotalHeatFlow(SysEqns(i), TempHeatLoss, TempMassHeat)
      HeatLoss = HeatLoss + TempHeatLoss
      MassHeat = MassHeat + TempMassHeat
   Next i
   
   Range("Total_MD_MassFlow").Value = MassFlow
   Range("TotalMDMassFlowRateInLitersPerHour").Value = FlowLperH
   
   Range("ColdInputTemperatureRange") = CWS.Temperature - mdlConstants.glbCelciusToKelvinOffset
   Range("HotWaterInputTemperatureRange") = HWS.Temperature - mdlConstants.glbCelciusToKelvinOffset
   
   AverageOutputTemperatures SysEqns(1), HotAvg, ColdAvg
   
   Range("AverageHotWaterTemperatureRange") = HotAvg - mdlConstants.glbCelciusToKelvinOffset
   Range("AverageColdOutputTemperatureRange") = ColdAvg - mdlConstants.glbCelciusToKelvinOffset
   
   ' Always use the first run.  These are not equivalent.
   HeatExchangerPerformanceCurves SysEqns(1)
   
   Range("TotalHeatFlowRange") = HeatFlow
   Range("HeatLossRange") = HeatLoss
   Range("HTDueToMassTransferRange") = MassHeat
   
   If glbEvaluateEquationsAtEnd Then
         DebugString = "OUTPUTS:" & vbCrLf
         DebugString = DebugString & FormatDebugColumns("heat_losses: Total Heat Losses (W):      " & HeatLoss & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("heat_from_mass_flow: Heat flow from mass transfer (W):      " & MassHeat & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("heat_from_mass_flow: Heat flow from membrane conduction (W):      " & HeatFlow - MassHeat & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("heat_from_mass_flow: Total Heat flow (W):      " & HeatFlow & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("dist_mass: Total distillate mass flow (kg/s):      " & MassFlow & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("dist_volu: Total distillate volume flow (Liter/hour):      " & FlowLperH & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("GOR: Overall Gain Output Ratio:      " & GOR & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("drain_hot.t:  Hot Outlet temperature (C): " & HotAvg - mdlConstants.glbCelciusToKelvinOffset & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("drain_cold.t: Cold Outlet temperature (C):    " & ColdAvg - mdlConstants.glbCelciusToKelvinOffset & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("drain_hot.m:  Hot Outlet mass flow (kg/s): " & HotMassFlowOut & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("drain_cold.m: Cold Outlet mass flow (kg/s):    " & ColdMassFlowOut & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("drain_hot.w:  Hot Outlet salinity (kg/kg): " & HotSalinityOut & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("drain_cold.w: Cold Outlet salinity (kg/kg):    " & ColdSalinityOut & vbCrLf)
         Debug.Print DebugString
         mdlValidation.WriteDebugInfoToFile DebugString
    End If
   
End Sub

Private Sub AverageOutputTemperatures(SysEq As clsSystemEquations, HotAvg As Double, ColdAvg As Double)

Dim Con As clsConnectivity
Dim HotPaths() As Long
Dim ColdPaths() As Long
Dim CVHotPaths() As Long
Dim CVColdPaths() As Long
Dim HSum As Double 'Sum of hot temperatures
Dim CSum As Double 'Sum of cold temperatures
Dim i As Long

Set Con = SysEq.Connectivity

Con.GetFlowPaths HotPaths, ColdPaths, CVHotPaths, CVColdPaths

HSum = 0
CSum = 0
For i = 1 To UBound(HotPaths, 2)
   HSum = HSum + SysEq.Temperature(HotPaths(UBound(HotPaths, 1), i))
Next i
For i = 1 To UBound(ColdPaths, 2)
   CSum = CSum + SysEq.Temperature(ColdPaths(UBound(ColdPaths, 1), i))
Next i

HotAvg = HSum / UBound(HotPaths, 2)
ColdAvg = CSum / UBound(ColdPaths, 2)

End Sub

Public Sub HeatExchangerPerformanceCurves(SysEq As clsSystemEquations)

Dim HTable As Variant
Dim CTable As Variant
Dim Con As clsConnectivity
Dim HotPaths() As Long
Dim ColdPaths() As Long
Dim CVHotPaths() As Long
Dim CVColdPaths() As Long
Dim HSum As Double 'Sum of hot temperatures
Dim CSum As Double 'Sum of cold temperatures
Dim i As Long
Dim j As Long
Dim HotPathLength As Long
Dim NumHotPaths As Long
Dim ColdPathLength As Long
Dim NumColdPaths As Long
Dim HReversals As Long
Dim CReversals As Long
Dim Ind As Long
Dim Div As Long

Set Con = SysEq.Connectivity

Con.GetFlowPaths HotPaths, ColdPaths, CVHotPaths, CVColdPaths

HotPathLength = UBound(HotPaths, 1)
NumHotPaths = UBound(HotPaths, 2)
ColdPathLength = UBound(ColdPaths, 1)
NumColdPaths = UBound(ColdPaths, 2)
HReversals = Con.HotReversals
CReversals = Con.ColdReversals

Sheet7.Application.Calculation = xlCalculationManual
Range("ClearMeRange").Clear

Range(Range("FlowPathRange"), Range("FlowPathRange").Offset(ColdPathLength - 1 - CReversals, 1)).Value = _
   ConstructHeatExchangerTable(Con.ColdFlowsHorizontal, Con, ColdPathLength, NumColdPaths, ColdPaths, SysEq, CReversals)
Range(Range("HotStreamRangeStart"), Range("HotStreamRangeStart").Offset(HotPathLength - 1 - HReversals, 1)).Value = _
   ConstructHeatExchangerTable(Con.HotFlowsHorizontal, Con, HotPathLength, NumHotPaths, HotPaths, SysEq, HReversals)

Sheet7.Application.Calculation = xlCalculationAutomatic
End Sub

Private Function ConstructHeatExchangerTable(FlowsHorizontal As Boolean, Con As clsConnectivity, PathLength As Long, _
                                             NumPath As Long, Paths() As Long, SysEq As clsSystemEquations, Reversals As Long) As Variant

Dim Div As Long
Dim Sum As Double
Dim Table As Variant
Dim Ind As Long
Dim j As Long
Dim i As Long

If FlowsHorizontal Then
    Div = Con.HorizontalDivisions
Else
    Div = Con.VerticalDivisions
End If

ReDim Table(1 To PathLength - Reversals, 1 To 2)

Ind = 1
For j = 1 To PathLength
   Sum = 0
   For i = 1 To NumPath
      Sum = Sum + SysEq.Temperature(Paths(j, i))
   Next i
   If j Mod (Div + 1) <> 0 Or j = PathLength Then
      Table(Ind, 2) = Sum / NumPath
      Table(Ind, 1) = (Ind - 1) / (PathLength - 1 - Reversals)
      Ind = Ind + 1
   End If
Next j

ConstructHeatExchangerTable = Table

End Function

Private Function TotalHeatFlow(SysEq As clsSystemEquations, Qloss As Double, Qmass As Double)

Dim CV As clsControlVolumePair
Dim i As Long
Dim Q As Double

Q = 0
Qloss = 0
Qmass = 0
For i = 1 To SysEq.Connectivity.NumberOfControlVolumePairs
   ' Order does not matter
   Set CV = SysEq.ControlVolumePair(i)
   Q = Q + CV.HeatTransferLateral
   Qloss = Qloss + CV.HeatLosses
   Qmass = Qmass + CV.HeatTransferDueToMassTransfer
Next i
   
TotalHeatFlow = Q

End Function










