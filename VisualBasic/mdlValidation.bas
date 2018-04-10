Attribute VB_Name = "mdlValidation"
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

Option Explicit
Option Base 1

Public Sub NREL_Sensitivity_Study_1()

' This study is set up for Sertac Akar at NREL. It presupposes the following:

'1. V7_Sub has just been run
'
' Here is the original study summary table. All parameters (Normal length was not varied!!) undergoes
' a range of changes as shown:

'Parameter                           Unit    -50.00% -25.00% -10.00% 0.00%   25.00%  50.00%  100.00%
'Spacer Thickness                    m       0.001   0.0015  0.0018  0.002   0.0025  0.003   0.004
'HTEX Area                           m2      0.0578  0.1300  0.1872  0.23    0.3611  0.5199  0.9244
'HTEX Length Paralel to Hot Flow     m       0.52    0.78    0.936   1.04    1.3     1.56    2.08
'HTEX Length Normal to Hot Flow      m       0.1111  0.16665 0.19998 0.2222  0.27775 0.3333  0.4444
'Feed Mass Flow (Hot Side)           kg/s    0.0082  0.0123  0.0147  0.01638 0.0205  0.0246  0.0328
'Feed Mass Flow (Cold Side)          kg/s    0.0083  0.0124  0.0149  0.0165  0.0207  0.0249  0.0332
'
'Input Temperature (Hot Side)    ?C  50  55  60  65.37   70  80  90
'Input Temperature (Cold Side)   ?C  15  20  25  29.35   30  35  40
'Salt Water Salinity kg/kg   0.001   0.002   0.003   0.004   0.06    0.08    0.016

Dim FSObj As FileSystemObject
Dim ParameterToChange() As String
Dim PercentChange() As Double
Dim SubCaseValue As Variant
Dim i As Long
Dim j As Long
Dim ParamVal() As Double

' Turn off ambient losses
Sheet4.CheckBox_IncludeAmbient.Value = False
' Set to a single cell run comparable to NREL's single control volume model.
Sheet3.Range("Horizontal_Divisions") = 1
Sheet3.Range("Vertical_Divisions") = 1

ReDim ParameterToChange(1 To 7)
' Indicate to the program that this is a custom routine
ParameterToChange(1) = "SpacerThickness"
ParameterToChange(2) = "LengthParallelToFlow"
ParameterToChange(3) = "HotMassFlow"
ParameterToChange(4) = "ColdMassFlow"
ParameterToChange(5) = "HotInputTemperature"
ParameterToChange(6) = "ColdInputTemperature"
ParameterToChange(7) = "HotWaterSalinity"
ReDim PercentChange(1 To 6)
PercentChange(1) = -50#
PercentChange(2) = -25#
PercentChange(3) = -10#
PercentChange(4) = 25#
PercentChange(5) = 50#
PercentChange(6) = 100#
'Input Temperature (Hot Side)    ?C  50  55  60  65.37   70  80  90
'Input Temperature (Cold Side)   ?C  15  20  25  29.35   30  35  40
'Salt Water Salinity kg/kg   0.001   0.002   0.003   0.004   0.06    0.08    0.016
ReDim ParamVal(1 To 3, 1 To 6)
ParamVal(1, 1) = 50#
ParamVal(1, 2) = 55#
ParamVal(1, 3) = 60#
ParamVal(1, 4) = 70#
ParamVal(1, 5) = 80#
ParamVal(1, 6) = 90#

ParamVal(2, 1) = 15#
ParamVal(2, 2) = 20#
ParamVal(2, 3) = 25#
ParamVal(2, 4) = 30#
ParamVal(2, 5) = 35#
ParamVal(2, 6) = 40#

ParamVal(3, 1) = 0.001
ParamVal(3, 2) = 0.002
ParamVal(3, 3) = 0.003
ParamVal(3, 4) = 0.006
ParamVal(3, 5) = 0.008
ParamVal(3, 6) = 0.016

ReDim SubCaseValue(0 To 1)
Set FSObj = New FileSystemObject

' Perform the base case run
mdlMain.MultiConfigurationMembraneDistillationModel

FSObj.CopyFile ThisWorkbook.path & "\" & glbDebugFileName, ThisWorkbook.path & "\NREL_Validation_Baseline.txt"
'
For i = 1 To UBound(ParameterToChange)
   For j = 1 To UBound(PercentChange)
        SubCaseValue(0) = ParameterToChange(i)
        If i <= 4 Then
           SubCaseValue(1) = PercentChange(j)
        Else
           SubCaseValue(1) = ParamVal(i - 4, j)
        End If
        mdlMain.MultiConfigurationMembraneDistillationModel True, "NREL_Validation", SubCaseValue
        If i <= 4 Then
            FSObj.CopyFile ThisWorkbook.path & "\" & glbDebugFileName, _
                           ThisWorkbook.path & "\NREL_Validation_" & ParameterToChange(i) & "_" & _
                           PercentChange(j) & "PercentChangeFromBaseline.txt"
        Else
            FSObj.CopyFile ThisWorkbook.path & "\" & glbDebugFileName, _
                           ThisWorkbook.path & "\NREL_Validation_" & ParameterToChange(i) & "_EqualsTo_" & _
                           SubCaseValue(1) & ".txt"
        End If
   Next j
Next i


End Sub

Public Sub NREL_Parameter_Changes(SysEqn As clsSystemEquations, SubCaseValue As Variant)

Dim ChgFactor As Double
Dim MFlow As Double
Dim Thick As Double
Dim ITemp As Double
Dim ISalinity As Double

ChgFactor = (1 + SubCaseValue(1) / 100)

' For this custom subroutine, SubCaseValue contains two entries 1. Key word 2. KeyValue
Select Case CStr(SubCaseValue(0))
    Case "SpacerThickness"
       ' Thickness comes out as meters but must be input as milimeters
        Thick = SysEqn.Inputs.Spacers(SysEqn.Inputs.HotSpacer).Thickness
        SysEqn.Inputs.Spacers(SysEqn.Inputs.HotSpacer).Thickness = Thick * ChgFactor / glbMiliMeterToMeter
        If SysEqn.Inputs.ColdSpacer <> SysEqn.Inputs.HotSpacer Then
           Thick = SysEqn.Inputs.Spacers(SysEqn.Inputs.ColdSpacer).Thickness
           SysEqn.Inputs.Spacers(SysEqn.Inputs.ColdSpacer).Thickness = Thick * ChgFactor / glbMiliMeterToMeter
        End If
    Case "LengthParallelToFlow"
        SysEqn.Inputs.VerticalLength = SysEqn.Inputs.VerticalLength * ChgFactor
    Case "HotMassFlow"
        MFlow = SysEqn.Inputs.WaterStreams(SysEqn.Inputs.HotWaterStream).MassFlow
        SysEqn.Inputs.WaterStreams(SysEqn.Inputs.HotWaterStream).MassFlow = MFlow * ChgFactor
    Case "ColdMassFlow"
        MFlow = SysEqn.Inputs.WaterStreams(SysEqn.Inputs.ColdWaterStream).MassFlow
        SysEqn.Inputs.WaterStreams(SysEqn.Inputs.ColdWaterStream).MassFlow = MFlow * ChgFactor
    Case "HotInputTemperature"
    'Must be in Celcius
    ' First get the mass flow before the temperature change. NREL's model holds mass flow constant
    ' instead of volumetric flow rate. Then change the temperature, then adjust the mass flow to the
    ' original value.
        MFlow = SysEqn.Inputs.WaterStreams(SysEqn.Inputs.HotWaterStream).MassFlow
        SysEqn.Inputs.WaterStreams(SysEqn.Inputs.HotWaterStream).Temperature = SubCaseValue(1) 'This varied the mass flow while hold volume flow constant
        SysEqn.Inputs.WaterStreams(SysEqn.Inputs.HotWaterStream).MassFlow = MFlow
    Case "ColdInputTemperature"
    ' Must be in celcius
        MFlow = SysEqn.Inputs.WaterStreams(SysEqn.Inputs.ColdWaterStream).MassFlow
        SysEqn.Inputs.WaterStreams(SysEqn.Inputs.ColdWaterStream).Temperature = SubCaseValue(1)
        SysEqn.Inputs.WaterStreams(SysEqn.Inputs.ColdWaterStream).MassFlow = MFlow
    Case "HotWaterSalinity"
       ' input must be in g/kg
       ' first collect mass flow otherwise mass flow will vary
       MFlow = SysEqn.Inputs.WaterStreams(SysEqn.Inputs.HotWaterStream).MassFlow
       SysEqn.Inputs.WaterStreams(SysEqn.Inputs.HotWaterStream).Salinity = SubCaseValue(1) / glbGramPerKilogramToFraction
       SysEqn.Inputs.WaterStreams(SysEqn.Inputs.HotWaterStream).MassFlow = MFlow
    Case Else
        Debug.Assert False
End Select


End Sub

Public Function IsFinalRun() As Boolean
    IsFinalRun = CBool(Range("IsFinalRunRange").Value)
End Function

Public Function IsFinalNewtonIteration() As Boolean

   IsFinalNewtonIteration = CBool(Range("IsFinalNewtonIterationRange").Value)

End Function

' This is used by direct contact, air gap will have a more complex adjustment scheme
Public Sub Dummy_AdjustInterfaceTemperaturesIfInvalid(T_avgh As Double, T_avgc As Double, T_ih As Double, T_ic As Double)
' THIS IS THE SAME AS THE FUNCTION ADDED TO clsHTFuncDirectContact And clsHTFuncAirGap!!!!
' This function is used to adjust the interface temperatures if they
' have physically inadmissible values after adjustment of the global
' equation set. A large jump in the global solution may make the previous
' interface temperatures completely incorrect. This procedure has
' made the model much more robust for heat exchangers that come close
' to saturation.
Dim Dum As Double
Dim Tdif As Double
Dim Frac_h As Double
Dim Frac_c As Double
Dim AdjustmentMade As Boolean
Const Big = 0.2
Const Small = 0.05

AdjustmentMade = False
'The following order of temperatures is required for the 2nd law of thermodynamics to be
'satisfied (from hottest to coldest)
' 1. T_avgh
' 2. T_ih
' 3. T_ic
' 4. T_avgc

' If this is not the case, force it!!!!

' Test cases (keep commented out unless you change this function)!

If T_avgh <= T_avgc Then
   Dum = mdlError.ReturnError("clsHTFuncDirectContact.AdjustInterfaceTemperaturesIfInvalid: The average hot temperature " & _
                              " and average cold temperature of the global solution are inverted! no way to fix this at the local level!", , True, True)
    
Else
   Tdif = T_avgh - T_avgc
   'First assire that T_ic and T_ih are inbetween T_avgh and T_avgc
   ' The entire objective is to return to a physically admissible solution
   If T_ih >= T_avgh Then
      T_ih = T_avgh - Small * Tdif
      AdjustmentMade = True
   End If
   If T_ic >= T_avgh Then
      T_ic = T_avgh - Big * Tdif
      AdjustmentMade = True
   End If
   If T_ih <= T_avgc Then
      T_ih = T_avgc + Big * Tdif
      AdjustmentMade = True
   End If
   If T_ic <= T_avgc Then
      T_ic = T_avgc + Small * Tdif
      AdjustmentMade = True
   End If
   
   ' Now assure that T_ic and T_ih are not reversed, move back the temperature
   ' that is encroaching closer to the average temperatures
   If T_ic >= T_ih Then
      ' Calculate the fraction distance
      AdjustmentMade = True
      Frac_h = (T_avgh - T_ih) / Tdif
      Frac_c = (T_ic - T_avgc) / Tdif
      
      If Frac_h > Frac_c Then 'move the hot temperature up
          T_ih = T_ic + Small * (T_avgh - T_ic)
      Else ' move the cold temperature down
          T_ic = T_ih - Small * (T_ih - T_avgc)
      End If
   End If
End If

End Sub

Public Function FormatDebugColumns(Str As String) As String

Dim StrArr() As String
Dim i As Long
Dim Dum As Double
Dim Out As String
Dim ColNS() As Long

CStrArr Split(Str, ":"), StrArr
Dim LnStr As Long
Dim TStr As String
ReDim ColNS(0 To 2)

ColNS(0) = 20
ColNS(1) = 60
ColNS(2) = 22

If NumElements(StrArr) <> 3 Then
   Dum = mdlError.ReturnError("mdlValidation.FormatColumns: This function only works for a three column case!", True, True)
End If

Out = ""
For i = 0 To UBound(StrArr)
    TStr = Trim(StrArr(i))
    LnStr = Len(TStr)
    If LnStr > ColNS(i) Then
       Dum = mdlError.ReturnError("mdlValidation.FormatColumns: One of the debug strings is longer than the designated column width!", , True, True)
    End If
    If i <> UBound(StrArr) Then
       Out = Out & TStr & "," & Space(ColNS(i) - LnStr)
    Else
       Out = Out & TStr
    End If

Next i

FormatDebugColumns = Out

End Function

Public Sub WriteDebugInfoToFile(StringToWrite As String, Optional DeleteFile As Boolean = False)

Dim FSObj As FileSystemObject
Dim TS As TextStream
Dim FileName As String
Dim Dum As Double

FileName = ThisWorkbook.path & "\" & glbDebugFileName

Set FSObj = New FileSystemObject

If DeleteFile Then
    If FSObj.FileExists(FileName) Then
       On Error GoTo DeleteFailed
       FSObj.DeleteFile (FileName)
       GoTo DeleteSucceeded
DeleteFailed:
       Dum = mdlError.ReturnError("mdlValidation.WriteDebugInfoToFile: The debug info file, """ & FileName & """ could not be deleted or written to. Another application may be using it!", , True, True)
DeleteSucceeded:
       On Error GoTo 0
    End If
End If

Set TS = FSObj.OpenTextFile(FileName, ForAppending, True)

TS.Write (StringToWrite)
TS.Close
Set FSObj = Nothing
Set TS = Nothing

End Sub

'This function was tested by dlvilla 6/30/2016
Public Function AssignStringIfInObjectCollection(Str As String, Col As Collection, _
                                         ObjectTypeStr As String, ModuleStr As String, SheetName As String, ByRef ErrMsg As String) As String
    Dim Obj As Object
    
    'This function ONLY works on collections of objects.  Collections of other types will fail!!!
    
    ErrMsg = ""
    
    If Len(Str) = 0 Then
        ErrMsg = ObjectTypeStr & " name must be non-blank and unique!"
    Else
        On Error GoTo InvalidName
        Set Obj = Col.Item(Str)
    End If
EndOfFunction:

If Len(ErrMsg) = 0 Then
    AssignStringIfInObjectCollection = Str
Else
    AssignStringIfInObjectCollection = ""
End If
    
Exit Function
InvalidName:
   ErrMsg = ObjectTypeStr & " for the " & ModuleStr & " is not a valid name found in the """ & SheetName & """ sheet."
   GoTo EndOfFunction
End Function

'This function was tested by dlvilla 6/30/2016
Public Function AssignValueIfInLimits(Var As Variant, ErrMsg As String, VarName As String, INVALID_VALUE As Variant, Unit As String, Optional LowerLimit As Variant, Optional UpperLimit As Variant) As Variant

Dim BeginStr As String
Dim UnitStr As String
Dim OrgErrMsg As String

If Len(ErrMsg) <> 0 Then
   OrgErrMsg = ErrMsg
   ErrMsg = ""
End If

BeginStr = VarName & " = " & CStr(Var) & " is out of its valid range which must be "
If Len(Unit) = 0 Then
    UnitStr = "."
Else
    UnitStr = " (" & Unit & ")."
End If

If IsMissing(UpperLimit) And IsMissing(LowerLimit) Then
    'do nothing
ElseIf IsMissing(UpperLimit) Then
    If Var < LowerLimit Then
        ErrMsg = BeginStr & "greater than " & CStr(LowerLimit) & UnitStr
    End If
ElseIf IsMissing(LowerLimit) Then
    If Var > UpperLimit Then
        ErrMsg = BeginStr & "less than " & CStr(UpperLimit) & UnitStr
    End If
Else
    If Var > UpperLimit Or Var < LowerLimit Then
        ErrMsg = BeginStr & "between " & CStr(LowerLimit) & " and " & CStr(UpperLimit) & UnitStr
    End If
End If

If Len(OrgErrMsg) <> 0 And Len(ErrMsg) <> 0 Then
   ErrMsg = OrgErrMsg & vbCrLf & vbCrLf & ErrMsg
ElseIf Len(OrgErrMsg) <> 0 Then
   ErrMsg = OrgErrMsg
End If

If Len(ErrMsg) = 0 Then
    AssignValueIfInLimits = Var
Else
    AssignValueIfInLimits = INVALID_VALUE
End If

End Function

Public Sub CSM_Validation(SheetName As String, RangeNamePrefix As String, IsCounterFlow As Boolean, NumDiscretization As Long, _
                          ColdWaterStream As String, _
                          ColdSpacer As String, _
                          HotWaterStream As String, _
                          HotSpacer As String, _
                          MembraneMaterial As String, _
                          FoilMaterial As String, _
                          AirGapSpacer As String, _
                          MD_Type As String, _
                          HorizontalLength As Double, _
                          VerticalLength As Double, _
                          ExternalMaterial As String, _
                          AmbientTemperature As Double, _
                          Optional NumberOfLayers As Double = 1, _
                          Optional AirGapPressure As Double = 1.0325)

' This is a generalized function that performs validation against the CSM

' To get this work, the worksheet must have the following data:
'1. Distillate Temperature Td,in, deg C and corresponding range RangeNamePrefix & "_DistillateTempIn"
'2. Feed Temperature       Tf, in,deg C and corresponding range RangeNamePrefix & "_FeedTempIn"
'3. In the WaterStreams worksheet ranges
'4. For output of the results: RangeNamePrefix & "_Model_FeedTempOut"
'                               RangeNamePrefix & "_Model_DistillateTempOut"
'                               RangeNamePrefix & "_Model_MDMassOut"

Dim FeedTempIn As Variant
Dim i As Long
Dim Wksh As Worksheet
Dim Dum As Double
Dim FeedTempOut As Variant
Dim DistillateTempOut As Variant
Dim DistillateTempIn As Variant
Dim MDMassFlow As Variant
Dim HL As Variant
Dim VL As Variant
Dim Area As Double
Dim PM As ProgressMeter
Dim Cmb As ComboBox
Dim YesNo As Long
Dim BeginRange As Long
Dim NumFound As Long
Dim GOR_Arr As Variant
Dim QTotal As Variant
Dim TotalArea

YesNo = MsgBox("This study resets settings automatically to those of the validation study. Are you sure you want to do this?", vbYesNo, "Settings will change!")

Select Case YesNo
   Case vbYes

        ' Set all of the inputs automatically

        'Configuration -
        If IsCounterFlow Then
           Range("Hot_Inflow_Edge").Value = 3
        Else
           Range("Hot_Inflow_Edge").Value = 1
        End If
        Range("Hot_Inflow_Edge_Side").Value = 0
        Range("Hot_Reversals").Value = 0
        Range("Cold_Inflow_Edge").Value = 1
        Range("Cold_Inflow_Edge_Side").Value = 0
        Range("Cold_Reversals").Value = 0
        'Control volume discretization
        Range("Horizontal_Divisions").Value = 1
        Range("Vertical_Divisions").Value = NumDiscretization
        Range("AirGapPressure").Value = AirGapPressure
        
        'Input sheet - must choose valid names!
        Sheet4.Combo_ColdWaterStreams.Value = ColdWaterStream
        Sheet4.Combo_ColdSpacer.Value = ColdSpacer
        Sheet4.Combo_HotWaterStream = HotWaterStream
        Sheet4.Combo_HotSpacer.Value = HotSpacer
        Sheet4.Combo_MembraneMaterial = MembraneMaterial
        Sheet4.Combo_MD_Type = MD_Type
        Sheet4.Combo_FoilMaterial = FoilMaterial
        Sheet4.Combo_AirGapSpacer = AirGapSpacer
        Sheet4.Range("NumberOfLayers").Value = NumberOfLayers
        
        Range("Horizontal_Length").Value = HorizontalLength
        Range("Vertical_Length").Value = VerticalLength

        Sheet4.CheckBox_IncludeAmbient = True
        Sheet4.Combo_ExternalMaterial = ExternalMaterial
        Sheet4.Range("AmbientTemperatureRange") = AmbientTemperature
        Sheet4.OptionButton2 = True 'Always 2 sides exposed for this case.
        Range("RangeGravityDirection").Value = 4 ' to the right.  This requires more discretizations
        If MD_Type = "Air Gap" Then ' We need to have some integration of the growth of the distillate 3 elements is enough.
           Range("Horizontal_Divisions") = 3
        Else
           Range("Horizontal_Divisions") = 1
        End If
        
        HL = Range("Horizontal_Length")
        VL = Range("Vertical_Length")
        Area = HL * VL
        
        Set Wksh = ThisWorkbook.Worksheets("WaterStreams")
        FeedTempIn = Range(RangeNamePrefix & "_FeedTempIn")
        DistillateTempIn = Range(RangeNamePrefix & "_DistillateTempIn")
        
        ReDim FeedTempOut(1 To UBound(FeedTempIn, 1), 1 To 1)
        ReDim DistillateTempOut(1 To UBound(FeedTempIn, 1), 1 To 1)
        ReDim MDMassFlow(1 To UBound(FeedTempIn, 1), 1 To 1)
        ReDim GOR_Arr(1 To UBound(FeedTempIn, 1), 1 To 1)
        ReDim QTotal(1 To UBound(FeedTempIn, 1), 1 To 1)
        
        Set PM = New ProgressMeter
        
        PM.NumberOfSteps = UBound(FeedTempIn, 1)
        PM.TargetFraction = 1
        PM.Show vbModeless
        
        For i = 1 To UBound(FeedTempIn, 1)
           
           ' This will not work if you change the order of the Water Streams!!! !@#$
           BeginRange = 4
           NumFound = 0
           Do While Len(Wksh.Range("A" & BeginRange).Value) <> 0
              If Wksh.Range("A" & BeginRange).Value = HotWaterStream Then
                 Wksh.Range("B" & BeginRange).Value = FeedTempIn(i, 1)
                 NumFound = NumFound + 1
              ElseIf Wksh.Range("A" & BeginRange).Value = ColdWaterStream Then
                 Wksh.Range("B" & BeginRange).Value = DistillateTempIn(i, 1)
                 NumFound = NumFound + 1
              End If
              If NumFound = 2 Then
                 Exit Do
              End If
              BeginRange = BeginRange + 1
           Loop
           If NumFound <> 2 Then
              DistillateTempIn(i, 1) = mdlError.ReturnError("mdlValidation.CSM_Validation: Either the hot or cold water stream indicated was not found and could" & _
                                   " therefore not be changed.  The validation study has failed!", , True, True)
           End If
           

           PM.Step "Performing analysis " & CStr(i) & " of " & CStr(UBound(FeedTempIn, 1)) & "."
           If PM.CancelPressed Then
              GoTo EndOfSub
           End If
           
           mdlMain.MultiConfigurationMembraneDistillationModel
           
           If mdlConstants.glbUseIndependentLayers Then
              TotalArea = Area * NumberOfLayers
           Else
              TotalArea = Area * (2 * NumberOfLayers - 1)
           End If
           
           'Now collect the results
           MDMassFlow(i, 1) = Range("Total_MD_MassFlow") / TotalArea ' normalize the result to area (even though the heat exchange does NOT scale unless you have parrallel units)
           DistillateTempOut(i, 1) = Range("AverageColdOutputTemperatureRange")
           FeedTempOut(i, 1) = Range("AverageHotWaterTemperatureRange")
           GOR_Arr(i, 1) = Range("GainOutputRatioRange")
           QTotal(i, 1) = Range("TotalHeatFlowRange") / TotalArea
        Next
        
        Range(RangeNamePrefix & "_Model_FeedTempOut") = FeedTempOut
        Range(RangeNamePrefix & "_Model_DistillateTempOut") = DistillateTempOut
        Range(RangeNamePrefix & "_Model_MDMassOut") = MDMassFlow
        Range(RangeNamePrefix & "_QTotal") = QTotal
        Range(RangeNamePrefix & "_GOR") = GOR_Arr
EndOfSub:
        If PM.CancelPressed Then
           MsgBox "None of the results have been updated.  You will have to rerun the entire analysis"
        End If
        PM.DoNotAllowClose = False
        Unload PM
        Set Wksh = ThisWorkbook.Worksheets(SheetName)
        
        Wksh.Activate

   Case vbNo
      ' Do nothing
End Select
End Sub

Public Sub VAll_Sub()

V1_Sub
V2_Sub
V3_Sub
V4_Sub
V5_Sub
V6_Sub
V7_Sub
V8_Sub

End Sub

Public Sub V1_Sub()

        CSM_Validation SheetName:="V1", _
                       RangeNamePrefix:="CSM_CounterCurrentDCMD", _
                       IsCounterFlow:=True, _
                       NumDiscretization:=5, _
                       ColdWaterStream:="CSM_Distillate20", _
                       ColdSpacer:="CSM_Spacer1", _
                       HotWaterStream:="CSM_Feed40", _
                       HotSpacer:="CSM_Spacer1", _
                       MembraneMaterial:="CLARCOR_QL822_PTFE", _
                       FoilMaterial:="Stainless_Steel_304", _
                       AirGapSpacer:="CSM_AGSpacer", _
                       MD_Type:="Direct Contact", _
                       HorizontalLength:=0.222, _
                       VerticalLength:=1.04, _
                       ExternalMaterial:="DELRIN150_OneInch", _
                       AmbientTemperature:=22
                       

End Sub

Public Sub V2_Sub()

        CSM_Validation SheetName:="V2", _
                       RangeNamePrefix:="CSM_CounterCurrentDCMD", _
                       IsCounterFlow:=False, _
                       NumDiscretization:=5, _
                       ColdWaterStream:="CSM_Distillate20", _
                       ColdSpacer:="CSM_Spacer1", _
                       HotWaterStream:="CSM_Feed40", _
                       HotSpacer:="CSM_Spacer1", _
                       MembraneMaterial:="CLARCOR_QL822_PTFE", _
                       FoilMaterial:="Stainless_Steel_304", _
                       AirGapSpacer:="CSM_AGSpacer", _
                       MD_Type:="Direct Contact", _
                       HorizontalLength:=0.222, _
                       VerticalLength:=1.04, _
                       ExternalMaterial:="DELRIN150_OneInch", _
                       AmbientTemperature:=22
                       'DO NOT WORRY ABOUT CSM_CounterCurrentDCMD it is automatically duplicated
                       'NOTE IsCounterFlow:=False, this is the only change from V1_Sub

End Sub

Public Sub V3_Sub()

        CSM_Validation SheetName:="V3", _
                       RangeNamePrefix:="CSM_CounterCurrentAGMD", _
                       IsCounterFlow:=True, _
                       NumDiscretization:=3, _
                       ColdWaterStream:="CSM_Distillate20", _
                       ColdSpacer:="CSM_AGSpacer", _
                       HotWaterStream:="CSM_Feed40", _
                       HotSpacer:="CSM_AGSpacer", _
                       MembraneMaterial:="CLARCOR_QL822_PTFE", _
                       FoilMaterial:="Stainless_Steel_304", _
                       AirGapSpacer:="CSM_AGSpacer", _
                       MD_Type:="Air Gap", _
                       HorizontalLength:=0.135, _
                       VerticalLength:=1.04, _
                       ExternalMaterial:="DELRIN150_OneInch", _
                       AmbientTemperature:=22
          ' Channel width is listed as 0.135 even though Spacer width is 0.222. This is because
          ' the assembly is on its side and does not fill the entire channel for the flow
          ' rate of 1.5LPM.  THIS IS GOING TO BE FIXED IN THE UPDATED VERSION.

End Sub

Public Sub V4_Sub()

        CSM_Validation SheetName:="V4", _
                       RangeNamePrefix:="CSM_CounterCurrentDCMD", _
                       IsCounterFlow:=True, _
                       NumDiscretization:=5, _
                       ColdWaterStream:="CSM_AquaStill_Spiral_D20", _
                       ColdSpacer:="CSM_Aquastill_Spiral_Spacer", _
                       HotWaterStream:="CSM_AquaStill_Spiral_F40", _
                       HotSpacer:="CSM_Aquastill_Spiral_Spacer", _
                       MembraneMaterial:="Aquastill_PE", _
                       FoilMaterial:="Stainless_Steel_304", _
                       AirGapSpacer:="CSM_AGSpacer", _
                       MD_Type:="Direct Contact", _
                       HorizontalLength:=0.4, _
                       VerticalLength:=0.5, _
                       ExternalMaterial:="DELRIN150_OneInch", _
                       AmbientTemperature:=22, _
                       NumberOfLayers:=9
                       
                       'Foil material and AirGapSpacer Do not matter for this run!
                       

End Sub

Public Sub V5_Sub()

        CSM_Validation SheetName:="V5", _
                       RangeNamePrefix:="CSM_CounterCurrentDCMD", _
                       IsCounterFlow:=True, _
                       NumDiscretization:=5, _
                       ColdWaterStream:="CSM_AquaStill_2LPM_D20", _
                       ColdSpacer:="CSM_Spacer1", _
                       HotWaterStream:="CSM_AquaStill_2LPM_F40", _
                       HotSpacer:="CSM_Spacer1", _
                       MembraneMaterial:="Aquastill_PE", _
                       FoilMaterial:="Stainless_Steel_304", _
                       AirGapSpacer:="CSM_AGSpacer", _
                       MD_Type:="Direct Contact", _
                       HorizontalLength:=0.222, _
                       VerticalLength:=1.04, _
                       ExternalMaterial:="DELRIN150_OneInch", _
                       AmbientTemperature:=22, _
                       NumberOfLayers:=2
                       
                       'Foil material and AirGapSpacer Do not matter for this run!
                       

End Sub

Public Sub V6_Sub()

        CSM_Validation SheetName:="V6", _
                       RangeNamePrefix:="CSM_CounterCurrentDCMD", _
                       IsCounterFlow:=True, _
                       NumDiscretization:=5, _
                       ColdWaterStream:="CSM_AquaStill_1LPM_D20", _
                       ColdSpacer:="CSM_Spacer1", _
                       HotWaterStream:="CSM_AquaStill_1LPM_F40", _
                       HotSpacer:="CSM_Spacer1", _
                       MembraneMaterial:="Aquastill_PE", _
                       FoilMaterial:="Stainless_Steel_304", _
                       AirGapSpacer:="CSM_AGSpacer", _
                       MD_Type:="Direct Contact", _
                       HorizontalLength:=0.222, _
                       VerticalLength:=1.04, _
                       ExternalMaterial:="DELRIN150_OneInch", _
                       AmbientTemperature:=22, _
                       NumberOfLayers:=1
                       
                       'Foil material and AirGapSpacer Do not matter for this run!
                       ' This is different from V5 only in the number of layers!

End Sub

Public Sub V7_Sub()

        CSM_Validation SheetName:="V7", _
                       RangeNamePrefix:="CSM_CounterCurrentDCMD", _
                       IsCounterFlow:=True, _
                       NumDiscretization:=5, _
                       ColdWaterStream:="CSM_AquaStill_1LPM_D20", _
                       ColdSpacer:="CSM_Spacer1", _
                       HotWaterStream:="CSM_AquaStill_1LPM_F40", _
                       HotSpacer:="CSM_Spacer1", _
                       MembraneMaterial:="3M_PolyPropylene", _
                       FoilMaterial:="Stainless_Steel_304", _
                       AirGapSpacer:="CSM_AGSpacer", _
                       MD_Type:="Direct Contact", _
                       HorizontalLength:=0.222, _
                       VerticalLength:=1.04, _
                       ExternalMaterial:="DELRIN150_OneInch", _
                       AmbientTemperature:=22, _
                       NumberOfLayers:=1
                       
                       'Foil material and AirGapSpacer Do not matter for this run!
                       ' Same as V6 but with a 3M polypropylene membrane!

End Sub

Public Sub V8_Sub()

        CSM_Validation SheetName:="V8", _
                       RangeNamePrefix:="CSM_CounterCurrentDCMD", _
                       IsCounterFlow:=True, _
                       NumDiscretization:=5, _
                       ColdWaterStream:="CSM_AquaStill_1LPM_D20", _
                       ColdSpacer:="CSM_Spacer1", _
                       HotWaterStream:="CSM_AquaStill_1LPM_F40", _
                       HotSpacer:="CSM_Spacer1", _
                       MembraneMaterial:="Aquastill_PE", _
                       FoilMaterial:="Stainless_Steel_304", _
                       AirGapSpacer:="CSM_AGSpacer", _
                       MD_Type:="Air Gap", _
                       HorizontalLength:=0.222, _
                       VerticalLength:=1.04, _
                       ExternalMaterial:="DELRIN150_OneInch", _
                       AmbientTemperature:=22
          ' Channel width is listed as 0.135 even though Spacer width is 0.222. This is because
          ' the assembly is on its side and does not fill the entire channel for the flow
          ' rate of 1.5LPM.

End Sub

Public Sub V11_Sub()

        CSM_Validation SheetName:="V11", _
                       RangeNamePrefix:="CSM_CounterCurrentDCMD", _
                       IsCounterFlow:=True, _
                       NumDiscretization:=5, _
                       ColdWaterStream:="CSM_AquaStill_1LPM_D20", _
                       ColdSpacer:="CSM_Spacer1", _
                       HotWaterStream:="CSM_AquaStill_1LPM_F40", _
                       HotSpacer:="CSM_Spacer1", _
                       MembraneMaterial:="3M_PolyPropylene", _
                       FoilMaterial:="Stainless_Steel_304", _
                       AirGapSpacer:="CSM_AGSpacer", _
                       MD_Type:="Air Gap", _
                       HorizontalLength:=0.222, _
                       VerticalLength:=1.04, _
                       ExternalMaterial:="DELRIN150_OneInch", _
                       AmbientTemperature:=22, _
                       AirGapPressure:=1.0325 - 0.1378951 '-2psi vacuum being applied to the air gap.
          ' Channel width is listed as 0.135 even though Spacer width is 0.222. This is because
          ' the assembly is on its side and does not fill the entire channel for the flow
          ' rate of 1.5LPM.

End Sub

Public Sub ModelSensitivityStudy()
   Dim Ind As Long
   Dim SensStudy As String
   mdlConstants.GlobalArrays
   Dim Low As Double
   Dim High As Double
   Dim Div As Long
   Dim Val As Double
   Dim Settings As Variant
   Dim AllRes As Variant
   Dim i As Long
   Dim j As Long
   Dim Rng As Range
   Dim PM As ProgressMeter
   
   SensStudy = Sheet23.ComboBox_ModelParameterToRun.Value
   
   Ind = FindIndex(SensStudy, mdlConstants.glbSensitivityStudies)
   
   Low = mdlConstants.glbSensitivityStudyLowValue(Ind)
   High = mdlConstants.glbSensitivityStudyHighValue(Ind)
   Div = mdlConstants.glbSensitivityStudyNumDiv(Ind)
   
   Settings = ReturnAllSettings
   
   Set Rng = Range(Range("SettingsStartPoint"), Range("SettingsStartPoint").Offset(UBound(Settings, 1) - 1, UBound(Settings, 2) - 1))
   Rng.Value = Settings
   
           ' Set up the progress meter
        Set PM = New ProgressMeter
        PM.NumberOfSteps = Ceiling(Div)
        PM.TargetFraction = 1
        PM.Show vbModeless
   
   For i = 1 To Div
           PM.Step "Performing analysis " & CStr(i) & " of " & CStr(PM.NumberOfSteps) & "."
           If PM.CancelPressed Then
              GoTo EndOfSub
           End If
   
      Val = Low + (i - 1) * (High - Low) / (Div - 1)
      Range(mdlConstants.glbSensitivityStudyAlterRange(Ind)).Value = Val
      ' Holding all other settings constant!  YOU NEED TO MAKE A GATHERINPUTS FUNCTION THAT GATHERS ALL OF THE INPUTS!!
      Range("SensitivityStartPointRange").Offset(i - 1, 0) = Val
      mdlMain.MultiConfigurationMembraneDistillationModel
      AllRes = Range("AllResults")
      For j = 1 To UBound(AllRes)
        Range("SensitivityStartPointRange").Offset(i - 1, j).Value = AllRes(j, 1)
      Next
      
      
   Next
   

EndOfSub:
        On Error Resume Next
        PM.DoNotAllowClose = False
        Unload PM
End Sub

Public Sub VerticalDivisionNumericalConvergenceStudy()

' Provides numerical convergence of the current configuration of the model
' for vertical divisions
Dim ConvergenceCriterion As Double
Dim ElementStep As Long
Dim MaxElements As Long
Dim NotConverged As Boolean
Dim NumElements As Long
Dim Settings As Variant
Dim PM As ProgressMeter
Dim Results As Variant
Dim Offset As Long
Dim YesNo As Long
Dim FirstTime As Boolean
Dim OldResults As Variant
Dim StudyDone As Boolean
Dim WroteStudyDone As Boolean

ReDim Results(11, 1)


YesNo = MsgBox("This study uses the current model's settings and takes awhile to run.  Are you sure you want to run it?", vbYesNo, "Proceed?")



Select Case YesNo
   Case vbYes

        MaxElements = CLng(Range("NumConvergMaxNumElement"))
        ElementStep = CLng(Range("NumConvergStepSize").Value)
        ConvergenceCriterion = Range("NumConvegConvergeCriterion")  'Keep on until 0.1% change has been obtained
        NotConverged = True
        NumElements = ElementStep
        
        
        
        ' Set up the progress meter
        Set PM = New ProgressMeter
        PM.NumberOfSteps = Ceiling(MaxElements / ElementStep)
        PM.TargetFraction = 1
        PM.Show vbModeless
        
        ' Record all of the settings for the current run (to assure the study results are documented.
        Settings = ReturnAllSettings
        
        Range(Range("NumericalConvergenceSettingsStartPointRange"), _
              Range("NumericalConvergenceSettingsStartPointRange").Offset(UBound(Settings, 1) - 1, UBound(Settings, 2) - 1)) = _
             Settings
        Offset = 0
        FirstTime = True
        WroteStudyDone = False
        Do While NotConverged And NumElements <= MaxElements
            Range("Vertical_Divisions") = NumElements
            
           PM.Step "Performing analysis " & CStr(NumElements / ElementStep) & " of " & CStr(PM.NumberOfSteps) & "."
           If PM.CancelPressed Then
              GoTo EndOfSub
           End If
            
           mdlMain.MultiConfigurationMembraneDistillationModel
           
           Results = Range("AllResults")
           
           
           'Write the study incrementally so that progress can be made
           If FirstTime Then
                Range(Range("NumericalConvergenceStartPoint"), Range("NumericalConvergenceStartPoint").Offset(UBound(Results, 1) - 1, PM.NumberOfSteps)).Clear
                Range(Range("Numconverg_MaxDiff_Start").Offset(0, 0), Range("Numconverg_MaxDiff_Start").Offset(0, PM.NumberOfSteps)).Clear
           End If
           Range(Range("NumericalConvergenceStartPoint").Offset(0, Offset), Range("NumericalConvergenceStartPoint").Offset(UBound(Results, 1) - 1, Offset)) = _
            Results
        
            ' Check to see what the convergence criterion is
            If Not FirstTime Then
               Range("Numconverg_MaxDiff_Start").Offset(0, Offset) = MaximumPercentChangeOfAllResults(Results, OldResults, ConvergenceCriterion, StudyDone)
               If StudyDone And Not WroteStudyDone Then
                   Range("NumElementThatMeetsCriterion") = NumElements
                   WroteStudyDone = True
               End If
            End If
            
            OldResults = Results
            NumElements = NumElements + ElementStep
            Offset = Offset + 1
            FirstTime = False
               
        Loop
        
    Case vbNo
        'Nothing to do
    End Select
EndOfSub:
        On Error Resume Next
        PM.DoNotAllowClose = False
        Unload PM

End Sub



Public Function ReturnAllSettings() As Variant

Dim V As Variant

ReDim V(1 To 26, 1 To 2)
' Write the current Settings to Range
V(1, 2) = Range("Hot_Inflow_Edge").Value
V(1, 1) = "Hot Inflow Edge"
V(2, 2) = Range("Hot_Inflow_Edge_Side").Value
V(2, 1) = "Hot Inflow Edge Side"
V(3, 2) = Range("Hot_Reversals").Value
V(3, 1) = "Hot Reversals"
V(4, 2) = Range("Cold_Inflow_Edge").Value
V(4, 1) = "Cold Inflow Edge"
V(5, 2) = Range("Cold_Inflow_Edge_Side").Value
V(5, 1) = "Cold Inflow Edge Side"
V(6, 2) = Range("Cold_Reversals").Value
V(6, 1) = "Cold Reversals"
'Control volume discretization
V(7, 2) = Range("Horizontal_Divisions").Value
V(7, 1) = "Horizontal Divisions"
V(8, 2) = Range("Vertical_Divisions").Value
V(8, 1) = "Vertical Divisions (Variable In this study)"

'Input sheet - must choose valid names!
V(9, 2) = Sheet4.Combo_ColdWaterStreams.Value
V(9, 1) = "Cold Water Stream"
V(10, 2) = Sheet4.Combo_ColdSpacer.Value
V(10, 1) = "Cold Spacer"
V(11, 2) = Sheet4.Combo_HotWaterStream
V(11, 1) = "Hot Water Stream"
V(12, 2) = Sheet4.Combo_HotSpacer.Value
V(12, 1) = "Hot Spacer"
V(13, 2) = Sheet4.Combo_MembraneMaterial
V(13, 1) = "Membrane Material"
V(14, 2) = Sheet4.Combo_MD_Type
V(14, 1) = "Membrane Distillation Type"
V(15, 2) = Sheet4.Combo_FoilMaterial
V(15, 1) = "Foil Material"
V(16, 2) = Sheet4.Combo_AirGapSpacer
V(16, 1) = "Air Gap Spacer"
V(17, 2) = Range("Horizontal_Length").Value
V(17, 1) = "Horizontal Length"
V(18, 2) = Range("Vertical_Length").Value
V(18, 1) = "Vertical Length"
V(19, 2) = Range("NumberOfLayers").Value
V(19, 1) = "Number Of Layers"
V(20, 2) = Sheet4.CheckBox_IncludeAmbient
V(20, 1) = "Include Ambient"
V(21, 2) = Sheet4.Combo_ExternalMaterial
V(21, 1) = "External Material"
V(22, 2) = Sheet4.Range("AmbientTemperatureRange")
V(22, 1) = "Abmbient Temperature"
V(23, 2) = Sheet4.OptionButton2
V(23, 1) = "1 Side exposed to ambient"
V(24, 2) = Sheet4.OptionButton1
V(24, 1) = "2 Sides exposed to ambient"
V(25, 2) = Sheet4.CheckBox_HotSideExposed
V(25, 1) = "If one side, hot side is exposed"
V(26, 2) = Range("RangeGravityDirection")
V(26, 1) = "Gravity Direction"


ReturnAllSettings = V
        

End Function

Private Function MaximumPercentChangeOfAllResults(Results As Variant, OldResults As Variant, CritPercDiff As Double, StudyDone As Boolean) As Double
    Dim i As Long
    Dim PercDiff As Double
    Dim MaxPercDiff As Double
    
    MaxPercDiff = 0
    For i = 1 To UBound(Results, 1)
        If OldResults(i, 1) <> 0 Then
           PercDiff = Abs((Results(i, 1) - OldResults(i, 1)) / (OldResults(i, 1)))
           If PercDiff > MaxPercDiff Then
              MaxPercDiff = PercDiff
           End If
        End If
    Next
    
    StudyDone = CritPercDiff > MaxPercDiff

    MaximumPercentChangeOfAllResults = MaxPercDiff
End Function


