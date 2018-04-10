VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSystemEquations"
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

Private Rslt() As Double  'This is the result of the evalution of a system of equations
Private T() As Double ' Global Temperature vector node 1 to NNode
Private m() As Double ' Global mass flow rate vector node 1 to NNode
Private S() As Double ' Global salinity vector node 1 To NNode
Private CWS As clsWaterStream 'Cold water stream
Private HWS As clsWaterStream 'Hot water stream
Private SpacerC As clsSpacer  'Cold water spacer
Private SpacerH As clsSpacer  'Hot water spacer
Private SpacerAG As clsSpacer 'Air Gap Spacer
Private Con As clsConnectivity 'Contains connectivity information for node numbers
Private Inp As clsInput
Private NEquations As Long
Private NVariables As Long
Private NNode As Long
Private CVP() As clsControlVolumePair
Private HTModel As Long ' (0) - 'Direct Contact', (1) - 'Air Gap', ... Add more later??
' Losses configuration - trickles down to clsControlVolumePair
Private LossesConfig As Long '0 = No losses, 1 = Losses on the hot side, 2 = losses on the cold side, 3 Losses on both sides
Private HotFlowMult As Double  'Indicates the multiplier on the cold flow (ussually 1/2 for symmetric, 1 for full flow simulation)
Private ColdFlowMult As Double 'Indicates the multiplier on the flow

Public Property Get HotSpacer() As clsSpacer
    Set HotSpacer = SpacerH
End Property

Public Property Let LossesConfiguration(Val As Long)
    Dim ErrMsg As String
    Dim Dum As Double
    
    LossesConfig = mdlValidation.AssignValueIfInLimits(Val, ErrMsg, "Loss Configuration = 0, 1, 2, or 3", mdlConstants.glbINVALID_VALUE, "", 0, 3)
    
    If Len(ErrMsg) <> 0 Then
       Dum = mdlError.ReturnError(ErrMsg, , True)
    End If
End Property

Public Property Get LossesConfiguration() As Long
  LossesConfiguration = LossesConfig
End Property

Public Property Get ColdSpacer() As clsSpacer
    Set ColdSpacer = SpacerC
End Property

Public Property Get AirGapSpacer() As clsSpacer
    Set AirGapSpacer = SpacerAG
End Property

Public Sub SetControlVolumePairsToNotCalulated()
      Dim i As Long
      For i = LBound(CVP) To UBound(CVP)
          CVP(i).CalculationComplete = False
      Next
End Sub

Public Property Get ControlVolumePair(CVNum As Long) As clsControlVolumePair
    Set ControlVolumePair = CVP(CVNum)
End Property

Public Property Get Temperature(i As Long) As Double
    If i < LBound(T) Or i > UBound(T) Then
        Temperature = mdlError.ReturnError("clsSystemEquations.Temperature: index " & CStr(i) & " is out of range of the Temperature vector range of " & LBound(T) & " to " & UBound(T))
    Else
        Temperature = T(i)
    End If
End Property

Public Sub SetTemperature(Temp As Double, i As Long)
    T(i) = Temp
End Sub

Public Property Get MassFlow(i As Long) As Double
    If i < LBound(m) Or i > UBound(m) Then
        MassFlow = mdlError.ReturnError("clsSystemEquations.MassFlow: index " & CStr(i) & " is out of range of the Temperature vector range of " & LBound(m) & " to " & UBound(m))
    Else
        MassFlow = m(i)
    End If
End Property

Public Sub SetMassFlow(MassFlo As Double, i As Long)
    m(i) = MassFlo
End Sub

Public Property Get Salinity(i As Long) As Double
    If i < LBound(m) Or i > UBound(m) Then
        Salinity = mdlError.ReturnError("clsSystemEquations.Salinity: index " & CStr(i) & " is out of range of the Temperature vector range of " & LBound(S) & " to " & UBound(S))
    Else
        Salinity = S(i)
    End If
End Property

Public Sub SetSalinity(Salt As Double, i As Long)
    S(i) = Salt
End Sub

Public Sub SetMassFlowMultipliers(HotMult As Double, ColdMult As Double)
    HotFlowMult = HotMult
    ColdFlowMult = ColdMult
End Sub

Public Property Get HotFlowMultiplier() As Double
    HotFlowMultiplier = HotFlowMult
End Property

Public Property Get ColdFlowMultiplier() As Double
    ColdFlowMultiplier = ColdFlowMult
End Property

Private Sub BoundaryConditions()
' The function returns the final index so that the equations function can take up the next index
' and continue to build the result of the function evaluation.

Dim i As Long
Dim Ind As Long
Dim BC() As Long
Dim ColdMassFlowBC As Double
Dim HotMassFlowBC As Double
Dim HotTemperatureBC As Double
Dim ColdTemperatureBC As Double
Dim HotSalinityBC As Double
Dim ColdSalinityBC As Double

' Avoid calling more complicated higher objects many times
ColdMassFlowBC = ColdFlowMult * CWS.MassFlow / (Con.ColdChannelWidth * Inp.NumberOfLayers)
HotMassFlowBC = HotFlowMult * HWS.MassFlow / (Con.HotChannelWidth * Inp.NumberOfLayers)
HotTemperatureBC = HWS.Temperature
ColdTemperatureBC = CWS.Temperature
HotSalinityBC = HWS.Salinity
ColdSalinityBC = CWS.Salinity

Ind = 1

Con.GetBC BC

For i = 1 To UBound(BC, 1) ' The number of boundary conditions is different than the number of nodes at which boundary conditions exist
    If BC(i, 2) = 0 Then ' This is a temperaure, mass, and salinity boundary conditions
        If BC(i, 3) = 0 Then ' 0 means this is a hot side boundary condition
            Rslt(Ind) = T(BC(i, 1)) - HotTemperatureBC
            Rslt(Ind + 1) = m(BC(i, 1)) - HotMassFlowBC
            Rslt(Ind + 2) = S(BC(i, 1)) - HotSalinityBC
        Else
            Rslt(Ind) = T(BC(i, 1)) - ColdTemperatureBC
            Rslt(Ind + 1) = m(BC(i, 1)) - ColdMassFlowBC
            Rslt(Ind + 2) = S(BC(i, 1)) - ColdSalinityBC
        End If
        Ind = Ind + 3
    Else ' This is a reversal boundary condition
        Rslt(Ind) = T(BC(i, 1)) - T(BC(i, 2))    'Temperature
        Rslt(Ind + 1) = m(BC(i, 1)) - m(BC(i, 2)) 'Mass Flow
        Rslt(Ind + 2) = S(BC(i, 1)) - S(BC(i, 2)) 'Salinity
        Ind = Ind + 3
    End If
Next

End Sub

Public Sub GetVariables(x() As Double)

'Global Variables vector is Temperatures, followed by masses, followed by Salinities

  Dim i As Long
  ReDim x(1 To 3 * NNode)
  
  
  For i = 1 To NNode
     x(i) = T(i)
  Next
  For i = 1 To NNode
     x(i + NNode) = m(i)
  Next
  For i = 1 To NNode
     x(i + NNode * 2) = S(i)
  Next

End Sub

Public Sub SetVariables(x() As Double)
  Dim i As Long
  Dim Dum As Double
  
  If UBound(x) - LBound(x) + 1 <> 3 * NNode Then
     Dum = mdlError.ReturnError("clsSystemEquations.SetVariables: The variable vector length has to be 3 times the number of nodes " & CStr(3 * NNode), , True)
     Exit Sub
  End If
  
  For i = 1 To NNode
     T(i) = x(i) ' temperature
  Next
  For i = 1 To NNode
     m(i) = x(i + NNode) 'mass flow
  Next
  For i = 1 To NNode
     S(i) = x(i + NNode * 2) 'Salinity
  Next

End Sub

Public Sub GetNodeVariables(Nodes() As Long, Temperatures() As Double, MassFlow() As Double, Salinity() As Double)

    Dim i As Long
    Dim n As Long
    
    n = UBound(Nodes)
    
    ReDim Temperatures(1 To n)
    ReDim MassFlow(1 To n)
    ReDim Salinity(1 To n)
    
    For i = 1 To n
        Temperatures(i) = T(Nodes(i))
        MassFlow(i) = m(Nodes(i))
        Salinity(i) = S(Nodes(i))
    Next i

End Sub


Public Sub EvaluateFunction(x() As Double, CV As clsControlVolumePair, SysEq As clsSystemEquations, Result() As Double)
  ' Most of the arguments are not used because the arguments are required for the Newton method but
  ' are not needed in every case.
  Dim i As Long
  
  ' Update temperatures in Me. This trickles down to the control volume pairs when they are calculated using clsControlVolumePair.CalculateControlVolumePair
  Me.SetVariables x
  
  ' Evaluate the equations
  Me.Equations Result

End Sub

Property Get NumberEquations() As Long
' This includes the number of boundary conditions.
    NumberEquations = NEquations
End Property

Public Sub Equations(Result() As Double)
   
   Dim Ind As Long
   Dim i As Long
   Dim j As Long
   Dim CV_ind As Long
   
   BoundaryConditions
   Ind = Con.NumberOfBoundaryConditions + 1
   ' solve in order of gravity so that we can keep track of air gap
   ' thickness. This doesn't effect direct contact but will become important
   ' if we quantify pressure drops
   For i = 1 To Con.NumberOfGravityPaths
      For j = 1 To Con.NumberOfGravityLayers
        CV_ind = Con.GravityControlVolumePath(j, i)
        CVP(CV_ind).CalculateControlVolumePair Me
        ' Return energy results
        Rslt(Ind) = CVP(CV_ind).EnergyEquation(Rslt(Ind + 1)) ' Hot result is the function value, cold is the input.
        Rslt(Ind + 2) = CVP(CV_ind).MassEquation(Rslt(Ind + 3))
        Rslt(Ind + 4) = CVP(CV_ind).SalinityEquation(Rslt(Ind + 5))
        ' So far there is no need for salinity in the cold stream.  No mass
        ' transfer is occuring and in direct contact the salinity is zero.  If
        ' we begin to model precipitation of salt or another chemical, then a species balance will be needed.
        Ind = Ind + 6
      Next
   Next
   
   ReDim Result(LBound(Rslt) To UBound(Rslt))
   For i = LBound(Rslt) To UBound(Rslt)
       Result(i) = Rslt(i)
   Next i

End Sub

Property Get Connectivity() As clsConnectivity
   Set Connectivity = Con
End Property

Property Get Inputs() As clsInput
   Set Inputs = Inp
End Property

Property Let IncludeExternalLosses(Val As Boolean)
    IncludeExtLoss = Val
End Property

Property Get IncludeExternalLosses() As Boolean
   IncludeExternalLosses = IncludeExtLoss
End Property



Public Sub InitlializeSystemEquations(Optional IncludeCustomInputSubRoutine As Boolean = False, _
                     Optional CustomSubroutineName As String = "", _
                     Optional SubCaseValue As Variant)
    Dim Dum As Double
    Dim i As Long
    Dim T_hot_out As Double
    Dim T_cold_out As Double
    Dim Ind As Double
    Dim T_hot_in As Double
    Dim T_cold_in As Double
    Dim cph As Double
    Dim cpc As Double
    Dim Lavg As Double
    
    
    ' These operations should only be performed once
    Set Inp = New clsInput
    Set Con = New clsConnectivity
    Set CWS = Inp.WaterStreams(Inp.ColdWaterStream)
    Set HWS = Inp.WaterStreams(Inp.HotWaterStream)
    Set SpacerC = Inp.Spacers(Inp.ColdSpacer)
    Set SpacerH = Inp.Spacers(Inp.HotSpacer)
    Set SpacerAG = Inp.Spacers(Inp.AirGapSpacer)
    
    ' Estimate the output temperatures based on the maximum amount of heat flow
    ' that is possible in order to
    T_hot_in = HWS.Temperature
    T_cold_in = CWS.Temperature
    cph = mdlProperties.SeaWaterSpecificHeat(T_hot_in, HWS.Salinity)
    cpc = mdlProperties.SeaWaterSpecificHeat(T_cold_in, CWS.Salinity)
    Lavg = mdlProperties.LatentHeatOfPureWater((T_hot_in + T_cold_in) / 2)
    

    
    mdlTransferCoefficient.EstimateOutputTemperatures T_hot_in, T_cold_in, HotFlowMult * HWS.MassFlow / (Con.HotChannelWidth * Inp.NumberOfLayers), _
                                                      ColdFlowMult * CWS.MassFlow / (Con.ColdChannelWidth * Inp.NumberOfLayers), HWS.Salinity, CWS.Salinity, 0#, cpc, _
                                                      cph, Lavg, T_hot_out, T_cold_out
                                                      
    'Before we initialize the control volume pairs,
    ' Apply a custom subroutine if requested. This custom subroutine
    ' is only for altering input to the model. It cannot have a direct effect
    ' on any of the calculated values because it is only in the model initialization
    If IncludeCustomInputSubRoutine Then
        CustomSubroutine CustomSubroutineName, SubCaseValue
    End If
                                                      

    'Initialize all control volume pairs
    ReDim CVP(1 To Con.NumberOfControlVolumePairs)
    
    For i = 1 To Con.NumberOfControlVolumePairs
        Set CVP(i) = New clsControlVolumePair
        CVP(i).InitializeControlVolumePair Me, i
    Next
    
    ' Adjust all of the temperatures to reflect the hot and cold output temperatures estimated
    SetCVInitialTemperatures T_hot_out, T_hot_in, T_cold_out, T_cold_in

    NNode = Con.NumberOfNodes
    
    ' !@#$ if more species need to be tracked (such as fouling of CaCO3) then more balance equations are needed.
    '      for direct contact, each control volume has an energy balance, mass balance, and salinity balance.
    '      there is no need for additional variables because the interface temperatures can be solved directly
    NEquations = (Con.NumberOfControlVolumePairs) * 6 + Con.NumberOfBoundaryConditions 'In this code, boundary conditions are treated as equations.
    NVariables = 3 * NNode  ' Temperature, Mass Flow, and Salinity at each node .
    
    If NEquations <> NVariables Then 'There is an inconsistency that needs to be resolved!
        Dum = mdlError.ReturnError("clsSystemEquations.Class_Initialize: The number of equations does not match the number of variables! " & _
                  "There must be an inconsistency in the Connectivity class clsConnectivity! This may be hard to troubleshoot.", , True)
    End If
    
    
    If Inp.MembraneDistillationType = "Direct Contact" Then
        HTModel = 0
    ElseIf Inp.MembraneDistillationType = "Air Gap" Then
        HTModel = 1
    End If
    
    ReDim Rslt(1 To NEquations)
    ReDim T(1 To NNode)
    ReDim m(1 To NNode)
    ReDim S(1 To NNode) ' - all cold side salinity just stays zero all the time.
    ' Now assign the proper values to the temperatures T, salinities S, and mass flows m
    For i = 1 To Con.NumberOfControlVolumePairs
        CVP(i).GetVariables Me
    Next i
    
End Sub

Private Sub CustomSubroutine(CustomSubroutineName As String, Optional SubCaseValue As Variant)

Dim Dum As Double
' If you are adding a new custom subroutine, then you must add a new entry to the case-select statement
' below!
Select Case CustomSubroutineName
   Case "NREL_Validation"
      NREL_Parameter_Changes Me, SubCaseValue
   Case Else
      Dum = mdlError.ReturnError("clsSystemEquations.CustomSubroutine: A custom subroutine was requested but the name supplied=""" & _
                                 CustomSubroutineName & """ is not a valid name. You must add this name manually to this function!", , True, True)
End Select

End Sub

Private Sub SetCVInitialTemperatures(T_hot_out As Double, T_hot_in As Double, T_cold_out As Double, T_cold_in As Double)

    Dim T_hot_inc As Double
    Dim T_cold_inc As Double
    Dim CVHotPaths() As Long
    Dim CVColdPaths() As Long
    Dim HotPaths() As Long
    Dim ColdPaths() As Long
    Dim Ind As Long
    ' get the flow paths and calculate the increment _inc between nodes as estimated
    ' by the initial condition being set by T_hot_out and T_cold_out
    
    Con.GetFlowPaths HotPaths, ColdPaths, CVHotPaths, CVColdPaths
    T_hot_inc = (T_hot_out - T_hot_in) / UBound(CVHotPaths, 1) ' This should come out negative!
    T_cold_inc = (T_cold_out - T_cold_in) / UBound(CVColdPaths, 1) ' This should come out positive!
    
    InitializeCVofOneSide CVHotPaths, T_hot_inc
    InitializeCVofOneSide CVColdPaths, T_cold_inc
    ReinitializeInterfaceTemperaturesForNewBoundaryCondition

End Sub

Private Sub InitializeCVofOneSide(CVPaths() As Long, T_inc As Double)
       ' Initialize all of the control volume pairs along the flow paths so that
       ' better initial conditions can be set.
       ' For HOT T_inc should be negative.
       ' For Cold T_inc should be possitive.
    Dim i As Long
    Dim j As Long

    
    For i = 1 To UBound(CVPaths, 2) ' Start with 2 so that each column is treated individually
        For j = 1 To UBound(CVPaths, 1)
           CVP(CVPaths(j, i)).IncrementTemperature T_inc, j  ' We assume a linear decrease
        Next
    Next
    
End Sub

Private Sub ReinitializeInterfaceTemperaturesForNewBoundaryCondition()
    Dim i As Long
    
    For i = 1 To UBound(CVP)
        CVP(i).InitializeGuessForInterfaceTemperatures
    Next
End Sub

Public Sub ReturnMaximumAndMinimumForVariables(XkMax() As Double, XkMin() As Double)
   ReDim XkMax(1 To Me.NumberEquations)
   ReDim XkMin(1 To Me.NumberEquations)
   
   Dim HWS As clsWaterStream
   Dim CWS As clsWaterStream
   Dim i As Long
   
   Set HWS = Me.Inputs.WaterStreams(Me.Inputs.HotWaterStream)
   Set CWS = Me.Inputs.WaterStreams(Me.Inputs.ColdWaterStream)
   
   ' Temperature
   For i = 1 To NNode
      XkMax(i) = HWS.Temperature
   Next
   'Mass Flow
   For i = 1 + NNode To 2 * NNode
      XkMax(i) = HotFlowMult * HWS.MassFlow / (Con.HotChannelWidth * Inp.NumberOfLayers) + ColdFlowMult * CWS.MassFlow / (Con.ColdChannelWidth * Inp.NumberOfLayers)
   Next
   ' Salinity
   For i = 1 + 2 * NNode To 3 * NNode
      XkMax(i) = 1
   Next
   
   
   For i = 1 To NNode
      XkMin(i) = CWS.Temperature
   Next
   For i = 1 + NNode To 2 * NNode
      XkMin(i) = 0
   Next
   For i = 1 + 2 * NNode To 3 * NNode
      XkMin(i) = 0
   Next


End Sub

Private Sub Class_Initialize()
   HotFlowMult = 1
   ColdFlowMult = 1
End Sub
