VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsControlVolumePair"
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

' This class provides equations for mass and energy based on one of many configurations
Private CVNum As Long
Private PreviousGravityLayerCVNum As Long
Private DCHT As clsHTFuncDirectContact 'Direct contact heat transfer function class (only used if direct contact heat transfer is used).
Private AGHT As clsHTFuncAirGap

' These place the CV in the node numbering scheme
' contained in system equations for the SysEq.Temperature, SysEq.MassFlow, SysEq.Salinity Vectors
Private HotInflowNode As Long
Private HotOutflowNode As Long
Private ColdInflowNode As Long
Private ColdOutflowNode As Long
Private Nodes() As Long 'in order 1 - hot inflow, 2 - hot outflow, 3 - coldinflow, 4 cold outflow

Private HotInflowNodePosition As Long ' This is the position w/r to the entire flow path used to help initialize the variables
Private ColdInflowNodePosition As Long 'This is the position w/r to the entire cold flow path used to help initialize the variables

Private HotInflowEdge As Long
Private ColdInflowEdge As Long

' Losses configuration
Private LossesConfig As Long '0 = No losses, 1 = Losses on the hot side, 2 = losses on the cold side, 3 Losses on both sides

' All of these are constant but may become non-constant for
' a spiral case in the future.
Private PrLengthNormalToHotFlow As Double
Private PrLengthNormalToColdFlow As Double
Private PrLengthParallelToHotFlow As Double
Private PrLengthParallelToColdFlow As Double
Private PrTotalLengthParallelToHotFlow As Double
Private PrTotalLengthParallelToColdFlow As Double
Private PrLengthParallelToGravity As Double
Private PrColdThickness As Double
Private PrHotThickness As Double
Private PrAirGapThickness As Double
Private PrCalcComplete As Boolean
Private ColdFlowMult As Double
Private HotFlowMult As Double

'Local variables that are accessed via properties if they are
'needed externally.
' Direct contact and air gap local variables (shared)
Private Tih As Double 'Hot interfacial temperature '
Private Tic As Double 'Cold interfacial temperature

' Air gap local variables
Private Tma As Double     ' temperature at the membrane/air gap interface (Kelvin)
Private Tf As Double      ' temperature at the air gap/condensate interface (Kelvin)
Private Tfw As Double     ' temperature at the condensate/cooling wall foil interface (Kelvin)
Private delta_w As Double ' average condensate thickness (condensate thickness variation is modeled as linear)
Private Pv_ma As Double   ' vapor pressure of air at the membrane/air gap interface (air is not necessarily saturated).
Private X_ma As Double     ' Quality at the membrane to air gap interface (ussually 1 but some applications may produce liquid aerosols at the interface)
'
Private Tc() As Double 'Cold Temperature in (0) out (1)
Private Th() As Double 'Hot Temperature in (0) out (1)
Private SC() As Double 'Cold Salinity in (0) out (1)
Private SH() As Double 'Hot Salinity in (0) out (1)
Private mC() As Double 'Cold side mass flow in (0) out (1)
Private mH() As Double 'Hot side mass flow in(0) out (1)
Private mM As Double 'Mass flow due to membrane distillation effect.
Private cph() As Double 'Hot side specific heat of water ( in - 0 out - 1)
Private cpc() As Double 'Cold side specific heat of water (in - 0 out - 1)
Private Hh As Double 'Hot side heat transfer coefficient
Private Hc As Double 'Cold side heat transfer coefficient
Private Hm As Double 'Membrane heat transfer coefficient
Private A_f As Double 'Membrane heat transfer area
Private Sih As Double ' interface salinity hot side (more concentrated due to mass boundary layer)
Private Sic As Double ' interface salinity cold side.
Private T_Air As Double ' External ambient temperature
Private Lavg As Double 'Latent heat of evaporation at the average of the control volume.

' Air gap only variables
Private delta_condensate_exit As Double ' this is the exit film thickness resulting from the cumulative flow of condensate for air gap calculations only! (opposite the gravitational vector) control volume
                   ' for layer 1 control volumes this is always zero. This is an actual variable in the air gap analysis. The value stored here
                   ' is the final converged value after newton's method is done on a CV level.
'Air gap only
Private TotalThickness As Double
Private AGThick As Double
Private MembThick As Double
Private FoilThick As Double
Private FoilMat As clsMaterial
Private MembMat As clsMaterial
                   
                   
' Other variables
Private HTModel As Long ' (0) - 'Direct Contact', (1) - 'Air Gap', ... Add more later??
Private Q As Double ' heat transfer rate through the membrane - depends on the local membrane heat transfer functions
Private Q_loss_hot As Double 'heat transfer rate to ambient that are lost from the hot side.
Private Q_loss_cold As Double 'heat transfer rate to ambient from the cold side

Private CWS As clsWaterStream ' cold water stream
Private HWS As clsWaterStream ' hot water stream

Property Get HeatTransferLateral() As Double
    HeatTransferLateral = Q
End Property

Property Get HeatLosses() As Double
    HeatLosses = Q_loss_hot + Q_loss_cold
End Property

Property Get HeatTransferDueToMassTransfer() As Double
   HeatTransferDueToMassTransfer = mM * Lavg
End Property

Property Get LengthParallelToGravity() As Double
    LengthParallelToGravity = PrLengthParallelToGravity
End Property

Property Get HeatTransferModel() As Long
    HeatTransferModel = HTModel
End Property

Property Get CondensateExitThicknessAirGapOnly() As Double
    CondensateExitThicknessAirGapOnly = delta_condensate_exit
End Property

Property Get PreviousGravityLayerCVNumber() As Long
     PreviousGravityLayerCVNumber = PreviousGravityLayerCVNum
End Property

Property Get ViolatesSecondLawOfThermodynamics(Optional KillOnError As Boolean = True) As Boolean
      
      Dim Dum As Double
      
      Dim ErrMsg As String
      If Tc(0) > Th(0) Then
          ErrMsg = "cold input temperature is greater than the hot input temperature."
      ElseIf Tc(0) > Th(1) Then
          ErrMsg = "cold input temperature is greater than the hot output temperature."
      ElseIf Tc(0) > Tc(1) Then
          ErrMsg = "cold output temperature is less than the cold input temperature."
      ElseIf Tc(1) > Th(0) Then
          ErrMsg = "cold output temperture is greater than the hot input temperature."
      ElseIf Th(1) > Th(0) Then
          ErrMsg = "hot output temperature is greater than the hot input temperature."
      ElseIf Tih < Tic Then
          ErrMsg = "hot interface temperature is less than cold interface temperature."
      ElseIf Tih > Th(0) Then
          ErrMsg = "hot interface temperature is greater than the hot input temperature."
      ElseIf Tic > Th(0) Then
          ErrMsg = "cold interface temperature is greater than the hot input temperature."
      ElseIf Tih < Tc(0) Then
          ErrMsg = "hot interface temperature is less than cold input temperature."
      ElseIf Tic < Tc(0) Then
          ErrMsg = "cold interface temperature is less than cold input temperature."
      End If
      
      If Len(ErrMsg) <> 0 Then
         ViolatesSecondLawOfThermodynamics = True
         Dum = mdlError.ReturnError("clsControlVolumePair.ViolatesSecondLawOfThermodynamics: Control volume " & CStr(CVNum) & _
                              " violates the second law of thermodynamics! The " & ErrMsg, , True, KillOnError)
      Else
         ViolatesSecondLawOfThermodynamics = False
      End If
      
End Property

Property Get ColdThickness() As Double
    ColdThickness = PrColdThickness
End Property

Property Get MembraneDistillationMassFlow() As Double
    MembraneDistillationMassFlow = mM
End Property

Property Get SpecificHeats(IsHot As Boolean, IsIn As Boolean) As Double
   If IsHot And IsIn Then
       SpecificHeats = cph(0)
   ElseIf IsHot Then
       SpecificHeats = cph(1)
   ElseIf IsIn Then
       SpecificHeats = cpc(0)
   Else
       SpecificHeats = cpc(1)
   End If
End Property

Property Get Area() As Double
    Area = A_f
End Property

Property Get HotThickness() As Double
   HotThickness = PrHotThickness
End Property

Property Get AirGapThickness() As Double
   AirGapThickness = PrAirGapThickness ' For Air gap heat transfer only (maybe in the future for vacuum membrane distillation and sweeping gas membrane distillation
End Property

Property Get CalculationComplete() As Boolean
     CalculationComplete = PrCalcComplete
End Property

Property Let CalculationComplete(Bool As Boolean)
     PrCalcComplete = Bool
End Property

Property Get LengthNormalToHotFlow() As Double
    LengthNormalToHotFlow = PrLengthNormalToHotFlow
End Property

Property Get LengthNormalToColdFlow() As Double
    LengthNormalToColdFlow = PrLengthNormalToColdFlow
End Property

Property Get LengthParallelToHotFlow() As Double
    LengthParallelToHotFlow = PrLengthParallelToHotFlow
End Property

Property Get LengthParallelToColdFlow() As Double
    LengthParallelToColdFlow = PrLengthParallelToColdFlow
End Property

Private Sub CalculateSpecificHeats()
   cph(0) = mdlProperties.SeaWaterSpecificHeat(Th(0), SH(0), "Jamieson_et_al")
   cph(1) = mdlProperties.SeaWaterSpecificHeat(Th(1), SH(1), "Jamieson_et_al")
   cpc(0) = mdlProperties.SeaWaterSpecificHeat(Tc(0), SC(0), "Jamieson_et_al")
   cpc(1) = mdlProperties.SeaWaterSpecificHeat(Tc(1), SC(1), "Jamieson_et_al")
End Sub

Public Property Get FlowType() As Long
' 1 = co-flow
' 2 = counter-flow
' 3 = cross-flow

' These values are strictly controlled to be between 1 and 4 elsewhere
' no need for error control.
Select Case HotInflowEdge
     Case 1
         Select Case ColdInflowEdge
             Case 1
                 FlowType = 1
             Case 2
                 FlowType = 3
             Case 3
                 FlowType = 2
             Case 4
                 FlowType = 3
         End Select
     Case 2
         Select Case ColdInflowEdge
             Case 1
                 FlowType = 3
             Case 2
                 FlowType = 1
             Case 3
                 FlowType = 3
             Case 4
                 FlowType = 2
         End Select
     Case 3
         Select Case ColdInflowEdge
             Case 1
                 FlowType = 2
             Case 2
                 FlowType = 3
             Case 3
                 FlowType = 1
             Case 4
                 FlowType = 3
         End Select
     Case 4
         Select Case ColdInflowEdge
             Case 1
                 FlowType = 3
             Case 2
                 FlowType = 2
             Case 3
                 FlowType = 3
             Case 4
                 FlowType = 1
         End Select
End Select

End Property

Public Function EnergyEquation(ColdResult As Double) As Double

Dim Tref As Double

Tref = mdlConstants.glbNISTReferenceTemperature

EnergyEquation = mH(0) * cph(0) * (Th(0) - Tref) - mH(1) * cph(1) * (Th(1) - Tref) - Q - Q_loss_hot
ColdResult = mC(0) * cpc(0) * (Tc(0) - Tref) - mC(1) * cpc(1) * (Tc(1) - Tref) + Q - Q_loss_cold

End Function

Public Function MassEquation(ColdResult As Double) As Double


    ' For both air gap and direct contact, MD mass transfer is subtracted
    MassEquation = mH(0) - mH(1) - mM

    Select Case HTModel
       Case 0 'For direct contact membrane distillation mM is added to the cold stream
           ColdResult = mC(0) - mC(1) + mM
       Case 1 ' For Air Gap (or vacuum) MD, mM becomes a separate stream and the cooling stream mass flow does not change.
           ColdResult = mC(0) - mC(1)
    End Select


End Function

Public Function SalinityEquation(ColdResult As Double) As Double

   SalinityEquation = mH(0) * SH(0) - mH(1) * SH(1)  ' Salinity only increases due to changes in mass flow from MD.  No precipitation or fouling modeled currently.

   ColdResult = mC(0) * SC(0) - mC(1) * SC(1)  ' No change in salinity is anticipated on the cold side unless fouling is added.  For Air gap salinity is expected

End Function

Public Sub IncrementTemperature(T_inc As Double, Position As Long)
    ' This is for setting the initial condition and should not be used afterward!
    If T_inc < 0 Then
        Th(0) = Th(0) + T_inc * (Position - 1)
        Th(1) = Th(1) + T_inc * Position
    Else
        Tc(0) = Tc(0) + T_inc * (Position - 1)
        Tc(1) = Tc(1) + T_inc * Position
    End If
End Sub

Public Sub InitializeGuessForInterfaceTemperatures()
      Dim Th_avg As Double
      Dim Tc_avg As Double
      Dim T_inc As Double
      
      Th_avg = (Th(0) + Th(1)) / 2
      Tc_avg = (Tc(0) + Tc(1)) / 2
      T_inc = glbInitialFractionToColdOrHotForInterfaceTemperature * (Th_avg - Tc_avg)
      
      Tih = Th_avg - T_inc
      Tic = Tc_avg + T_inc
      
      If HTModel = 1 Then 'Air gap variables need to be initialized
         ' Just make the temperature drop as a linear function of distance. Across the air gap temperatures.
         Tma = Tih - (Tih - Tic) * (MembThick / TotalThickness)
         Tf = Tih - (Tih - Tic) * (MembThick + AGThick - delta_w) / TotalThickness
         Tfw = Tih - (Tih - Tic) * (MembThick + AGThick) / TotalThickness
         Pv_ma = mdlConstants.glbInitialGuessCondensateThickFraction * mdlProperties.SaturatedPressurePureWater(Tma)
      End If
      
      
End Sub

Public Sub InitializeControlVolumePair(SysEq As clsSystemEquations, CVNumber As Long)
  'We are going to assume that the initial temperature distribution, salinity and mass distribution do not change at all from the inputs
  'A couple of iterations will correct this assumption

  'T_offset is the amount of temperature to offset from the input temperature
  Dim Inp As clsInput
  Dim Con As clsConnectivity
  Dim HotInflowNodeInd As Long
  Dim HotOutflowNodeInd As Long
  Dim ColdInflowNodeInd As Long
  Dim ColdOutflowNodeInd As Long
  Dim GravDir As Long
  Dim Dum As Double

  
  ColdFlowMult = SysEq.ColdFlowMultiplier
  HotFlowMult = SysEq.HotFlowMultiplier
  
  ' This has to be initialized for the values to stick.
  mdlConstants.GlobalArrays
  
  If SysEq.Inputs.MembraneDistillationType = mdlConstants.glbMembraneDistillationTypes(1) Then
     HTModel = 0
  ElseIf SysEq.Inputs.MembraneDistillationType = mdlConstants.glbMembraneDistillationTypes(2) Then
     HTModel = 1
  Else
     Dum = mdlError.ReturnError("clsControlVolumePair.InitializeControlVolumePair: The SysEq.Inputs.MembraneDistillationType = " & SysEq.Inputs.MembraneDistillationType & " is not a valid option!", , True)
  End If
  
  'CVNum is the internal private property that will be retained.
  CVNum = CVNumber
  
  Set Inp = SysEq.Inputs
  Set Con = SysEq.Connectivity

  'These are measures that do not need to be recalculated.
  A_f = (Inp.HorizontalLength / Con.HorizontalDivisions) * (Inp.VerticalLength / Con.VerticalDivisions)
  
  If Con.HotFlowsHorizontal Then
     PrLengthNormalToHotFlow = Inp.VerticalLength / Con.VerticalDivisions
     PrLengthParallelToHotFlow = Inp.HorizontalLength / Con.HorizontalDivisions
     PrTotalLengthParallelToHotFlow = Inp.HorizontalLength
  Else
     PrLengthNormalToHotFlow = Inp.HorizontalLength / Con.HorizontalDivisions
     PrLengthParallelToHotFlow = Inp.VerticalLength / Con.VerticalDivisions
     PrTotalLengthParallelToHotFlow = Inp.VerticalLength
  End If
  
  If Con.ColdFlowsHorizontal Then
     PrLengthNormalToColdFlow = Inp.VerticalLength / Con.VerticalDivisions
     PrLengthParallelToColdFlow = Inp.HorizontalLength / Con.HorizontalDivisions
     PrTotalLengthParallelToColdFlow = Inp.HorizontalLength
  Else
     PrLengthNormalToColdFlow = Inp.HorizontalLength / Con.HorizontalDivisions
     PrLengthParallelToColdFlow = Inp.VerticalLength / Con.VerticalDivisions
     PrTotalLengthParallelToColdFlow = Inp.VerticalLength
  End If
  
  PrColdThickness = SysEq.ColdSpacer.Thickness * ColdFlowMult 'reduced to 1/2 when symmetry applied on one side but not the other.
  PrHotThickness = SysEq.HotSpacer.Thickness * HotFlowMult
  PrAirGapThickness = SysEq.AirGapSpacer.Thickness
  
  Set CWS = Inp.WaterStreams(Inp.ColdWaterStream)
  Set HWS = Inp.WaterStreams(Inp.HotWaterStream)
  
  HotInflowNodeInd = Con.NodeOrder(1, CVNum)
  HotOutflowNodeInd = Con.NodeOrder(2, CVNum)
  ColdInflowNodeInd = Con.NodeOrder(3, CVNum)
  ColdOutflowNodeInd = Con.NodeOrder(4, CVNum)
  
  ReDim Nodes(1 To 4)
  
  HotInflowNode = Con.c(CVNum, HotInflowNodeInd)
  HotOutflowNode = Con.c(CVNum, HotOutflowNodeInd)
  ColdInflowNode = Con.c(CVNum, ColdInflowNodeInd)
  ColdOutflowNode = Con.c(CVNum, ColdOutflowNodeInd)
  
  Nodes(1) = HotInflowNode
  Nodes(2) = HotOutflowNode
  Nodes(3) = ColdInflowNode
  Nodes(4) = ColdOutflowNode
  
  HotInflowEdge = Con.InflowEdge(True, CVNum)
  ColdInflowEdge = Con.InflowEdge(False, CVNum)
  
  GravDir = SysEq.Connectivity.GravityDirection
  If GravDir = 1 Or GravDir = 3 Then
     PrLengthParallelToGravity = Inp.VerticalLength / Con.VerticalDivisions
  Else
     PrLengthParallelToGravity = Inp.HorizontalLength / Con.HorizontalDivisions
  End If
  ' These are variables that need an initial value
  ReDim Th(0 To 1)
  ReDim mH(0 To 1)
  ReDim SH(0 To 1)
  Th(0) = HWS.Temperature
  Th(1) = HWS.Temperature
  
  ' Some of These values will be adjusted later on IN InitializeGuessForInterfaceTemperatures!
  Tih = HWS.Temperature - glbInitialFractionToColdOrHotForInterfaceTemperature * (HWS.Temperature - CWS.Temperature)
  mH(0) = HotFlowMult * HWS.MassFlow / (Con.HotChannelWidth * Inp.NumberOfLayers)
  mH(1) = HotFlowMult * HWS.MassFlow / (Con.HotChannelWidth * Inp.NumberOfLayers)
  SH(0) = HWS.Salinity
  SH(1) = HWS.Salinity
  
  ReDim Tc(0 To 1)
  ReDim mC(0 To 1)
  ReDim SC(0 To 1)
  Tc(0) = CWS.Temperature
  Tc(1) = CWS.Temperature
  Tic = CWS.Temperature + glbInitialFractionToColdOrHotForInterfaceTemperature * (HWS.Temperature - CWS.Temperature)
  mC(0) = ColdFlowMult * CWS.MassFlow / (Con.ColdChannelWidth * Inp.NumberOfLayers)
  mC(1) = ColdFlowMult * CWS.MassFlow / (Con.ColdChannelWidth * Inp.NumberOfLayers)
  SC(0) = CWS.Salinity
  SC(1) = CWS.Salinity
  
  T_Air = Inp.ExternalAmbientTemperature
  LossesConfig = SysEq.LossesConfiguration
  
  ' These values will be adjusted to a more accurate guess later! IN InitializeGuessForInterfaceTemperatures
  If HTModel = 1 Then 'Air gap variables need to be initialized
     Set MembMat = SysEq.Inputs.Materials(SysEq.Inputs.MembraneMaterial)
     MembThick = MembMat.Thickness
     Set FoilMat = SysEq.Inputs.Materials(SysEq.Inputs.FoilMaterial)
     FoilThick = FoilMat.Thickness
     AGThick = SysEq.AirGapSpacer.Thickness
     
     delta_w = mdlConstants.glbInitialGuessCondensateThickFraction * AGThick  '
     
     TotalThickness = MembThick + FoilThick + AGThick
     
     ' Just make the temperature drop as a linear function of distance.
     Tma = Tih - (Tih - Tic) * (MembThick / TotalThickness)
     Tf = Tih - (Tih - Tic) * (MembThick + AGThick - delta_w) / TotalThickness
     Tfw = Tih - (Tih - Tic) * (MembThick + AGThick) / TotalThickness
     Pv_ma = mdlConstants.glbInitialGuessCondensateThickFraction * mdlProperties.SaturatedPressurePureWater(Tma)
     X_ma = 1
     
  End If
  
  ' This is just purely made up to start the calculations
  mM = mdlConstants.glbInitialGuessAtFractionofMDMassFlow * mH(0)
  
  '
  ' Establish the CV from which air gap MD condensate flows from.  If this is a layer
  ' 1 (w/r to gravity) CV then a glbINVALID_VALUE is assigned.
  PreviousGravityLayerCVNum = SysEq.Connectivity.FindPreviousGravityCVGivenCVNum(CVNum)
  
  PrCalcComplete = False
  
End Sub

Public Property Get TotalLengthParrallelToHotFlow() As Double
    TotalLengthParrallelToHotFlow = PrTotalLengthParallelToHotFlow
End Property

Public Property Get TotalLengthParrallelToColdFlow() As Double
    TotalLengthParrallelToColdFlow = PrTotalLengthParallelToColdFlow
End Property

Public Sub GetVariables(SysEq As clsSystemEquations)
  ' This function transfers all of the variables back to the
  ' global numbering system for the SysEq temperature, MassFlow, and Salinity vectors
  ' These vectors are later compiled into a single vector for the complete global variable vector.
  
   ' Temperatures
   SysEq.SetTemperature Th(0), HotInflowNode
   SysEq.SetTemperature Th(1), HotOutflowNode
   SysEq.SetTemperature Tc(0), ColdInflowNode
   SysEq.SetTemperature Tc(1), ColdOutflowNode
   
   ' Mass Flows
   SysEq.SetMassFlow mH(0), HotInflowNode
   SysEq.SetMassFlow mH(1), HotOutflowNode
   SysEq.SetMassFlow mC(0), ColdInflowNode
   SysEq.SetMassFlow mC(1), ColdOutflowNode
   
   ' Salinity
   SysEq.SetSalinity SH(0), HotInflowNode
   SysEq.SetSalinity SH(1), HotOutflowNode
   SysEq.SetSalinity SC(0), ColdInflowNode
   SysEq.SetSalinity SC(1), ColdOutflowNode

End Sub

Private Sub SetInternalVariables(T_G As Double, m_G As Double, s_G As Double, Temperature As Double, MassFlow As Double, Salinity As Double)

   
   'T_G, m_G, and s_G are the global (SysEq) temperature, mass flow and salinity.
   
   ' This function checks to see if there is a change
   ' between the global vector and the local values in the control volume.
   ' if there is a change then PrCalcComplete is set to false so that the control volume
   ' local heat exchange equations will be solved again.  If there is no change PrCalcComplete will remain
   ' true and no new solution is needed. This happens a lot if newton's method is being used because the Jacobian is being reformed
   ' but variation on a single variable only effects a small subset of the control volumes.
           
   If T_G <> Temperature Then
      Temperature = T_G
      PrCalcComplete = False
   End If
   If m_G <> MassFlow Then
      MassFlow = m_G
      PrCalcComplete = False
   End If
   If s_G <> Salinity Then
      Salinity = s_G
      PrCalcComplete = False
   End If

   
End Sub

Public Property Get CVNumber() As Long
       CVNumber = CVNum
End Property


Public Sub SetVariables(T_G() As Double, m_G() As Double, s_G() As Double)
  
  Dim Temp As Double
   'Temperature - each variable has to be checked because if nothing has changed
   ' then there is no reason to recalculate the control volume internal variables (which can be costly to repeat a lot)
   
   SetInternalVariables T_G(1), m_G(1), s_G(1), Th(0), mH(0), SH(0)
   SetInternalVariables T_G(2), m_G(2), s_G(2), Th(1), mH(1), SH(1)
   SetInternalVariables T_G(3), m_G(3), s_G(3), Tc(0), mC(0), SC(0)
   SetInternalVariables T_G(4), m_G(4), s_G(4), Tc(1), mC(1), SC(1)
   '
End Sub

Public Sub CalculateControlVolumePair(SysEq As clsSystemEquations)

   ' I wish this was less verbose but it is the best way to transfer the information
   Dim Tavgh As Double 'hot side average temperature
   Dim Tavgc As Double 'cold side average temperature
   Dim Tavg As Double ' overall average temperature of hot and cold sides
   Dim Pavgh As Double ' hot side average pressure
   Dim Pavgc As Double ' cold side average pressure
   Dim Mavgh As Double ' hot side average mass flow rate
   Dim Mavgc As Double ' cold side average mass flow rate
   Dim Savgh As Double ' hot side average salinity (not a simple function of input and output salinity
   Dim Savgc As Double 'cold side average salinity - ussually not as important - no concentration difference is calculated here.
   
   Dim LR() As Double 'Local Result vector (for nonlinear equations to be solved at sub-CV's within each CV level)
   Dim LX() As Double 'Local Solution vector for heat transfer
   Dim LXmax() As Double
   Dim LXmin() As Double
   Dim TGlobal() As Double 'global temperatures
   Dim mGlobal() As Double 'global mass flows
   Dim SGlobal() As Double 'global salinities
   Dim Bool As Boolean
   Dim ConvergCrit() As Double
   Dim AGMaxVapPressure As Double
   Dim HotSpacer As clsSpacer
   Dim ColdSpacer As clsSpacer
   Dim AGSatVapPressure As Double
   Dim C_star As Double
   Dim DebugString As String
   Dim NTU As Double
   Dim Dum As Double
   Dim Emax As Double
   Dim C_min As Double
   Dim delThmax As Double
   Dim delTcmax As Double
   Dim cpminh As Double
   Dim cpminc As Double
     
   ' Update the variables if they are different
   
   SysEq.GetNodeVariables Nodes, TGlobal, mGlobal, SGlobal
   Me.SetVariables TGlobal, mGlobal, SGlobal
   
   Set HotSpacer = SysEq.HotSpacer
   Set ColdSpacer = SysEq.ColdSpacer
   
   ' Check to see if the variable values are valid given constraints of the second law of thermodynamics the code will end with an error if
   ' 2nd law is violated
   
   If Not PrCalcComplete Then 'And Not Me.ViolatesSecondLawOfThermodynamics Then ' Only calculate if the external boundary conditions are different (when calculating a Jacobian matrix, most of the control volumes will have the
   '                            exact same boundary conditions and should not be recalculated.)
        If glbEvaluateEquationsAtEnd Then
           If IsFinalRun Then
              DebugString = ""
              DebugString = DebugString & "Control Volume: " & Me.CVNumber & ": CALCULATED FINAL VALUES NOT USSUALLY OUTPUT" & vbCrLf
              Debug.Print DebugString
              mdlValidation.WriteDebugInfoToFile DebugString
           End If
        End If
   
       CalculateSpecificHeats
       
        ' This needs adjustment
       Tavgh = 0.5 * (Th(0) + Th(1))
       Tavgc = 0.5 * (Tc(0) + Tc(1))
       Tavg = 0.5 * (Tavgh + Tavgc)
       
       Mavgh = 0.5 * (mH(0) + mH(1))
       Mavgc = 0.5 * (mC(0) + mC(1))
       
       ' Salinity (not just a simple average)
       Savgh = (2 * mH(0) / (mH(0) + mH(1))) * SH(0)
       Savgc = (2 * mC(0) / (mC(0) + mC(1))) * SC(0)
       
       
       Sih = InterfaceSalinityConcentration(Tavgh, Mavgh, Savgh, mM, A_f, PrLengthNormalToHotFlow, PrTotalLengthParallelToHotFlow, PrHotThickness, HotSpacer, "Hot Side")
       Sic = InterfaceSalinityConcentration(Tavgc, Mavgc, Savgc, mM, A_f, PrLengthNormalToColdFlow, PrTotalLengthParallelToColdFlow, PrColdThickness, ColdSpacer, "Cold Side")
       
       'New routine to limit the Newtonian solution by the maximum amount of heat transfer possible
       C_min = mdlMath.DblMin(mH(0) * cph(0), mC(0) * cpc(0))
       Emax = C_min * (Th(0) - Tc(0))
       cpminh = mdlMath.DblMin(cph(0), cph(1))
       cpminc = mdlMath.DblMin(cpc(0), cpc(1))
       delThmax = Emax / (mH(0) * cpminh)
       delTcmax = Emax / (mC(0) * cpminc)
       
              
      If glbEvaluateEquationsAtEnd Then
          If IsFinalRun Then
             DebugString = ""
             C_star = mdlMath.DblMin(mH(0) * cph(0), mC(0) * cpc(0)) / mdlMath.DblMax(mH(0) * cph(0), mC(0) * cpc(0))
             DebugString = DebugString & FormatDebugColumns("C_hot:     Heat transfer Rate-H (W/K):     " & mH(0) * cph(0) & vbCrLf)
             DebugString = DebugString & FormatDebugColumns("C_cold:    Heat transfer Rate-C (W/K):   " & mC(0) * cpc(0) & vbCrLf)
             DebugString = DebugString & FormatDebugColumns("C_min:    Minimum Heat transfer Rate (W/K):   " & mdlMath.DblMin(mH(0) * cph(0), mC(0) * cpc(0)) & vbCrLf)
             DebugString = DebugString & FormatDebugColumns("C_max:     Minimum Heat transfer Rate (W/K):   " & mdlMath.DblMax(mH(0) * cph(0), mC(0) * cpc(0)) & vbCrLf)
             DebugString = DebugString & FormatDebugColumns("C_star:    Heat capacity ratio:    " & C_star & vbCrLf)
             DebugString = DebugString & FormatDebugColumns("Si_hot:   Interface Salinity hot side:    " & Sih & vbCrLf)
             DebugString = DebugString & FormatDebugColumns("Si_cold:  Interface Salinity cold side:   " & Sic & vbCrLf)
             DebugString = DebugString & FormatDebugColumns("Lhp:  HTEX Length Parallel to Hot Flow:   " & Me.LengthParallelToHotFlow & vbCrLf)
             DebugString = DebugString & FormatDebugColumns("Lhn:  HTEX Lnegth normal to Hot Flow: " & Me.LengthNormalToHotFlow & vbCrLf)
             DebugString = DebugString & FormatDebugColumns("WP_h:   Width of the hot flow passage:  " & Me.LengthNormalToHotFlow & vbCrLf)
             DebugString = DebugString & FormatDebugColumns("WP_c: Width of the cold flow passage: " & Me.LengthNormalToColdFlow & vbCrLf)
             Debug.Print DebugString
             mdlValidation.WriteDebugInfoToFile DebugString
          End If
       End If
       

       
       ' Move on to HTModel specific calculations
       Select Case HTModel
           Case 0 ' Direct contact MD
                ReDim LR(1 To 2) ' We are solving 2 local equations here for Tih, and Tic
                ReDim LX(1 To 2) ' , Tih, Tic
                ReDim LXmax(1 To 2) ' , Tih, Tic
                ReDim LXmin(1 To 2) ' , Tih, Tic
                
                
                           
                'Initialize solution vector
                LX(1) = Tih
                LX(2) = Tic
                
                LXmax(1) = Tavgh
                LXmax(2) = mdlMath.DblMin(Tavgh, Tavgc + delTcmax)
                
                LXmin(1) = mdlMath.DblMax(Tavgc, Tavgh - delThmax)
                LXmin(2) = Tavgc
                
                Set DCHT = New clsHTFuncDirectContact
                
                DCHT.ConvergenceCriteria ConvergCrit, A_f
                
                ' This only captures all of the constants already calculated.  These do not change
                DCHT.NeedToReadConstants = True ' YOU MUST RESET THIS - THIS WAS A HUGE BUG THAT KEPT ME GUESSING FOR A LONG TIME
                DCHT.EvaluateFunction LX, Me, SysEq, LR, Tavgc, Tavgh, Mavgc, Mavgh, mM, Tic, Tih, Sic, Sih, Savgc, Savgh
                
                ' Solve for the interface temperatures such that the heat transfer is constant through all
                ' three regions
                mdlMath.NewtonsMethod DCHT, LX, Me, SysEq, mdlConstants.glbNewtonConvergenceCriterion, mdlConstants.glbMaxNewtonIterations, _
                                      mdlConstants.glbDerivativeIncrement, LXmax, LXmin, ConvergCrit, True
                'mdlMath.ConjugateGradient DCHT, LX, Me, SysEq, 0.001, 1000, LXmax, LXmin
    
                'Reassign to global variables
                Tih = LX(1)
                Tic = LX(2)
                mM = DCHT.MembraneMassTransfer
                Lavg = mdlProperties.LatentHeatOfPureWater((Tih + Tic) / 2)
                
                Q = DCHT.HeatTransferRate 'Watts
                
               If glbEvaluateEquationsAtEnd Then
                  If IsFinalRun Then
                     DebugString = ""
                     NTU = DCHT.OverallHeatTransferCoefficient * Me.Area / mdlMath.DblMin(mH(0) * cph(0), mC(0) * cpc(0))
                     DebugString = DebugString & FormatDebugColumns("NTU   : Number of transfer units:     " & NTU & vbCrLf)
                     DebugString = DebugString & FormatDebugColumns("A : Control volume heat exchange area (m2):   " & Me.Area & vbCrLf)
                     DebugString = DebugString & FormatDebugColumns("U : Overall Heat Transfer Coefficient (W/m2/K):   " & DCHT.OverallHeatTransferCoefficient & vbCrLf)
                     DebugString = DebugString & FormatDebugColumns("E : HTEX (counter-flow) Effectiveness:    " & (1 - Exp(-NTU * (1 - C_star))) / (1 - C_star * Exp(-NTU * (1 - C_star))) & vbCrLf)
                     DebugString = DebugString & FormatDebugColumns("Th    : Hot side average temperature (K):     " & Tavgh & vbCrLf)
                     DebugString = DebugString & FormatDebugColumns("Tc    : Cold side average temperature (K):     " & Tavgc & vbCrLf)
                     DebugString = DebugString & "END Control Volume:" & Me.CVNumber & vbCrLf
                     Debug.Print DebugString
                     mdlValidation.WriteDebugInfoToFile DebugString
                  End If
               End If

         Case 1 ' Air gap membrane distillation
                AGMaxVapPressure = mdlProperties.SeaWaterVaporPressure(Tavgh, 0)

                ReDim LR(1 To 7) ' We are solving 7 local equations (See clsAirGapHTFunc)
                                 ' 5 heat flow equations,
                                 ' 1 mass flow equation (MD mass transfer versus diffussion across the air gap)
                                 ' 1 condensate film thickness

                ReDim LX(1 To 7) ' LX(1) - Tih temperature at the hot feed/membrane interface (Kelvin) - VARIABLE SHARED IN COMMON WITH DIRECT CONTACT
                                 ' LX(2) - Tma temperature at the membrane/air gap interface (Kelvin)
                                 ' LX(3) - Tf temperature at the air gap/condensate interface (Kelvin)
                                 ' LX(4) - Tfw temperature at the condensate/cooling wall foil interface (Kelvin)
                                 ' LX(5) - Tic temperature at the cooling wall foil/cooling water interface (Kelvin) - VARIABLE SHARED IN COMMON WITH DIRECT CONTACT
                                 ' LX(6) - delta_w average condensate thickness (condensate thickness variation is modeled as linear)
                                 ' LX(7) - Pv_ma water vapor partial pressure in moist air at the membrane/air gap interface (air is not necessarily saturated).
                ReDim LXmax(1 To 7) ' maximums for LX
                ReDim LXmin(1 To 7) ' minimums for LX

                'Initialize solution variable with either the initial guess or the previous solution
                LX(1) = Tih ' temperature at the hot feed/membrane interface (Kelvin)
                LX(2) = Tma ' temperature at the membrane/air gap interface (Kelvin)
                LX(3) = Tf  ' temperature at the air gap/condensate interface (Kelvin)
                LX(4) = Tfw ' temperature at the condensate/cooling wall foil interface (Kelvin)
                LX(5) = Tic ' temperature at the cooling wall foil/cooling water interface (Kelvin)
                LX(6) = delta_w ' average condensate thickness (condensate thickness variation is modeled as linear)
                LX(7) = Pv_ma / (1 + X_ma) ' water vapor partial pressure in moist air at the membrane/air gap interface (air is not necessarily saturated i.e. X_ma = 0).
                
                ' These bounds have to be re-established every step since the global solution is iterating.
                LXmax(1) = Tavgh
                LXmax(2) = Tavgh
                LXmax(3) = Tavgh
                LXmax(4) = Tavgh
                LXmax(5) = Tavgh
                LXmax(6) = SysEq.AirGapSpacer.Thickness ' The Condensate thickness cannot exceed the entire air gap
                LXmax(7) = AGMaxVapPressure 'SysEq.Inputs.AirGapPressure ' The vapor pressure cannot exceed the total pressure!
                                                                         'In fact, it cannot exceed the saturation pressure at Tma but that has to be enforced
                                                                         'through a new equation (#8) and new slack variable SV1!
                'X(7) = Pv_ma / ( 1+ X) the max of Pv_ma is Psat(Tavgh) and the max of X is 1 this produces a maximum value of Psat(Tavgh)
                LXmin(1) = Tavgc
                LXmin(2) = Tavgc
                LXmin(3) = Tavgc
                LXmin(4) = Tavgc
                LXmin(5) = Tavgc
                LXmin(6) = 0# ' Never expect zero but this is the lower bound!
                LXmin(7) = 0# ' Never expect zero but this is the lower bound!
                
                ' Added for Air gap only... The Maximum limit for the
                Set AGHT = New clsHTFuncAirGap
                
                ' Added 7/12/2017
                AGHT.ConvergenceCriteria ConvergCrit, A_f
                
                ' This only captures all of the constants already calculated.  These do not change
                AGHT.NeedToReadConstants = True ' YOU MUST RESET THIS - THIS WAS A HUGE BUG THAT KEPT ME GUESSING FOR A LONG TIME
                AGHT.EvaluateFunction LX, Me, SysEq, LR, Tavgc, Tavgh, Mavgc, Mavgh, Sic, Sih, Savgc, Savgh

                ' Solve for the interface temperatures such that the heat transfer is constant through all
                ' three regions - THIS HAS CONSTRAINTS ENABLES WHILE DirectContact doesn't need constraints
                mdlMath.NewtonsMethod AGHT, LX, Me, SysEq, mdlConstants.glbNewtonConvergenceCriterion, _
                        mdlConstants.glbMaxNewtonIterations * 3, mdlConstants.glbDerivativeIncrement, LXmax, LXmin, ConvergCrit, _
                        EvaluateConstraints:=False
                'Reassign to local variables
                Tih = LX(1) ' temperature at the hot feed/membrane interface (Kelvin)
                Tma = LX(2) ' temperature at the membrane/air gap interface (Kelvin)
                Tf = LX(3)  ' temperature at the air gap/condensate interface (Kelvin)
                Tfw = LX(4) ' temperature at the condensate/cooling wall foil interface (Kelvin)
                Tic = LX(5) ' temperature at the cooling wall foil/cooling water interface (Kelvin)
                delta_w = LX(6) ' average condensate thickness (condensate thickness variation is modeled as linear)
                AGSatVapPressure = mdlProperties.SeaWaterVaporPressure(Tma, 0)
                If LX(7) > AGSatVapPressure / 2 Then
                   Pv_ma = AGSatVapPressure
                   X_ma = AGSatVapPressure / LX(7) - 1
                Else
                   Pv_ma = 2 * LX(7)
                   X_ma = 1
                End If

                'Items used in the global heat transfer equations that need to be updated now that we have an equillibrated local solution
                mM = AGHT.MembraneMassTransfer
                Q = AGHT.HeatTransferRate 'Watts
                delta_condensate_exit = AGHT.CondensateExitThickness
                Lavg = mdlProperties.LatentHeatOfPureWater((Tih + Tma) / 2)
                'Recheck control volume pair for an accurate configuration
                'Bool = Me.ViolatesSecondLawOfThermodynamics
       End Select
       
       ' Added 7/17/2017 - Include losses because they are significant in the lab testing
       Select Case LossesConfig
          Case 0 ' No losses, do nothing
              Q_loss_hot = 0
              Q_loss_cold = 0
          Case 1 ' only hot side is losing heat to external
              Q_loss_hot = Q_losses(Tavgh, Mavgh, Savgh, Sih, True, SysEq)
              Q_loss_cold = 0
          Case 2 ' only cold side is losing heat to external
              Q_loss_hot = 0
              Q_loss_cold = Q_losses(Tavgc, Mavgc, Savgc, Sic, False, SysEq)
          Case 3 ' both sides are losing heat to external (only happens when there is only one layer and both sides are exposed)
              Q_loss_hot = Q_losses(Tavgh, Mavgh, Savgh, Sih, True, SysEq)
              Q_loss_cold = Q_losses(Tavgc, Mavgc, Savgc, Sic, False, SysEq)
       End Select
       

       
       PrCalcComplete = True
   End If
EndOfSub:
    
Exit Sub
ErrorOccured:
    If mdlConstants.glbIncludeMsgBox Then
        MsgBox err.Source, vbCritical, "clsControlVolumePair.CalculateControlVolumePair"
    End If
GoTo EndOfSub
End Sub

Public Function Q_losses(T_water As Double, M_water As Double, S_water As Double, S_iw As Double, _
                         IsHotSide As Boolean, SysEq As clsSystemEquations) As Double
    Dim LHT As clsHTFuncLosses
    Dim LR() As Double
    Dim LX() As Double
    Dim LXmax() As Double
    Dim LXmin() As Double
    Dim TempJump As Double
    Dim TempMax As Double
    Dim TempMin As Double
    Dim ConvergCrit() As Double
    
    Set LHT = New clsHTFuncLosses
    
    
    ReDim LR(1 To 2) ' We are solving 2 local equations here for T_ia, and T_iw
    ReDim LX(1 To 2) ' , Tih, Tic
    ReDim LXmax(1 To 2) ' , Tih, Tic
    ReDim LXmin(1 To 2) ' , Tih, Tic
    
    'Initialize solution vector
    TempJump = mdlConstants.glbInitialFractionToColdOrHotForInterfaceTemperature * (T_water - T_Air)
    LX(1) = T_water - TempJump
    LX(2) = T_Air + TempJump
    
    TempMax = mdlMath.DblMax(T_water, T_Air)
    TempMin = mdlMath.DblMin(T_water, T_Air)
    
    LXmax(1) = TempMax
    LXmax(2) = TempMax
    
    LXmin(1) = TempMin
    LXmin(2) = TempMin
    
    LHT.ConvergenceCriteria ConvergCrit, A_f
    
    ' This only captures all of the constants already calculated.  These do not change
    LHT.NeedToReadConstants = True ' YOU MUST RESET THIS - THIS WAS A HUGE BUG THAT KEPT ME GUESSING FOR A LONG TIME
    LHT.EvaluateFunction LX, Me, SysEq, LR, IsHotSide, T_Air, T_water, M_water, S_water, S_iw
    
    ' Solve for the interface temperatures such that the heat transfer is constant through all
    ' three regions
    mdlMath.NewtonsMethod LHT, LX, Me, SysEq, mdlConstants.glbNewtonConvergenceCriterion, mdlConstants.glbMaxNewtonIterations, _
                          mdlConstants.glbDerivativeIncrement, LXmax, LXmin, ConvergCrit

    Q_losses = LHT.HeatTransferRate 'Watts
End Function

Public Property Get MembraneAverageTemperature() As Double
   ' Use the interface temperatures
   MembraneAverageTemperature = (Tih + Tic) / 2
End Property

Public Property Get AverageSalinity() As Double
       Savgh = (2 * mH(0) / (mH(0) + mH(1))) * SH(0)
       Savgc = (2 * mC(0) / (mC(0) + mC(1))) * SC(0)
       AverageSalinity = (Savgh + Savgc)
End Property

Public Property Get MembraneAverageDensity() As Double
   MembraneAverageDensity = mdlProperties.SeaWaterDensity(Me.MembraneAverageTemperature, 0)
End Property

Public Property Get MassFlow() As Double
    MassFlow = mM
End Property


Private Sub Class_Initialize()
    ReDim cpc(0 To 1)
    ReDim cph(0 To 1)
End Sub


