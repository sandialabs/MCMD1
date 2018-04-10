VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHTFuncDirectContact"
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



Private PrNeedToReadConstants As Boolean
Private IsFirstTimeThrough As Boolean

Private T_avgc As Double 'Average temperature of bulk stream flow on cold side (K)
Private T_avgh As Double 'Average temperature of bulk stream flow on hot side (K)
Private M_avgc As Double 'Average mass flow on cold side (kg/s)
Private M_avgh As Double 'Average mass flow on hot side (kg/s)
Private m_M As Double ' Membrane distilation mass flow
Private T_ih As Double 'Interface temperature (temperature at membrane wall) on the hot side (K)
Private T_ic As Double 'Interface temperature (temperature at membrane wall) on the cold side (K)
Private S_ic As Double 'Interface Salinity cold side (grams/gram)
Private S_ih As Double 'Interface Salinity hot side (grams/gram)
Private P_c As Double 'Pressure cold side (Pa)
Private P_h As Double 'PRessure hot side (Pa)
Private U As Double  'Overall heat transfer coefficient (W/(m2*K))
Private A_f As Double ' mass transfer area
Private Q As Double ' heat transfer rate via effectiveness method
Private cp_hi As Double ' specific heat hot side inflow
Private cp_ci As Double ' specific heat cold side inflow
Private FlowType As Long ' 1 = coflow, 2 = counter flow, 3 = cross-flow
Private L_Norm_c As Double ' Length normal to flow cold side (m)
Private Lzc As Double ' thickness of cold side (from cold side spacers) (m)
Private L_TotalPara_c As Double ' Length parallel to flow cold side (m)
Private L_Norm_h As Double ' Length normal to flow hot side (m)
Private Lzh As Double ' thickness of hot side (from hot side spacers) (m)
Private L_TotalPara_h As Double ' Length parallel to flow hot side (m)
Private Phi As Double ' Membrane porosity
Private k_mf As Double ' Membrane thermal conductivity
Private Lz_m As Double ' Membrane thickness
Private r_mean As Double 'Membrane mean pore thickness
Private K_m As Double 'Membrane manually input mass transfer coefficient.
Private Use_K_m As Boolean ' Boolean to indicate if we should use the mass transfer coefficient.
Private Hh As Double
Private Hm As Double
Private Hc As Double
Private HTModel As Long
Private S_avgc As Double
Private S_avgh As Double

Private HSpacer As clsSpacer ' Hot and cold spacers
Private CSpacer As clsSpacer

Public Property Let NeedToReadConstants(Val As Boolean)
    PrNeedToReadConstants = Val
End Property

Public Property Get OverallHeatTransferCoefficient() As Double
    OverallHeatTransferCoefficient = U
End Property

Public Property Get MembraneMassTransfer() As Double
    MembraneMassTransfer = m_M
End Property

Public Property Get HeatTransferRate() As Double
    HeatTransferRate = Q
End Property

Public Property Get HeatTransferIsBalanced() As Boolean
    Dim Tol As Double
    
    If Q <> 0 Then
       If mdlMath.AbsDbl(((Hm * A_f * (T_ih - T_ic) - Q) / Q)) < 0.01 Then
          HeatTransferIsBalanced = True
       Else
          HeatTransferIsBalanced = False
       End If
    Else
       HeatTransferIsBalanced = False
    End If
    
End Property

Public Sub EvaluateConstraints(x() As Double, xp1() As Double, MaxReduction As Double) 'Return a step reduction size factor to keep from violating constraints

' This constraints are specific to direct contact.  The constraint is x(1) >= x(2)
'
Dim c As Double

  
  If (x(1) + xp1(1) <= x(2) + xp1(2)) Then
    ' We have a constraint violation
    c = (x(1) - x(2)) / (xp1(2) - xp1(1))
    If c > 1 Or c < 0 Then
       Debug.Assert False
    End If
    xp1(1) = MaxReduction * c * xp1(1)
    xp1(2) = MaxReduction * c * xp1(2)
    
    
  End If
End Sub


Private Sub AdjustInterfaceTemperaturesIfInvalid(Optional AdjustmentMade As Boolean)
' This function is used to adjust the interface temperatures if they
' have physically inadmissible values after adjustment of the global
' equation set. A large jump in the global solution may make the previous
' interface temperatures completely incorrect. This procedure has
' made the model much more robust for heat exchangers that come close
' to saturation.
' THIS FUNCTION HAS BEEN TESTED THROUGH ALL OF ITS IF-THEN LINES.
Dim Dum As Double
Dim Tdif As Double
Dim Frac_h As Double
Dim Frac_c As Double
Const Big = 0.25
Const Small = 0.15

AdjustmentMade = False
'The following order of temperatures is required for the 2nd law of thermodynamics to be
'satisfied (from hottest to coldest)
' 1. T_avgh
' 2. T_ih
' 3. T_ic
' 4. T_avgc

' If this is not the case, force it!!!!

' Test cases (keep commented out unless you change this function)!


If T_avgh < T_avgc Then
   Dum = mdlError.ReturnError("clsHTFuncDirectContact.AdjustInterfaceTemperaturesIfInvalid: The average hot temperature " & _
                              " and average cold temperature of the global solution are inverted! no way to fix this at the local level!")
    
Else
   Tdif = T_avgh - T_avgc
   'First assire that T_ic and T_ih are inbetween T_avgh and T_avgc
   ' The entire objective is to return to a physically admissible solution
   If T_ih > T_avgh Then
      T_ih = T_avgh - Small * Tdif
      AdjustmentMade = True
   End If
   If T_ic > T_avgh Then
      T_ic = T_avgh - Big * Tdif
      AdjustmentMade = True
   End If
   If T_ih < T_avgc Then
      T_ih = T_avgc + Big * Tdif
      AdjustmentMade = True
   End If
   If T_ic < T_avgc Then
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


' This function name must be shared by all objects
' It represents the set of nonlinear local equations that must be solved
' In order to resolve terms in the global energy and mass transfer equations.
' There can be many different setups. In this case, the solution could be accomplished through
' successive substitution since a solution for the interface temperatures is straightforward but
' I used Newton's method anyway to keep the methodology generalized and to test the methods on a
' very simple case before moving to more complicated methods.
Public Sub EvaluateFunction(x() As Double, CV As clsControlVolumePair, SysEq As clsSystemEquations, Result() As Double, _
                                               Optional Tavgc As Double, _
                                               Optional Tavgh As Double, _
                                               Optional Mavgc As Double, _
                                               Optional Mavgh As Double, _
                                               Optional mM As Double, _
                                               Optional Tic As Double, _
                                               Optional Tih As Double, _
                                               Optional Sic As Double, _
                                               Optional Sih As Double, _
                                               Optional Savgc As Double, _
                                               Optional Savgh As Double)
' Use optional variables to set values to provide a means of initialization.
    Dim Dum As Double
    Dim DebugString As String
    
    If PrNeedToReadConstants Then
      If (IsMissing(Tavgc) Or IsMissing(Tavgh) Or IsMissing(Mavgc) Or IsMissing(Mavgh) Or _
         IsMissing(mM) Or IsMissing(Tic) Or IsMissing(Tih) Or IsMissing(Sic) Or IsMissing(Sih)) Then
            Dum = mdlError.ReturnError("clsHTFuncDirectContact.EvaluateHeatTransfer: All optional arguments must be included if this is the first time that the class is being called. " _
                                      & " Subsequent calls do not need the optional arguments.", , True)
      Else
        ' Initialize values (some of these change and others stay the same value)
         T_avgc = Tavgc ' cold average temperature
         T_avgh = Tavgh ' hot average temperature
         M_avgc = Mavgc ' cold average mass flow
         M_avgh = Mavgh ' hot average mass flow
         m_M = mM  ' membrane distillation mass flow
         T_ih = Tih ' hot interface temperature
         T_ic = Tic ' cold interface temperature
         
         ' Added 1/9/2018 to compensate for when a global solution
         '                jump makes the previous interface temperatures
         '                physically unrealistic.
         AdjustInterfaceTemperaturesIfInvalid
         
         S_ic = Sic ' cold side interface salinity
         S_ih = Sih ' hot side interface salinity
         S_avgc = Savgc
         S_avgh = Savgh
         cp_hi = CV.SpecificHeats(True, True)
         cp_ci = CV.SpecificHeats(False, True)
         
         
         ' These are values that NEVER change and take more time to extract over and over again
         If IsFirstTimeThrough Then
            
            A_f = CV.Area
            L_Norm_c = CV.LengthNormalToColdFlow
            Lzc = CV.ColdThickness
            L_TotalPara_c = CV.TotalLengthParrallelToColdFlow
            L_Norm_h = CV.LengthNormalToHotFlow
            Lzh = CV.HotThickness
            L_TotalPara_h = CV.TotalLengthParrallelToHotFlow
         
            Dim MembMat As clsMaterial
            Set MembMat = SysEq.Inputs.Materials(SysEq.Inputs.MembraneMaterial)
            Phi = MembMat.Porosity
            k_mf = MembMat.ThermalConductivity
            Lz_m = MembMat.Thickness
            r_mean = MembMat.MeanPoreRadius
            K_m = MembMat.DirectMassTransferCoefficient
            Use_K_m = MembMat.UseDirectInputForMassTransferCoefficient
         
            Dim HWS As clsWaterStream
            Dim CWS As clsWaterStream
            
            '!@#$Eventually pressures may change and become a variable in the
            'solution set.  For now they are constant!
            Set HWS = SysEq.Inputs.WaterStreams(SysEq.Inputs.HotWaterStream)
            Set CWS = SysEq.Inputs.WaterStreams(SysEq.Inputs.ColdWaterStream)
            
            P_c = CWS.Pressure
            P_h = HWS.Pressure
            
            Set HSpacer = SysEq.HotSpacer
            Set CSpacer = SysEq.ColdSpacer
            
            FlowType = CV.FlowType
            HTModel = CV.HeatTransferModel
         
          End If
         

         IsFirstTimeThrough = False ' This never gets reset
         PrNeedToReadConstants = False ' This gets reset for each solution iteration
      End If
    Else
       ' Only the variables have to be updated.  Everything else is constant or gets recalculated.
        Dim L_avg As Double
        T_ih = x(1)
        T_ic = x(2)

        L_avg = mdlProperties.LatentHeatOfPureWater((T_ih + T_ic) / 2)
                
        'Calculate the Mass transfer
        m_M = MembraneDistillationMassTransfer(T_ih, S_ih, P_h, T_ic, S_ic, P_c, A_f, Phi, r_mean, Lz_m, K_m, Use_K_m)
        
        'Overall heat transfer Coefficient (not used anymore - we use the local heat transfer coefficients
        U = mdlTransferCoefficient.HTC(T_avgc, T_avgh, M_avgc, M_avgh, L_avg, m_M, T_ih, T_ic, Hc, Hh, Hm, _
                                       L_Norm_c, CSpacer, Lzc, L_TotalPara_c, L_Norm_h, HSpacer, Lzh, L_TotalPara_h, A_f, _
                                        Phi, k_mf, Lz_m, S_ic, S_ih, S_avgc, S_avgh, Use_K_m, HTModel)
        
        
        
        ' THIS IS NOT NEEDED BECAUSE WE ARE MODELING THE GEOMETRY OF THE FLOW IN THE HEAT EXCHANGER ALREADY AND HAVE A LOCAL 1-D HEAT TRANSFER
        ' MODEL OF THE TRAVERSE HEAT FLOW THAT IS BEING SOLVED VIA NEWTON'S METHOD
        'Q = mdlTransferCoefficient.EffectivenessHeatTransferRate(FlowType, T_avgh, T_avgc, M_avgh, M_avgc, cp_hi, cp_ci, U, A_f, True, True)
        
        ReDim Result(1 To 2)
        ' Calculate the total heat transfer several ways these are the equations that need to be driven to zero if the interface temperatures are consistent
        Result(1) = Hh * (T_avgh - T_ih) - Hm * (T_ih - T_ic)
        Result(2) = Hc * (T_ic - T_avgc) - Hm * (T_ih - T_ic)
        
        Q = U * A_f * (T_avgh - T_avgc)
        
        If glbEvaluateEquationsAtEnd Then
           If IsFinalRun And IsFinalNewtonIteration Then
               
               DebugString = DebugString & FormatDebugColumns("Hv: H2O Latent heat at avg. interface temp (J/kg):      " & L_avg & vbCrLf)
               DebugString = DebugString & FormatDebugColumns("T_ave: Average Temperature of hot and cold bulk streams (K): " & (T_avgh + T_avgc) / 2 & vbCrLf)
               DebugString = DebugString & FormatDebugColumns("q: Transverse Heat Transfer (W): " & Q & vbCrLf)
               DebugString = DebugString & FormatDebugColumns("dist_mass: Mass flow of the distilate (kg/s): " & m_M & vbCrLf)
               DebugString = DebugString & FormatDebugColumns("Hm: Membrane heat transfer coefficient (W/m2/K): " & Hm & vbCrLf)
               DebugString = DebugString & FormatDebugColumns("Tci: Membrane to cold flow interface temperature (K):   " & T_ic & vbCrLf)
               DebugString = DebugString & FormatDebugColumns("Thi: Membrane to hot flow interface temperature (K):    " & T_ih & vbCrLf)
               Debug.Print DebugString
               mdlValidation.WriteDebugInfoToFile DebugString

           End If
        End If
    End If
End Sub

Private Sub Class_Initialize()
   PrNeedToReadConstants = True
   IsFirstTimeThrough = True
End Sub


Public Sub ConvergenceCriteria(ConvergCrit() As Double, CVArea As Double)

ReDim ConvergCrit(1 To 2)
Dim EnergyCritOverArea As Double

EnergyCritOverArea = mdlConstants.glbAbsConvergCriterionEnergy / CVArea

ConvergCrit(1) = EnergyCritOverArea
ConvergCrit(2) = EnergyCritOverArea

End Sub


''THIS FUNCTION IS NOT USED ANYMORE!!! It is being retained if we ever want to return to
'' a more direct method for direct contact MD heat transfer.
'Private Sub DirectContactInterfacialTemperatures(SysEqn As clsSystemEquations, Con As clsConnectivity, CVNum As Long, _
'                                   Hh As Double, Hc As Double, Hm As Double, HotInterfaceTemperature As Double, _
'                                   ColdInterfaceTemperature As Double)
'
'Dim Hv As Double
'Dim Var As Variant
'Dim Th_in As Double
'Dim Th_out As Double
'Dim Tc_in As Double
'Dim Tc_out As Double
'Dim Th As Double
'Dim Tc As Double
'
'
'   Th_in = SysEqn.Temperature(Con.c(CVNum, Con.NodeOrder(1, CVNum)))
'   Th_out = SysEqn.Temperature(Con.c(CVNum, Con.NodeOrder(2, CVNum)))
'   Tc_in = SysEqn.Temperature(Con.c(CVNum, Con.NodeOrder(3, CVNum)))
'   Tc_out = SysEqn.Temperature(Con.c(CVNum, Con.NodeOrder(4, CVNum)))
'
'   Th = (Th_in + Th_out) / 2
'   Tc = (Tc_in + Tc_out) / 2
'
'   ' These equations come from Hitsov, I., Maere, T., Sitter, K., Dotremont, C., Nopens, I.,
'   '                           "Modelling approaches in membrane distillation: A critical review"
'   '                           Separation and Purification Technology, 142 (2015) 48-64 equations 6 and 7
'   ' Hh hot side heat transfer coefficient (W/K)
'   ' Hm membrane heat transfer coefficient (W/K)
'   ' Hc cold side heat transfer coefficient (W/K)
'   ' Tc - cold side temperature
'   ' Th - hot side temperature
''OLD - Hm did not contain the latent heat of evaporation in the original equations which messes up the heat transfer elsewhere.
''   It is best to include Hv * m_M in Hm and to modify these equations - this was a major error in the heat transfer before 9/13/2016
'   ' hot = 2
''   T_i(i, 2) = (Th * Hh(i) + _
''                  Hm(i) * (Tc + Th * (Hh(i) / Hc(i))) - _
''                  Hv * m_M(i)) / _
''                  (Hh(i) * (1 + (Hm(i) / Hc(i))) + Hm(i))
''   ' cold = 1
''   T_i(i, 1) = (Tc * Hc(i) + _
''                   Hm(i) * (Th + Tc * (Hc(i) / Hh(i))) + _
''                   Hv * m_M(i)) / _
''                   (Hc(i) * (1 + (Hm(i) / Hh(i))) + Hm(i))
'' NEW ' Hm contains the latent heat of evaporation already!!
'
'   HotInterfaceTemperature = (Th * Hh + Hm * (Tc + Th * (Hh / Hc))) / (Hh * (1 + (Hm / Hc)) + Hm)
'
'   ColdInterfaceTemperature = (Tc * Hc + Hm * (Th + Tc * (Hc / Hh))) / (Hc * (1 + (Hm / Hh)) + Hm)
'
'   If ColdInterfaceTemperature > HotInterfaceTemperature Then
'        HotInterfaceTemperature = mdlError.ReturnError("mdlDirectContact.DirectContactInterfacialTemperatures: The Cold interface temperature" & _
'                                 "has been calculated to be hotter than the hot interface temperature. Investigation is needed", , True, True)
'   End If
'
'End Sub
