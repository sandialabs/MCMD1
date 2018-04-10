VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHTFuncAirGap"
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

' The air gap heat transfer function has to be solved in a specific order to be able to properly resolve the fluid film thickness properly.
' The fluid film thickness is a function of all previous mass flows upward
'
' This model comes mostly from the derivation of Alsaadi et. al., 2013 "Modeling of Air Gap Membrane Distillation Process: A theoretical and experimental study" Journal of Membrane Science.

Private PrNeedToReadConstants As Boolean
Private IsFirstTimeThrough As Boolean

Private Tcb As Double 'Average temperature of bulk stream flow on cold side (K) =
Private Thb As Double 'Average temperature of bulk stream flow on hot side (K)
Private M_avgc As Double 'Average mass flow on cold side (kg/s)
Private M_avgh As Double 'Average mass flow on hot side (kg/s)
Private T_m_M_in As Double ' This is the temperature of
Private T_m_M_out As Double ' After variables have been resolved, this is the

Private delta_w_in As Double 'thickness of the condensate fluid at input
Private delta_w_out As Double ' thickness of the condensate fluid at output

'Unknown variables of the air - gap analysis
Private Pv_ma As Double ' Interface vapor pressure exiting the membrane (not saturated air)
Private Tih As Double 'Interface temperature (temperature at membrane wall) on the hot side (K)
Private Tma As Double 'Interface temperature (temperature at membrane wall) of the air gap (K)
Private Tf As Double 'Saturated water temperature at the condensing fluid-air-gap interface (K)
Private Tfw As Double 'Subcooled water temperature at the condensing fluid to cooling wall foil interface (K)
Private Tic As Double 'Cooling fluid to foil interface temperature
Private delta_w As Double 'thickness of the condensate fluid (average)
Private X_ma As Double '

Private m_M As Double 'Membrane distilation mass flow - determines the thickness of the condensate flow based on film condensation theory.
Private m_D As Double 'Mass flow due to diffusion
Private S_ic As Double 'Interface Salinity cold side (grams/gram)
Private S_ih As Double 'Interface Salinity hot side (grams/gram)
Private S_avgc As Double ' Average Bulk Salinity (cold)
Private S_avgh As Double ' Average Bulk salinity (hot)
Private P_c As Double 'Pressure cold side (Pa)
Private P_h As Double 'Pressure hot side (Pa)
Private U As Double  'Overall heat transfer coefficient (W/(m2*K))
Private A_f As Double ' mass transfer area
Private Q As Double ' heat transfer rate
Private cp_hi As Double ' specific heat hot side inflow
Private cp_ci As Double ' specific heat cold side inflow
Private FlowType As Long ' 1 = coflow, 2 = counter flow, 3 = cross-flow
Private L_Norm_c As Double ' Length normal to flow cold side (m)
Private Lzc As Double ' thickness of cold side (from cold side spacers) (m)
Private L_TotalPara_c As Double ' Length parallel to flow cold side (m)
Private L_Norm_h As Double ' Length normal to flow hot side (m)
Private Lzh As Double ' thickness of hot side (from hot side spacers) (m)
Private L_TotalPara_h As Double ' Length parallel to flow hot side (m)
Private L_Para_g As Double ' Length parallel to gravity (m)
Private Phi As Double ' Membrane porosity
Private k_mf As Double ' Membrane thermal conductivity W/(m*K) (this is of the matrial without accounting for the resistive
                       '    effect of the air if Use_K_m is true. Otherwise it is the total thermal conductivity of the membrane
Private k_foil As Double ' cooling wall foil thickness W/(m*K)
Private Lz_m As Double ' Membrane thickness (m)
Private Lz_a As Double ' Air gap thickness (air + condensate film thickness) (m)
Private Lz_f As Double ' Foil (cooling wall) thickness (m)
Private r_mean As Double 'Membrane mean pore thickness
Private K_m As Double 'Membrane manually input mass transfer coefficient.
Private Use_K_m As Boolean ' Boolean to indicate if we should use the mass transfer coefficient.
Private Hh As Double ' hot side heat transfer coefficient W/(m2*K)
Private Hm As Double ' membrane heat transfer coefficient W/(m2*K)
Private Hc As Double ' cold side heat transfer coefficient W/(m2*K)
Private Ha As Double ' heat transfer coefficient for air gap - air W/(m2*K)
Private Hw As Double ' heat transfer coefficient for cooling wall - W/(m2*K)
Private Hf As Double ' heat transfer coefficient for condensate fluid in air gap - W/(m2*K)

Private HTModel As Long  'Heat transfer type - 0 direct heat transfer, 1 air gap heat transfer
Private P_air_gap As Double ' pressure in the air gap

Private HSpacer As clsSpacer ' Hot and cold spacers
Private CSpacer As clsSpacer
Private ASpacer As clsSpacer ' Air gap spacer

Public Property Let NeedToReadConstants(Val As Boolean)
    PrNeedToReadConstants = Val
End Property

Public Property Get MembraneMassTransfer() As Double
    MembraneMassTransfer = m_M
End Property

Public Property Get HeatTransferRate() As Double
    HeatTransferRate = Q
End Property

Public Property Get CondensateExitThickness() As Double
    CondensateExitThickness = delta_w_out
End Property

'Public Property Get HeatTransferIsBalanced() As Boolean
'    Dim Tol As Double
'
'    If Q <> 0 Then
'       If mdlMath.AbsDbl(((Hm * A_f * (Tih - Ta) - Q) / Q)) < 0.01 Then
'          HeatTransferIsBalanced = True
'       Else
'          HeatTransferIsBalanced = False
'       End If
'    Else
'       HeatTransferIsBalanced = False
'    End If
'
'End Property


' This function name must be shared by all objects
' It represents the set of nonlinear local equations that must be solved
' In order to resolve terms in the global energy and mass transfer equations.
' There can be many different setups. In this case, the solution requires 6 variables to be
' solved:

' 1. Tih - hot side membrane interface temperature
' 2. Tma  - air gap/membrane interface temperature
' 3. Tf   - temperature of saturated fluid at the condensate fluid/air interface
' 4. Tfw  - temperature of the sub-cooled fluid at the condensate cooling wall foil interface
' 5. Tic  - temperature of subcooled cooling fluid at the cooling wall foil bulk cooling flow interface
' 6. delta_w - average condensate film thickness
' 7. Pv_ma   - vapor pressure of the exit of the membrane into the air gap.
' - condensation film thickness ( in the equations, a linear slope is assumed whose start point is the thickness required to support
'              all of the previous mass flow traveling in the direction of gravity.
'
Public Sub EvaluateConstraints(x() As Double, xp1() As Double, MaxReduction As Double) 'Return a step reduction size factor to keep from violating constraints

' These constraints are specific to air gap.  The constraint is x(7) [Pv_ma] <= Psat(x(2) [Tma])
'
'
Dim Psat2 As Double
Dim Psat1 As Double
Dim Psat As Double
Dim Pv_ma_ As Double
Dim Pv_ma_2 As Double
Dim Pv_ma_1 As Double
Dim X_ma As Double 'Quality at membrane air gap interface.  If the pressure reaches saturation this becomes non-zero and changes the heat balances
Dim Iter As Long
Dim TopMult As Double
Dim BotMult As Double
Dim MidMult As Double
Dim OnBoundary As Boolean
Dim Tol As Double
Dim MaxIter As Long
Dim ReducFact As Double
Dim i As Long
Dim Dum As Double


   
  Psat2 = mdlProperties.SaturatedPressurePureWater(x(2) + xp1(2))
  Pv_ma_2 = x(7) + xp1(7)
  
  If Psat2 < Pv_ma_2 Then ' The constraint is NOT satisfied and the step will have to be reduced!
     
      Pv_ma_1 = x(7)
       
      Psat1 = mdlProperties.SaturatedPressurePureWater(x(2))
      
      
      If Pv_ma_1 > Psat1 Then
         Dum = mdlError.ReturnError("clsHTFuncAirGap.EvaluateConstraints: The partial pressure of water is beyond the saturation pressure at the corresponding temperature." & _
                              " There is a problem with your Air gap algorithm or you have used a solver that doesn't handle the constraint correctly!", , True)
      End If
      
      TopMult = 1#
      BotMult = 0#
      Iter = 0
      MaxIter = 100
      Tol = 1 'Pascals
       
      ' Do a binary search
      OnBoundary = False
      Do While Not OnBoundary And Iter < MaxIter 'We will want to reduce the step size (but not direction) until the constraint is just satisfied
         MidMult = (TopMult + BotMult) / 2
         
         Pv_ma_ = x(7) + MidMult * xp1(7)
         Psat = mdlProperties.SaturatedPressurePureWater(x(2) + MidMult * xp1(2))
         
         If Pv_ma_ > Psat Then ' We are still violating the boundary and we need to reduce the Multiplier
            TopMult = MidMult
         Else ' We are no longer violating the boundary but we may want to get closer to it
            If ((TopMult - BotMult) * xp1(7)) < Tol Then ' We are done!
                OnBoundary = True
                ReducFact = MidMult
            Else
                BotMult = MidMult
            End If
         End If
                   
         Iter = Iter + 1
      Loop
      
      If ReducFact < MaxReduction Then ' Only reduce the current constraint and let everyone else only be slowed down by MaxReduction (this keeps the search from just stopping)
         For i = LBound(xp1) To UBound(xp1)
            If i = 7 Then
               xp1(7) = ReducFact * xp1(7)
               ' Watch the order! this only works because 7 is after 2!
               Psat = mdlProperties.SaturatedPressurePureWater(x(2) + xp1(2))
               If x(7) + xp1(7) > Psat Then 'We missed a moving target.  Go back to the boundary!
                  xp1(7) = Psat - x(7)
               End If
            Else
               xp1(i) = MaxReduction * xp1(i)
            End If
         Next i
      Else
         For i = LBound(xp1) To UBound(xp1)
            xp1(i) = ReducFact * xp1(i)
         Next
      End If

  End If

End Sub

Public Sub EvaluateFunction(x() As Double, CV As clsControlVolumePair, SysEq As clsSystemEquations, Result() As Double, _
                                               Optional T_cb As Double, _
                                               Optional T_hb As Double, _
                                               Optional Mavgc As Double, _
                                               Optional Mavgh As Double, _
                                               Optional Sic As Double, _
                                               Optional Sih As Double, _
                                               Optional Savgc As Double, _
                                               Optional Savgh As Double)
' Use optional variables to set values to provide a means of initialization.
    Dim Dum As Double
    If PrNeedToReadConstants Then
      If (IsMissing(T_cb) Or IsMissing(T_hb) Or IsMissing(Mavgc) Or IsMissing(Mavgh) Or _
          IsMissing(Sic) Or IsMissing(Sih)) Then
            Dum = mdlError.ReturnError("clsDirectContactHTFunc.EvaluateHeatTransfer: All optional arguments must be included if this is the first time that the class is being called. " _
                                      & " Subsequent calls do not need the optional arguments.", , True)
      Else
        ' Initialize values (the values below change with each solution cycle)
         Tcb = T_cb ' cold average temperature
         Thb = T_hb ' hot average temperature
         M_avgc = Mavgc ' cold average mass flow
         M_avgh = Mavgh ' hot average mass flow

         S_ic = Sic ' cold side interface salinity
         S_ih = Sih ' hot side interface salinity
         S_avgc = Savgc ' bulk flow salinities.
         S_avgh = Savgh
         cp_hi = CV.SpecificHeats(True, True)
         cp_ci = CV.SpecificHeats(False, True)
         ' delta_w_in is only up to date because the Control volume pairs are being calculated by gravity layers so that
         ' we can avoid turning the lateral heat transfer into a global set of equations (even though it might be best to do so - this would ruin my entire approach!).
         If CV.PreviousGravityLayerCVNumber = mdlConstants.glbINVALID_VALUE Then
             ' This means that there is NO previous CV number and:
             delta_w_in = 0
         Else
             delta_w_in = SysEq.ControlVolumePair(CV.PreviousGravityLayerCVNumber).CondensateExitThicknessAirGapOnly
         End If
         
         ' These are values that NEVER change and take more time to extract over and over again
         If IsFirstTimeThrough Then
            
            A_f = CV.Area
            L_Norm_c = CV.LengthNormalToColdFlow
            Lzc = CV.ColdThickness
            L_TotalPara_c = CV.TotalLengthParrallelToColdFlow
            L_Norm_h = CV.LengthNormalToHotFlow
            Lzh = CV.HotThickness
            L_TotalPara_h = CV.TotalLengthParrallelToHotFlow
            L_Para_g = CV.LengthParallelToGravity
            ' This is the TOTAL thickness including the condensate thickness
            Lz_a = CV.AirGapThickness
         
            Dim MembMat As clsMaterial
            Set MembMat = SysEq.Inputs.Materials(SysEq.Inputs.MembraneMaterial)
            Phi = MembMat.Porosity
            k_mf = MembMat.ThermalConductivity
            Lz_m = MembMat.Thickness
            r_mean = MembMat.MeanPoreRadius
            K_m = MembMat.DirectMassTransferCoefficient
            Use_K_m = MembMat.UseDirectInputForMassTransferCoefficient
            
            Dim FoilMat As clsMaterial
            Set FoilMat = SysEq.Inputs.Materials(SysEq.Inputs.FoilMaterial)
            Lz_f = FoilMat.Thickness
            k_foil = FoilMat.ThermalConductivity
         
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
            Set ASpacer = SysEq.AirGapSpacer
            
            P_air_gap = SysEq.Inputs.AirGapPressure
            
            FlowType = CV.FlowType
            HTModel = CV.HeatTransferModel ' 0 = direct contact, 1 = air gap
            
          End If

         IsFirstTimeThrough = False ' This never gets reset
         PrNeedToReadConstants = False ' This gets reset for each solution iteration
      End If
    Else
        Dim AirOnlyThickness As Double
        Dim D_ab As Double 'Air-Water diffusion coefficient
        Dim D_ab2 As Double ' Air -Water diffusion coefficient times total air gap pressure NOT THE SAME UNITS as D_ab
        Dim cM As Double ' Molar concentration
        Dim yma As Double ' Mole fraction of water at Tma
        Dim yf As Double ' Mole fraction of water at Tf
        Dim yavg As Double ' average molde fraction between Tma and Tf
        Dim Mav As Double ' Molecular weight of air-water mixture
        Dim Mv As Double ' Molecular weight of vapor (is a global constant but included here to reduce the length of
                         ' equations
        Dim A_slope As Double ' slope of the condensate film thickness rate of change delta = A_slope * xx + delta_w_in where xx ranges from 0 to L_Para_g
                              ' delta_w = 0.5 * (delta_w_in + (A_slope * L_Para_g + delta_w_in))
        Dim B0 As Double ' a constant from equation 24 of A.S. Alsaadi
        'Calculate relevant parameters
        Dim rho_liq As Double
        Dim rho_gas As Double
        Dim mu As Double ' dynamic viscosity of the water condensed on the cooling wall
        Dim Tf_avg As Double ' average temperature in the condensate on the cooling wall
        Dim rho_mix As Double ' density of air-water mixture.
        Dim Psat_Tma As Double ' Saturated pressure of water at Tma
        Dim K_j_gap As Double
        Dim K_j_gap2 As Double
        Dim Tgap As Double 'Average temperature in the gap
        Dim L_avg As Double ' Average latent heat of vaporization
        
       ' Only the variables have to be updated.  Everything else is constant or gets recalculated.
        ' See figure 1 of
        Tih = x(1)
        Tma = x(2)
        

        
        'Saturated pressure of water at Tma ' HAVE TO USE THE SEAWATER RELATIONSHIP TO KEEP
        ' COMPARISON FAIR!
        Psat_Tma = mdlProperties.SeaWaterVaporPressure(Tma, 0)
        
        Tf = x(3)
        Tfw = x(4)
        Tic = x(5)
        delta_w = x(6)
        '(x(7) = Pv_ma / (X_ma+1) is a variable that
        ' combines two variables of which only one or the other is active.
        ' X_ma is the quality of the water x_ma = 1 indicates 100% water vapor
        ' while X_ma = 0 indicates 100% liquid water. .)
        ' This is a new condition 7/19/2017 that allows
        ' the vapor to partially immediately condense (in aerosol form?) at the air gap interface such
        ' that the full latent heat of vaporization is not transfered at the main sight of condensation at the
        ' cooling wall. Liquid water droplets are assumed to still all gather on the cooling wall. This is the
        ' transition to permeate gap membrane distillation. If X_ma reaches a value of 0, then Air gap membrane
        ' distillation is approaching permeate gap membrane distillation depending on the gap thickness.
        If x(7) > Psat_Tma / 2 Then
           X_ma = Psat_Tma / x(7) - 1
           Pv_ma = Psat_Tma
        Else
           X_ma = 1
           Pv_ma = 2 * x(7)
        End If
        
        If X_ma < 0 Or Pv_ma > Psat_Tma Or X_ma > 1 Then
            Pv_ma = mdlError.ReturnError("clsHTFuncAirGap.EvaluateFunction: The air gap water quality or saturated" & _
                   " pressure is incorrect!", , True)
        End If
                    
        AirOnlyThickness = Lz_a - delta_w
        'Calculate the Mass transfer - this is ONLY one estimate for this that must be balanced against the diffusion of the air gap!
        '   before it becomes a valid evaluation of the

        m_M = MembraneDistillationMassTransfer(Tih, S_ih, P_h, Tma, S_ic, P_c, A_f, Phi, r_mean, Lz_m, K_m, Use_K_m, HTModel, Pv_ma)
        
        If m_M <= 0 Then
           m_M = mdlError.ReturnError("clsHTFuncAirGap.EvaluateFunction: The membrane mass flow is negative! The problem" & _
                   " is probably ill posed or numeric settings need to be changed on the initial guess!", , True)
        End If
        
        Tgap = (Tma + Tf) / 2
        ' Diffusion coefficient of water vapor mixture. Per Alsaadi equation 16.  Need to know the range of validity for this equation!!!
        ' !@#$ - THIS MODEL FOR DIFFUSION COEFFICIENT NEEDS TO BE KEPT CONSISTENT WITH the MembraneDistillationMassTransfer function!!
        D_ab = 0.00001895 * (Tgap) ^ 2.072 / P_air_gap
        '
        Mv = mdlConstants.glbWaterMolecularWeight
        ' Mole fraction of water vapor
        yma = Pv_ma / P_air_gap '
        yf = mdlProperties.SeaWaterVaporPressure(Tf, 0) / P_air_gap ' This assumes 0 salinity-in the future there may need to be a model for salinity encroachment
        yavg = (yma + yf) / 2
        ' Molecular weight of air-water mixture
        Mav = yavg * Mv + (1 - yavg) * mdlConstants.glbAirMolecularWeight
        ' Average density of the air-water mixture via the ideal gas law
        rho_mix = P_air_gap * Mav / (mdlConstants.glbGasConstant * Tgap)
        ' Molar concentration
        cM = rho_mix / Mv
        
        L_avg = mdlProperties.LatentHeatOfPureWater((Tma + Tih) / 2)
        'Overall heat transfer Coefficient (we use the local heat transfer coefficients Hc, Hh, Hm, Ha, Hf, Hw)
        U = mdlTransferCoefficient.HTC(Tcb, Thb, M_avgc, M_avgh, L_avg, m_M, Tih, Tic, Hc, Hh, Hm, _
                                                     L_Norm_c, CSpacer, Lzc, L_TotalPara_c, L_Norm_h, HSpacer, Lzh, L_TotalPara_h, A_f, Phi, k_mf, Lz_m, S_ic, S_ih, S_avgc, S_avgh, Use_K_m, HTModel, _
                                                     False, Ha, Hf, Hw, Tma, Tf, Tfw, P_air_gap, ASpacer, delta_w, Lz_f, k_foil, X_ma)
        ' Condensate film thickness - this is equation 24 of Alsaadi where it is assumed that the film condensate thickness can be approximated as piecewise linear
        ' (a new linear slope for every element is assumed.
        '
        Tf_avg = (Tf + Tfw) / 2
        rho_liq = mdlProperties.DensitySaturatedLiquidPureWater(Tf_avg)
        rho_gas = mdlProperties.AirWaterSaturatedMixtureDensity(Tf_avg, P_air_gap)
        mu = mdlProperties.SeaWaterViscosity(Tf_avg, 0)
        B0 = A_f * rho_liq * (rho_liq - rho_gas) * mdlConstants.glbGravity / mu
        '
        A_slope = mdlMath.SmallestPositiveRealCubicRoot(L_Para_g ^ 2, 2 * L_Para_g * delta_w_in, delta_w_in ^ 2, -(m_M / B0), glbZeroThresholdForCubicSolution)
        ' Make sure a good root has been found
        If A_slope = mdlConstants.glbINVALID_VALUE Then
           A_slope = mdlError.ReturnError("clsHTFuncAirGap.EvaluateFunction: The function mdlMath.SmallestPositiveRealCubicRoot did not find a real positive root to" _
                                          & " the cubic polynomial used to solve for the consensate film thickness. It is therefore necessary to troubleshoot why.", , True)
        End If
        delta_w_out = A_slope * L_Para_g + delta_w_in
        
        ' Overall heat transfer rate
        Q = U * A_f * (Thb - Tcb)
        
        ' This is comparable to the K_j for membrane mass transfer
        K_j_gap = (cM * D_ab * Mv / AirOnlyThickness) / P_air_gap
        
        ' Let's try Hitsov,2017 formulation. THIS FORMULATION IS INCORRECT!
        D_ab2 = 0.00000446 * Tgap ^ 2.334
        K_j_gap2 = (D_ab2 / (AirOnlyThickness * mdlConstants.glbGasConstant / Mav * Tgap)) * Log((P_air_gap - yf * P_air_gap) / (P_air_gap - Pv_ma)) * 1 / (Pv_ma - yf * P_air_gap)
        
        ReDim Result(1 To 7)
        ' Calculate the total heat transfer several ways these are the equations that need to be driven to zero if the interface temperatures are consistent
        
        ' No need to multiply by A_f since balancing W/m2 will balance W!
        ' 1. Heat transfer from hot feed flow to membrane heat flow comparison
        Result(1) = Hh * (Thb - Tih) - Hm * (Tih - Tma)
        ' 2. Heat transfer from membrane to air gap comparison
        Result(2) = Hm * (Tih - Tma) - Ha * (Tma - Tf)
        ' 3. Heat transfer from air gap to condensing film comparison
        Result(3) = Ha * (Tma - Tf) - Hf * (Tf - Tfw)
        ' 4. Heat transfer from condensing film to cooling wall interface comparison
        Result(4) = Hf * (Tf - Tfw) - Hw * (Tfw - Tic)
        ' 5. Heat transfer from cooling wall interface to bulk stream cooling temperature
        Result(5) = Hw * (Tfw - Tic) - Hc * (Tic - Tcb)
        ' 6. Mass transfer balance between resistance to mass flow through membrane and diffusive resistance across the air gap.
        Result(6) = m_M - A_f * K_j_gap2 * (yma - yf) * P_air_gap
        ' 7. condsate film thickness - the first term represents the variable as it is being guided
        '    by the mathematical solution procedure.  The second term is the estimate of what the thickness should be based on Nusselt film condensation
        '    on a vertical plate (equation 24 of Alsaadi et. al, 2013.)
        Result(7) = delta_w - (delta_w_in + delta_w_out) / 2
    End If
End Sub

Public Sub ConvergenceCriteria(ConvergCrit() As Double, CVArea As Double)

ReDim ConvergCrit(1 To 8)
Dim EnergyCritOverArea As Double

EnergyCritOverArea = mdlConstants.glbAbsConvergCriterionEnergy / CVArea

ConvergCrit(1) = EnergyCritOverArea
ConvergCrit(2) = EnergyCritOverArea
ConvergCrit(3) = EnergyCritOverArea
ConvergCrit(4) = EnergyCritOverArea
ConvergCrit(5) = EnergyCritOverArea
ConvergCrit(6) = mdlConstants.glbAbsConvergCriterionMassFlow
ConvergCrit(7) = mdlConstants.glbAbsConvergCriterionThickness

End Sub

Private Sub Class_Initialize()
   PrNeedToReadConstants = True
   IsFirstTimeThrough = True
End Sub

