Attribute VB_Name = "mdlTransferCoefficient"
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


'Transfer coefficient module (K - mass transfer coefficient H - heat transfer coefficient)

Option Explicit

Private Sub EvaluateProperties(ByVal Temperature As Double, _
                               Optional ByVal Salinity As Double, _
                               Optional ByRef Density As Double, _
                               Optional ByRef Viscosity As Double, _
                               Optional ByRef ThermalConductivity As Double, _
                               Optional ByRef SpecificHeat As Double, _
                               Optional ByRef Prandtl As Double)
                               ' THIS PROCEDURE IS NO LONGER USED BECAUSE IT IS INEFFICIENT!.
    Dim TempVar As Variant
    
    If IsMissing(Salinity) Then
        If Not IsMissing(Density) Then
           Density = mdlCharlieProp.WaterDen(Temperature)
        End If
        If Not IsMissing(Viscosity) Then
           Viscosity = mdlCharlieProp.WaterViscosity(Temperature)
        End If
        If Not IsMissing(ThermalConductivity) Then
           ThermalConductivity = mdlCharlieProp.WaterThermCond(Temperature)
        End If
        If Not IsMissing(Prandtl) Then
           Prandtl = mdlCharlieProp.WaterPrandtl(Temperature)
        End If
        If Not IsMissing(SpecificHeat) Then
           SpecificHeat = mdlCharlieProp.WaterSpHeat(Temperature)
        End If
    Else
        If Not IsMissing(Density) Then
           Density = mdlProperties.SeaWaterDensity(Temperature, Salinity, "Sun_et_al")
        End If
        If Not IsMissing(Viscosity) Or Not IsMissing(Prandtl) Then
           Viscosity = mdlProperties.SeaWaterViscosity(Temperature, Salinity, "Isdale_et_al")
        End If
        If Not IsMissing(ThermalConductivity) Or Not IsMissing(Prandtl) Then
           ThermalConductivity = mdlProperties.SeaWaterThermalConductivity(Temperature, Salinity, "JamiesonAndTudhope")
        End If
        If Not IsMissing(SpecificHeat) Or Not IsMissing(Prandtl) Then
           SpecificHeat = mdlProperties.SeaWaterSpecificHeat(Temperature, Salinity, "Jamieson_et_al")
        End If
        
        ' This is a very quick calculation so if/then is not worth it
        Prandtl = SpecificHeat * Viscosity / ThermalConductivity
    End If
    
End Sub
                               
Public Function Reynolds(MassFlow As Double, Density As Double, Viscosity As Double, height As Double, width As Double, _
                          Optional ByRef HydraulicDiameter As Double, _
                          Optional ByRef AverageVelocity As Double, Optional UseSpacerPorosityMethodHydraulicDiameter As Boolean = False, _
                          Optional Spacer As clsSpacer, _
                          Optional SpacerThickness As Double)
                          
' DO NOT USE Spacer.Thickness - This is the original, unaltered value read from the INPUT. The thickness is varied
' according to a multiplier based on whether symmetry conditions are being applied for multi-layer runs!
    Dim Area As Double
    Dim rho As Double
    Dim Re As Double

        Area = width * height
        
        If UseSpacerPorosityMethodHydraulicDiameter Then
            If IsMissing(Spacer) Then
              Re = mdlError.ReturnError("mdlTransferCoefficient.Reynolds: A spacer hydraulic diameter was requested, but no spacer was included" & _
                                        " in the input arguments!")
            Else
              HydraulicDiameter = 4 * Spacer.Porosity * Spacer.FilamentDiameter * (SpacerThickness) / _
                                 (2 * Spacer.FilamentDiameter + 4 * (1 - Spacer.Porosity) * (SpacerThickness))  'DO NOT USE Spacer.Thickness!!!!
              
            End If
        Else
           HydraulicDiameter = 2 * Area / (width + height)
        End If
        AverageVelocity = MassFlow / (Density * Area)
        Re = Density * AverageVelocity * HydraulicDiameter / Viscosity
        
        If Re > mdlConstants.glbMaxReynolds Then
            Reynolds = mdlError.ReturnError("mdlTC.Reynolds: The reynolds number (" & Re & ") is above " & _
                         mdlConstants.glbMaxReynolds & " which indicates" & _
                         " the beginning of the transition to turbulent.  This code is only valid for " & _
                         "laminar flows!")
        Else
            Reynolds = Re
        End If
        
End Function


Public Function K_MassTransferCoef(Temperature As Double, _
                                  MassFlow As Double, _
                                  Salinity As Double, _
                                  width As Double, _
                                  height As Double, _
                                  Length As Double, _
                                  Spacer As clsSpacer, _
                                  Optional DescStr As String = "") As Double
'm/s
Dim Density As Double
Dim Renold As Double
Dim Viscosity As Double
Dim ThermalConductivity As Double
Dim SpecificHeat As Double
Dim HydraulicDiameter As Double 'hydraulic diameter
Dim Schmidt As Double
Dim Sherwood As Double
Dim SoluteDiffusivity As Double
Dim DebugString As String

'EvaluateProperties Temperature, Salinity, Density, Viscosity, ThermalConductivity
Density = mdlProperties.SeaWaterDensity(Temperature, Salinity)
Viscosity = mdlProperties.SeaWaterViscosity(Temperature, Salinity)
ThermalConductivity = mdlProperties.SeaWaterThermalConductivity(Temperature, Salinity)

'
Renold = Reynolds(MassFlow, Density, Viscosity, height, width, HydraulicDiameter, , True, Spacer, height)
'
If Not mdlError.NoError Then
   GoTo ErrorHappened
End If

SoluteDiffusivity = mdlProperties.SoluteDiffusivityNaClSolution(Temperature, Salinity, "ChiamAndSarbatly")
'
Schmidt = Viscosity / (Density * SoluteDiffusivity)

Sherwood = 1.86 * (Renold * Schmidt * HydraulicDiameter / Length) ^ 0.33

If mdlError.NoError Then
     K_MassTransferCoef = Sherwood * SoluteDiffusivity / HydraulicDiameter
Else
     GoTo ErrorHappened
End If

If glbEvaluateEquationsAtEnd Then
   If IsFinalRun Then
   
     DebugString = DebugString & FormatDebugColumns("Sc    : " & DescStr & " Schmidt number:   " & Schmidt & vbCrLf)
     DebugString = DebugString & FormatDebugColumns("Sh    : " & DescStr & " Sherwood number:  " & Sherwood & vbCrLf)
     DebugString = DebugString & FormatDebugColumns("K : " & DescStr & " Salt concentration mass transfer coeff. (m/s): " & K_MassTransferCoef & vbCrLf)
     Debug.Print DebugString
     mdlValidation.WriteDebugInfoToFile DebugString
   End If
End If


EndOfFunction:

Exit Function
ErrorHappened:
   K_MassTransferCoef = mdlError.ReturnError(IncludeMsgBox:=True)
GoTo EndOfFunction
End Function


Sub CrossFlowHeatExchanger(A_f, U, e, Cpih, Cpic, Mih, Mic, Cmin)

Dim Ch As Double
Dim Cc As Double
Dim Cmax As Double
Dim Cstar As Double
Dim NTU As Double
Dim gamma As Double


Ch = Cpih * Mih
Cc = Cpic * Mic

Cmin = WorksheetFunction.Min(Ch, Cc)
Cmax = WorksheetFunction.Max(Ch, Cc)

Cstar = Cmin / Cmax

NTU = U * A_f / Cmin

gamma = 1 - Exp(-NTU)
e = (1 - Exp(-gamma * Cstar)) / Cstar

End Sub

Function HTC(T_c As Double, _
                           T_h As Double, _
                           M_c As Double, _
                           M_h As Double, _
                           Lavg As Double, _
                           m_MD As Double, _
                           T_ih As Double, _
                           T_ic As Double, _
                           Hci As Double, _
                           Hhi As Double, _
                           Hmi As Double, _
                           LengthNormalToColdFlow As Double, _
                           ColdSpacer As clsSpacer, ColdSpacerThickness As Double, _
                           LengthParallelToColdFlow As Double, _
                           LengthNormalToHotFlow As Double, _
                           HotSpacer As clsSpacer, HotSpacerThickness As Double, _
                           LengthParallelToHotFlow As Double, _
                           Area As Double, _
                           MembranePorosity As Double, MembraneThermalConductivity As Double, MembraneThickness As Double, _
                           Salinity_c As Double, Salinity_h As Double, BulkSalinity_c As Double, BulkSalinity_h As Double, _
                           UseDirectMassTransferCoefficient, HTModel As Long, Optional IsParallelMembraneThermalConductivityModel As Boolean = False, _
                           Optional Hai As Double = 1E+20, Optional Hfi As Double = 1E+20, Optional Hwi As Double = 1E+20, _
                           Optional Ta As Double, Optional Tf As Double, Optional Tfw As Double, Optional Pressure As Double, _
                           Optional AirGapSpacer As clsSpacer, Optional CondensateThickness As Double, Optional WallThickness As Double, _
                           Optional FoilThermalConductivity As Double, Optional AirGapStartQuality) As Double 'U W/(m2*K)
                           

' Inputs T_c - cold side temperature of bulk flow (K)
'        T_h - hot side temperature of bulk flow (K)
'        M_c - cold side mass flow rate (kg/s)
'        M_h - hot side mass flow rate (kg/s)
'        Lavg - Latent heat of vaporization for average of hot and cold sides (J/kg)
'        m_MD - membrane distillation mass flow rate that is vaporizing and flowing across the membrane (kg/s)
'        T_ih - interface temperature at membrane wall on the hot side (K)
'        T_ic - interface temperature at membrane wall on the cold side (K)
'        Hci  - (for output) - cold side convective heat transfer coefficient between the bulk flow and membrane ( or for air gap cooling wall foil) interface (W/(m2*K))
'        Hhi  - (for output) - hot side ""
'        Hmi  - (for output) - through membrane heat transfer coefficient that includes latent heat of vaporization that occurs as m_MD flows
'                              across the membrane
'    ' For air gap only - (optional)
'
'        Hai  - (for output) - air gap heat transfer coefficient that includes enthalpy of water vapor due to mass transport
'        Hfi  - (for output) - condensed fluid heat transfer coefficient
'        Hwi  - (for output) - cooling wall heat transfer coefficient
'        Ta   - temperature at the membrane-air gap interface
'        Tf   - temperature at the air gap-condensed fluid interface
'        Tfw  - temperature at the condensed fluid-cooling wall foil interface
'
'        Hci, Hhi, Hmi are 0 and are passed to recieve a value
    
    Dim h_c As Double
    Dim h_h As Double
    Dim keff As Double
    Dim T_avg As Double
    Dim k_air As Double
    Dim rft As Double
    Dim Leh As Double
    Dim Lec As Double
    Dim beta As Double
    Dim Dh_h As Double ' Hot side hydraulic diameter
    Dim Dh_c As Double ' Cold side hydraulic diameter
    Dim Re_h As Double
    Dim Re_c As Double ' hot and cold side reynolds numbers
    Dim Nu_h As Double
    Dim Nu_c As Double
    Dim Pr_h As Double
    Dim Pr_c As Double
    Dim DebugString As String
    
    'DO NOT USE ""Spacer.Thickness - it is the original input value but the thickness is varied by a multiplier to account for
    '                                multiple layers
    
    Hci = H_RectPassage(T_c, M_c, LengthNormalToColdFlow, ColdSpacerThickness, LengthParallelToColdFlow, BulkSalinity_c, ColdSpacer, _
                            T_ic, Salinity_c, glbUseTurbulentNusseltForForcedConvection, Dh_c, Re_c, Nu_c, Pr_c)
    Hhi = H_RectPassage(T_h, M_h, LengthNormalToHotFlow, HotSpacerThickness, LengthParallelToHotFlow, BulkSalinity_h, HotSpacer, _
                            T_ih, Salinity_h, glbUseTurbulentNusseltForForcedConvection, Dh_h, Re_h, Nu_h, Pr_h)
    
    'Daniel Villa 9-13-2016.  Previously there was a serious error in the heat
    ' transfer.  The latent heat of evaporation was included in the interface temperatures
    ' but not in the overall heat transfer coefficient.  It plays an important role in the
    ' resistance in the membrane but not in the cold and hot flows.
    'OLD
    'Hmi = keff / Lz_m
    'NEW
    If HTModel = 0 Then
       k_air = DryAirThermalConductivityAt1Atm((T_ih + T_ic) / 2)
    ElseIf HTModel = 1 Then 'There is a chance that condensation begins Before exiting the membrane
       k_air = DryAirThermalConductivityAt1Atm((T_ih + Ta) / 2)
    End If
        
    If UseDirectMassTransferCoefficient And MembranePorosity = mdlConstants.glbINVALID_VALUE Then
        keff = MembraneThermalConductivity
    Else
        If IsParallelMembraneThermalConductivityModel Then
           'Hitsov et. al. argue that this simple parrallel mixture law tends to largly overestimate the
           ' effective membrane thermal conductivity.
                    'Page 52: "García-Payo and Izquierdo-Gil [33] performed an extensive
                    'evaluation of 9 different models for prediction of thermal conductivity
                    'of the membrane matrices of 2 PVDF, 2 PTFE and 2 supported
                    'PTFE membranes and compared them to experimental data. The
                    'authors concluded that the commonly used parallel model largely
                    'overestimates the thermal conductivity, whereas the series model
                    'slightly underestimates it."
           keff = MembranePorosity * k_air + (1 - MembranePorosity) * MembraneThermalConductivity ' should be kair
        Else
           ' This is the model recommended by Hitsov et. al.
           'Maxwell type I equation presented by 2015 Hitsov et. al. "Modeling approaches in MD - a critical review" Separation and Purification Technology
           beta = (MembraneThermalConductivity - k_air) / (MembraneThermalConductivity + 2 * k_air)
           keff = k_air * (1 + 2 * beta * (1 - MembranePorosity)) / (1 - beta * (1 - MembranePorosity))
        End If
    End If
    
    If HTModel = 0 Then
       
       Hmi = keff / MembraneThickness + (Lavg * m_MD / (Area * (T_ih - T_ic)))
    ElseIf HTModel = 1 Then 'There is a chance that condensation begins Before exiting the membrane
       Hmi = keff / MembraneThickness + (Lavg * (AirGapStartQuality) * m_MD / (Area * (T_ih - Ta)))
    End If
    
    ' Additional calculations are needed for Air gap membrane distillation
    If HTModel = 1 Then
    
        Dim kAB As Double ' thermal conductivity of saturated air-water vapor mixture
        Dim kS As Double 'thermal conductivity of the spacer
        Dim Enthalpy_a As Double ' enthalpy of the vapor at the membrane/air gap interface
        Dim Enthalpy_a_ma As Double
        Dim Enthalpy_f_ma As Double
        Dim Enthalpy_f As Double ' enthalpy of the vapor at the air gap/condensed fluid interface
        Dim kf As Double ' thermal conductivity of the falling film (we neglect mass transport equations to avoid having to move this analysis to a global set of equations)
        Dim eS As Double ' air gap spacer porosity
        Dim AirGapThickness As Double
        
        kAB = mdlProperties.AirWaterSaturatedMixtureThermalConductivity((Ta + Tf) / 2, Pressure)
        kS = AirGapSpacer.ThermalConductivity
        eS = AirGapSpacer.Porosity
        AirGapThickness = AirGapSpacer.Thickness - CondensateThickness
        
        Enthalpy_a_ma = mdlProperties.SaturatedVaporEnthalpyPureWater(Ta)
        Enthalpy_f_ma = mdlProperties.SaturatedLiquidEnthalpyPureWater(Ta)
        ' Added 7/19/2017. If entering vapor immediately condenses, then we will have a non-zero quality.
        ' Quality is the fractio nof water that is in vapor form.
        Enthalpy_a = Enthalpy_a_ma * AirGapStartQuality + Enthalpy_f_ma * (1 - AirGapStartQuality)
        
        'Added 7/11/2017
        'Enthalpy_f = mdlProperties.SaturatedVaporEnthalpyPureWater(Tf)
        Enthalpy_f = mdlProperties.SaturatedLiquidEnthalpyPureWater(Tf)
        
        kf = mdlProperties.SubCooledWaterThermalConductivity(Pressure, (Tf + Tfw) / 2)
        
        Hai = (kAB * eS + (1 - eS) * kS) / AirGapThickness + m_MD * (Enthalpy_a - Enthalpy_f) / (Area * (Ta - Tf))
        Hfi = kf / CondensateThickness
        Hwi = FoilThermalConductivity / WallThickness

    End If
    
    ' Hai, Hfi, and Hwi are all related to air-gap membrane distillation only.  Otherwise they are just 1e20 to respresent 0 resistance
    HTC = 1 / (1 / Hhi + 1 / Hmi + 1 / Hci + 1 / Hai + 1 / Hfi + 1 / Hwi)
    
    If glbEvaluateEquationsAtEnd Then
        If IsFinalRun And IsFinalNewtonIteration Then
            DebugString = DebugString & FormatDebugColumns("Ka  : Air Thermal Conductivity at average interface temp (W/m/K):    " & k_air & vbCrLf)
            DebugString = DebugString & FormatDebugColumns("Km    : Effective thermal conductivity of the membrane (W/m/K):   " & keff & vbCrLf)
            DebugString = DebugString & FormatDebugColumns("Dh_h   : Hot side hydraulic diameter (m):  " & Dh_h & vbCrLf)
            DebugString = DebugString & FormatDebugColumns("Dh_c   : Cold side hydraulic diameter (m): " & Dh_c & vbCrLf)
            DebugString = DebugString & FormatDebugColumns("Re_hot : Hot side reynolds number: " & Re_h & vbCrLf)
            DebugString = DebugString & FormatDebugColumns("Re_cold    : Cold side reynolds number:    " & Re_c & vbCrLf)
            DebugString = DebugString & FormatDebugColumns("Nu_hot : Hot side Nusselt number:  " & Nu_h & vbCrLf)
            DebugString = DebugString & FormatDebugColumns("Nu_cold    : Cold side Nusselt number: " & Nu_c & vbCrLf)
            DebugString = DebugString & FormatDebugColumns("Pr_hot : Hot side Prandtl number:  " & Pr_h & vbCrLf)
            DebugString = DebugString & FormatDebugColumns("Pr_cold    : Cold side Prandtl number: " & Pr_c & vbCrLf)
            DebugString = DebugString & FormatDebugColumns("h_hot  : Hot side heat transfer coefficient (W/m2/K):    " & Hhi & vbCrLf)
            DebugString = DebugString & FormatDebugColumns("h_cold : Cold side heat transfer coefficient (W/m2/K): " & Hci & vbCrLf)
            Debug.Print DebugString
            mdlValidation.WriteDebugInfoToFile DebugString
        End If
    End If
    
    
    
End Function

Function HTC_Natural_Convection(T_Air As Double, _
                           T_water As Double, _
                           M_water As Double, _
                           T_ia As Double, _
                           T_iw As Double, _
                           Hwi As Double, _
                           Hii As Double, _
                           Hai As Double, _
                           LengthNormalToWaterFlow As Double, _
                           WaterSpacer As clsSpacer, _
                           WaterThickness As Double, _
                           LengthParallelToWaterFlow As Double, _
                           Area As Double, _
                           GravityDirectionLength As Double, _
                           ExternalInsulationMaterial As clsMaterial, _
                           Salinity_Water As Double, _
                           Interface_Salinity_Water As Double) As Double 'U W/(m2*K)
                           
' This is an overall heat transfer coefficient calculation for natural convection (i.e. unforced convection)
' Inputs T_air - Ambient external air temperture (K)
'        T_water - water bulk flow temperature (K)
'        M_water - water mass flow rate (i.e. water side is forced) (kg/s)
'        T_ia - Air/External insulation interface temperature (K)
'        T_iw - Water/External Insulation interface temperature (K)
'        Hwi  - Water to external insulation interface heat transfer coefficient (W/K) (output)
'        Hii  - External insulation heat transfer coefficient (W/K) (output)
'        Hai  - Air to external insulation heat transfer coefficient (W/K) (output)

    ' Do not use WaterSpacer.Thickness!

    Hwi = H_RectPassage(T_water, M_water, LengthNormalToWaterFlow, WaterThickness, LengthParallelToWaterFlow, Salinity_Water, WaterSpacer, T_iw, Interface_Salinity_Water)
    
    Hai = H_Natural_Convection(T_Air, T_ia, GravityDirectionLength) 'YOU HAVE TO USE THE TOTAL LENGTH OF THE ASSEMBLY, OTHERWISE YOU ARE GOING TO MAKE
     'THE SIMULATION UNSTABLE BECAUSE THE RAYLEIGH NUMBER
    
    Hii = ExternalInsulationMaterial.ThermalConductivity / ExternalInsulationMaterial.Thickness
    
    HTC_Natural_Convection = 1 / (1 / Hwi + 1 / Hai + 1 / Hii)
End Function

Function H_RectPassage(Temperature As Double, _
                                  MassFlow As Double, _
                                  width As Double, _
                                  height As Double, _
                                  Length As Double, _
                                  Salinity As Double, _
                                  Spacer As clsSpacer, _
                                  InterfaceTemperature As Double, _
                                  InterfaceSalinity As Double, _
                                  Optional UseTurbulentNusselt As Boolean = True, _
                                  Optional HydraulicDiameter As Double, _
                                  Optional Reyn As Double, _
                                  Optional Nuss As Double, _
                                  Optional Pran As Double) As Double
    'Temperature in Kelvin
    'MassFlow in kg/s
    'L - height of passage
    'Lz - width of passage
    
    Dim rho As Double
    Dim Re As Double
    Dim mu As Double
    Dim k As Double
    Dim Pr As Double
    Dim Nu As Double
    Dim Cp As Double
    Dim Dh As Double 'hydraulic diameter
    Dim rho_i As Double
    Dim mu_i As Double
    Dim k_i As Double
    Dim Cp_i As Double
    Dim Pr_i As Double
    
    rho = SeaWaterDensity(Temperature, Salinity) 'From mdlProperties
    mu = SeaWaterViscosity(Temperature, Salinity)
    k = SeaWaterThermalConductivity(Temperature, Salinity)
    Cp = SeaWaterSpecificHeat(Temperature, Salinity)
    
'    rho_i = SeaWaterDensity(InterfaceTemperature, InterfaceSalinity) 'From mdlProperties
'    mu_i = SeaWaterViscosity(InterfaceTemperature, InterfaceSalinity)
'    k_i = SeaWaterThermalConductivity(InterfaceTemperature, InterfaceSalinity)
'    Cp_i = SeaWaterSpecificHeat(InterfaceTemperature, InterfaceSalinity)
    
    Pr = Cp * mu / k
    Pran = Pr
    'Pr_i = Cp_i * mu_i / k_i
    
    Re = Reynolds(MassFlow, rho, mu, height, width, Dh, , True, Spacer, height)
    Reyn = Re
    HydraulicDiameter = Dh
    If UseTurbulentNusselt Then
         Nu = FlatSheetTurbulentFlow(Re, Pr, "Alsaadi_et_al", height, Spacer)
    Else
         Nu = FlatSheetLaminarFlow(Re, Pr, Dh, Length, "Gryta_et_al_Table1No11") '"Gryta_et_al_Table1No6")
    End If
    Nuss = Nu
    
    If mdlError.NoError Then
        H_RectPassage = Nu * k / Dh
    Else
        MsgBox err.Source
        Debug.Assert False
    End If

End Function

Function RayleighOfAir(TemperatureAir As Double, TemperatureSurface As Double, Length As Double) As Double

    Dim alpha As Double  'thermal diffusivity
    Dim Nu As Double     'kinematic viscosity
    Dim beta As Double   'coefficient of thermal expansion
    Dim Tfilm As Double
    
    'Approximate the film temperature
    Tfilm = (TemperatureAir + TemperatureSurface) / 2
    
    alpha = mdlProperties.DryAirThermalDiffusivityAt1Atm(Tfilm)
    beta = 1 / Tfilm ' Ideal gas assumption
    Nu = mdlProperties.DryAirKinematicViscosityAt1Atm(Tfilm)
    
    'Natural convection can work either way (thus the absolute function)
    RayleighOfAir = mdlConstants.glbGravity * beta * mdlMath.AbsDbl(TemperatureSurface - TemperatureAir) * Length ^ 3 / (alpha * Nu)

End Function

Function H_Natural_Convection(TemperatureAir As Double, TemperatureSurface As Double, Length As Double) As Double

    Dim Ra As Double
    Dim Tfilm As Double
    Dim Nu As Double
    Dim k_air As Double
    Dim Pr As Double
    
    'Approximate the film temperature
    Tfilm = (TemperatureAir + TemperatureSurface) / 2
    
    Ra = RayleighOfAir(TemperatureAir, TemperatureSurface, Length)
    Pr = mdlProperties.DryAirPrandtlAt1Atm(Tfilm)
    k_air = DryAirThermalConductivityAt1Atm(Tfilm) ' mdlCharlieProp.AirThermalCond(Tfilm)
    
    Nu = mdlNusselt.NusseltNumberNaturalConvection(Ra, Pr)
    
    If mdlError.NoError Then
        H_Natural_Convection = Nu * k_air / Length
    Else
        H_Natural_Convection = mdlConstants.glbINVALID_VALUE
    End If

End Function

Function H_Condensation_VerticalConstantTemperatureSurface(TsatPvk As Double, TIcond As Double, Pvk As Double, _
                                                           mcond As Double, Lcond As Double, _
                                                           wframe As Double, delfilmi As Double)

'TsatPvk - saturated water temperature at pressure Pvk (Kelvin)
'TIcond - interface temperature of condensing surface (Kelvin)
'Pvk    - vacuum pressure of module k (Pa)
'mcond  - condensation mass flow rate
' ...
' Definitions same as those in mdlMath.ModuleProperties

' This is from Lindon C. Thomas "Heat Transfer" 2nd Edition page 618 - 622
Dim Tavg As Double
Dim Ref As Double 'condensate reynolds number
Dim Prf As Double 'condensate prandtl number
Dim rhof As Double
Dim rhofg As Double
Dim kf As Double 'thermal conductivity
Dim muf As Double 'viscosity
Dim pcond As Double 'average perimeter of condensate film
Dim Ans As Double
Dim hfvH2O As Double 'latent heat of vaporization
Dim Jaf As Double ' Ja number (cannot remember the name) dimensionless ratio of sensible heat to latent heat for a unit mass
Dim cpf As Double ' specific heat of water at Tavg

Tavg = (TsatPvk + TIcond) / 2
    
rhof = mdlProperties.SubCooledWaterDensity(Pvk, Tavg)
rhofg = mdlProperties.DensityChangeOnCondensationPureWater(Tavg)
kf = mdlProperties.SubCooledWaterThermalConductivity(Pvk, Tavg)
muf = mdlProperties.SubCooledWaterViscosity(Pvk, Tavg)
cpf = mdlProperties.SubCooledWaterSpecificHeat(Pvk, Tavg)



' Even though the second term is complicated it is very small
pcond = 2 * wframe + 2 * delfilmi

'Condensate reynolds number
Ref = 4 * mcond / (pcond * muf)

If mdlError.NoError Then
    If Ref < 0 Then
        Ans = mdlError.ReturnError("mdlTC.H_Condensation_VerticalConstantTemperatureSurface:" & _
                      " the condensate Reynold's number is negative which is invalid!")
    ElseIf Re < 30 Then
        hfvH2O = mdlProperties.LatentHeatOfPureWater(TsatPvk)
        Jaf = cpf * (TsatPvk - TIcond) / hfvH2O
        
        Ans = 0.943 * ((rhof * mdlConstants.glbGravity * rhofg * kf ^ 3 * hfvH2O * (1 + 0.68 * Jaf)) _
                       / (muf * (TsatPvk - TIcond) * Lcond)) ^ 0.25
    ElseIf Re < 1800 Then
        Ans = (Ref / (1.08 * Ref ^ 1.22 - 5.2)) * ((rhof * mdlConstants.glbGravity * rhofg * kf ^ 3) / muf ^ 2) ^ (1 / 3)
    ElseIf Re < 100000 Then
        Ans = (Ref / (8750# + 58# * Prf ^ -0.5 * (Ref ^ 0.75 - 253#))) * (rhof * mdlConstants.glbGravity * (rhofg) * kf ^ 3 / muf ^ 2) ^ (1 / 3)
    Else
        Ans = mdlError.ReturnError("mdlTC.H_Condensation_VerticalConstantTemperatureSurface:" & _
                      " the condensate Reynold's number (" & Ref & ") is greater than 1e5 which is beyond the valid range of the equations used!")
    End If
Else
    Ans = mdlConstants.glbINVALID_VALUE
End If

H_Condensation_VerticalConstantTemperatureSurface = Ans

End Function


Function MembraneDistillationMassTransfer(HotTemperature As Double, _
                                          HotSalinity As Double, _
                                          HotPressure As Double, _
                                          ColdTemperature As Double, _
                                          ColdSalinity As Double, _
                                          ColdPressure As Double, _
                                          Area As Double, _
                                          Porosity As Double, _
                                          MeanPoreRadius As Double, _
                                          Thickness As Double, _
                                          DirectMassTransferCoefficient As Double, _
                                          UseDirectInputForMassTransferCoefficient As Boolean, _
                                          Optional HTModel As Long = 0, _
                                          Optional Pv_air As Double = 0, _
                                          Optional K_j As Double = 0) ' Default mass transfer model is direct contact = 0.  1 is air gap mass transfer
    Dim dDel As Double              'm
    Dim MwMix As Double             'kg/kmol
    Dim Xavg As Double
    Dim Xh As Double
    Dim Xc As Double
    Dim Pvh As Double               'bara
    Dim Pvc As Double
    Dim rho As Double               'kg/m³
    Dim A As Double
    Dim frac As Double
    Dim tou As Double
    Dim lamda As Double
    Dim Knudsen As Double
    Dim PD As Double
    Dim p_a As Double
    Dim Pavg As Double
    Dim Tavg As Double
    Dim DebugString As String
    
    If HotTemperature <= ColdTemperature Then
        MembraneDistillationMassTransfer = mdlError.ReturnError("mdlTransferCoefficient.MembraneDistillationMassTransfer:" & _
                            " The Hot Temperature is less than or equal to the Cold Temperature!" & _
                            " this is not the intended mode of operation for the model.", , True, True)
                            
    End If
    
    Pvh = mdlProperties.SeaWaterVaporPressure(HotTemperature, HotSalinity) 'Pa
    If HTModel = 0 Then
       Pvc = mdlProperties.SeaWaterVaporPressure(ColdTemperature, ColdSalinity) 'Pa
    End If
    ' For now, assume there is no difference between air gap and direct contact membrane distillation
    
    If Not UseDirectInputForMassTransferCoefficient Then 'If are assuming both convection and diffusion

       'References:
       ' Alkhudhiri, A., N. Darwish, N. Hilal, 2012. "Membrane distillation: A comprehensive review."
       '            Desalination 287: 2-18.
       ' Khayet, M., 2011. "Membranes and theoretical modeling of membrane distillation: A review."
       '             Advances in Colloid and Interface Science 164: 56-88
       Pavg = (HotPressure + ColdPressure) / 2
       Tavg = (HotTemperature + ColdTemperature) / 2
       
       ' Tortuosity per Alkhudhiri
       tou = (2 - Porosity) ^ 2 / Porosity
       
       ' mean free path
       lamda = mdlConstants.glbBoltzmann * Tavg / (Sqr(2) * mdlConstants.glbPi * Pavg * mdlConstants.glbWaterVaporCollisionDiameter ^ 2)
       Knudsen = lamda / (2 * MeanPoreRadius)
              
       If Knudsen > 1 Then
        'Membrane mass flux coefficient per Khayet
           K_j = (2 * Porosity * MeanPoreRadius / (3 * tou * Thickness)) * Sqr(8 * mdlConstants.glbWaterMolecularWeight / (mdlConstants.glbPi * mdlConstants.glbGasConstant * Tavg))
       ElseIf Knudsen > 0.01 And Knudsen < 1# Then
           ' Pressure time ordinary diffusivity coefficient
           ' from Alkhudhiri
           PD = 0.00001895 * Tavg ^ 2.072
           ' Partial pressure of air ( in vacuum this has to be nearly zero )
           p_a = Pavg - mdlProperties.SeaWaterVaporPressure(Tavg, 0#)
           If p_a < 0 Then
              p_a = 0
           End If
        'Membrane mass flux coefficient per Khayet
           K_j = (mdlConstants.glbWaterMolecularWeight / (mdlConstants.glbGasConstant * Tavg * Thickness)) * ((3 * tou / (2 * Porosity * MeanPoreRadius)) * _
                  Sqr(mdlConstants.glbPi * mdlConstants.glbWaterMolecularWeight / (8 * mdlConstants.glbGasConstant * Tavg)) + p_a * tou / (Porosity * PD)) ^ (-1)
       ElseIf Knudsen < 0.01 And Knudsen > 0 Then
          ' Added 8/16/2017 for completeness.  I do not expect this region to
          ' ever be exercised.  If this is dominating, then we have a problem. THIS LINE IS NOT VALIDATED AS OF 8/16/2017.
           K_j = mdlConstants.glbPi * PD * MeanPoreRadius ^ 2 _
                / (mdlConstants.glbGasConstant * Tavg * p_a * tou * Thickness)
       Else
           MembraneDistillationMassTransfer = mdlError.ReturnError("mdlTC.MembraneDistillationMassTransfer: The Knudsen number (" _
                                         & Knudsen & ") is outside of its valid range of being greater than 0")
           GoTo ErrorHappened
       End If
    Else
        K_j = DirectMassTransferCoefficient
    End If

    If HTModel = 0 Then ' Direct contact, use saturated vapor pressures.
       If Pvh < Pvc Then
       ' THIS IS OK FOR EXTREME CASES
'          MembraneDistillationMassTransfer = mdlError.ReturnError("mdlTransferCoefficient.MembraneDistillationMassTransfer:" & _
'                 " The hot vapor pressure is less than the cold vapor pressure which implies mass flow in the opposite direction intended!", _
'                 , True, True)
                 
       End If
       MembraneDistillationMassTransfer = Area * K_j * (Pvh - Pvc)
    ElseIf HTModel = 1 Then ' Air gap membrane mass transfer.
       If Pvh > Pv_air Then
          MembraneDistillationMassTransfer = Area * K_j * (Pvh - Pv_air)
       Else
         ' If this is stagnating, then indicate no mass change
            MembraneDistillationMassTransfer = mdlError.ReturnError("mdlTransferCoefficient.MembraneDistillationMassTransfer:" & _
                     " The air gap mass transfer is stagnating! You need to figure out why!", , True, True)
       End If
    End If
    
    If glbEvaluateEquationsAtEnd Then
       If IsFinalRun And IsFinalNewtonIteration Then
       
         DebugString = DebugString & FormatDebugColumns("tau   : Membrane tortuosity:     " & tou & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("Pa    : Partial pressure of Air in the membrane (Pa):     " & p_a & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("Kn    : Knudsen number:   " & Knudsen & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("lambda    : mean free path of water vapor molecules in the MD Membrane:   " & lamda & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("d : Membrane mean pore diameter:  " & 2 * MeanPoreRadius & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("C : Membrane transfer coefficient (kg/m2/s/Pa):   " & K_j & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("PD    : Water Vapor Diffusivity (Pa*m2/s):    " & PD & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("p_mean    : Mean pressure (Pa):   " & Pavg & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("pv_ave    : Average partial pressure of vapor in Membrane:    " & (Pvh + Pvc) / 2 & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("pv_h    : Hot side vapor pressure at  interface temp,salin:    " & Pvh & vbCrLf)
         DebugString = DebugString & FormatDebugColumns("pv_c    : Cold side vapor pressure at  interface temp,salin:    " & Pvc & vbCrLf)
         Debug.Print DebugString
         mdlValidation.WriteDebugInfoToFile DebugString
       End If
    End If

EndOfFunction:

Exit Function
ErrorHappened:
            MembraneDistillationMassTransfer = mdlConstants.glbINVALID_VALUE
GoTo EndOfFunction
End Function

Public Function EffectivenessHeatTransferRate(FlowType As Long, Th As Double, Tc As Double, mH As Double, mC As Double, cph As Double, cpc As Double, U As Double, Area As Double, _
                                             Optional HotMixed As Boolean = True, Optional ColdMixed As Boolean = True)

Dim NTU As Double ' Number of transfer units
Dim Cstar As Double ' Cmin / Cmax
Dim Chot As Double ' hot thermal capacity rate (mass flow rate * specific heat)
Dim Ccold As Double ' cold thermal capacity rate
Dim Cmin As Double ' minimum thermal capacity rate
Dim Cmax As Double ' maximum thermal capacity rate
Dim epsilon As Double 'effectiveness
Dim gamma As Double ' Function of NTU's
' Flow type 1 = coflow, 2 = counter flow, 3 = cross flow
' Th = hot side inflow temperature (K)
' Tc = cold side inflow temperature (K)
' mH = hot side mass flow rate (kg/s)
' mC = cold side mass flow rate (kg/s)
' cph = hot side specific heat (J/(kg*K))
' cpc = cold side specific heat (J/(kg*K))
' U = overall heat transfer coefficient W/(m^2 * K)
' Area = heat exchange area
' HotMixed - for cross flow heat exchangers only - Indicates if hot side is mixed or unmixed
' ColdMixed - for cross flow heat exchangers only - Indicates if cold side is mixed or unmixed
Chot = mH * cph
Ccold = mC * cpc

Cmin = Application.WorksheetFunction.Min(Chot, Ccold)
Cmax = Application.WorksheetFunction.Max(Chot, Ccold)

Cstar = Cmin / Cmax

NTU = U * Area / Cmin

Select Case FlowType
   Case 1 'coflow
       ' From Michael R. Lindburg "Mechanical Engineering Reference Manual Twelth Edition" Page 36-21.
       epsilon = (1 - Exp(-NTU * (1 + Cstar))) / (1 + Cstar)
   Case 2
       ' From Michael R. Lindburg "Mechanical Engineering Reference Manual Twelth Edition" Page 36-21.
       epsilon = (1 - Exp(-NTU * (1 - Cstar))) / (1 - Cstar * Exp(-NTU * (1 - Cstar)))
   Case 3 ' From  Lindon C. Thomas, Heat Transfer 2nd edition Page 799
       If HotMixed And ColdMixed Then
           epsilon = 1 / ((1 / (1 - Exp(-NTU))) + (Cstar / (1 - Exp(-Cstar * NTU))) - 1 / NTU)
       ElseIf HotMixed Or ColdMixed Then
           gamma = 1 - Exp(-NTU)
           epsilon = (1 - Exp(-gamma * Cstar)) / Cstar
       Else
           epsilon = BothUnmixedCrossFlowEpsilon(NTU, Cstar)
       End If
   Case Else
       'NTU is just being used as a dummy variable here.
       NTU = mdlError.ReturnError("mdlTransferCoefficient.EffectivenessHeatTransferRate: Input ""FlowType"" must be an integer equal to 1, 2, or 3. " _
                                  & CStr(FlowType) & " was entered!", , True)
End Select

EffectivenessHeatTransferRate = epsilon * Cmin * (Th - Tc)


End Function

Private Function BothUnmixedCrossFlowEpsilon(NTU As Double, Cstar As Double) As Double
   ' From  Lindon C. Thomas, Heat Transfer 2nd edition Page 799
   ' for an unmixed, unmixed cross flow heat exchanger.
   Dim gamma As Double
   Dim epsilon As Double
   Dim epsilon_m1 As Double
   Dim j As Long
   Dim Sum1 As Double
   Dim Sum2 As Double
   Dim Intermediat As Double
   Const MaxTerm = 20
   Const Tolerance = 0.000001
   
   If NTU < 5 Then ' Approximation
      gamma = NTU ^ -0.22
      epsilon = 1 - Exp((Exp(-gamma * Cstar * NTU) - 1) / (gamma * Cstar))
   Else
      epsilon = 0
      Term = 0
      Sum1 = 0
      Sum2 = 0
      Do
          Intermediat = NTU ^ j / (Application.WorksheetFunction.Fact(j))
          Sum1 = Sum1 + Intermediate
          Sum2 = Sum2 + Intermediate * Cstar ^ j
          epsilon = epsilon + (1 - Exp(-NTU) * Sum1) * (1 - Exp(-Cstar * NTU) * Sum2) / (Cstar * NTU)
          j = j + 1
      Loop Until epsilon - epsilon_m1 < Tolerance Or j > MaxTerm
   End If
   
   BothUnmixedCrossFlowEpsilon = epsilon
   
End Function

Public Function InterfaceSalinityConcentration(Tavg As Double, Mavg As Double, Savg As Double, mM As Double, Area As Double, _
                                               LengthNormalToHotFlow As Double, LengthParallelToHotFlow As Double, Thickness As Double, Spacer As clsSpacer, Optional DescStr As String = "") As Double
       Dim Mass_K As Double 'Mass Transfer coefficient
       Dim Mw As Double 'Molecular weight of saline water
       Dim rho As Double 'density of saline water
       Dim NmM As Double ' molar membrane distillation mass flow

       ' Calculate the increased salinity concentration at the interface on the hot side
       Mass_K = mdlTransferCoefficient.K_MassTransferCoef(Tavg, Mavg, Savg, LengthNormalToHotFlow, Thickness, LengthParallelToHotFlow, Spacer, DescStr)

       Mw = mdlProperties.MolarWeightOfWaterNaClMixture(Savg)
       NmM = mM / (Mw * Area)
       rho = mdlProperties.SeaWaterDensity(Tavg, Savg)
       
       If mdlError.NoError Then
       ' From Hitsov, 2015 Separation and Purification Technology 142:pp 48-64 Equation (18)
          InterfaceSalinityConcentration = Savg * Exp(Mw * NmM / (rho * Mass_K))
       Else
          GoTo ErrorOccured
       End If

EndOfFunction:
     
Exit Function
ErrorOccured:
      InterfaceSalinityConcentration = mdlConstants.glbINVALID_VALUE
GoTo EndOfFunction
End Function

Public Function MaximumHeatFlow(T_hot As Double, T_cold As Double, m_hot As Double, m_cold As Double, _
                                mMD As Double, cpc As Double, cph As Double, Lavg As Double, Optional ExcludeLatentHeat As Boolean = True) As Double

' This function calculates the maximum heat flow given input temperatures, mass flows, and salinities into a heat exchanger
' this bound for heat flow can be used to make sure that the 2nd law of thermodynamics is not violated by calculation of an unrealistic heat flow
'
' The fact that latent heat transfer is occuring limits the heat flow a little bit more than a conventional heat exchanger (no mass exchange)
' Setting mMD to zero

Dim T_ref As Double
Dim Qmax As Double
Dim Q_cold_max As Double
Dim Q_hot_max As Double
Dim Qlat As Double
Dim Ecold As Double
Dim Ehot As Double

T_ref = mdlConstants.glbNISTReferenceTemperature
'
'cph = mdlProperties.SeaWaterSpecificHeat(T_hot, S_hot)
'cpc = mdlProperties.SeaWaterSpecificHeat(T_cold, S_cold)
'Lavg = mdlProperties.LatentHeatOfPureWater((T_hot + T_cold) / 2)

Ecold = cpc * (T_cold - T_ref)
Ehot = cph * (T_hot - T_ref)

If ExcludeLatentHeat Then 'Exlude from the maximum heat flow rate! Not from the calculation
   
   Qlat = mMD * (Lavg + Ehot - Ecold)
   
Else

   Qlat = 0

End If

'!@#$ BE CAREFUL WITH THIS PROCEDURE, THE LATENT HEAT IS EXLCUDED FROM THIS

' Cool down all of the hot input water to the cold input temperature

Q_hot_max = m_hot * Ehot - (m_hot - mMD) * Ecold - Qlat
' Heat up all of the cold input water to the hot input temperature

Q_cold_max = (m_cold + mMD) * Ehot - m_cold * Ecold - Qlat

' The maximum heat flow is the mininum of these two maximums (i.e. one stream is the heat limiting stream)
MaximumHeatFlow = mdlMath.DblMin(Q_hot_max, Q_cold_max)

End Function

Public Sub EstimateOutputTemperatures(T_hot As Double, T_cold As Double, m_hot As Double, m_cold As Double, S_hot As Double, S_cold As Double, _
                                mMD As Double, cpc As Double, cph As Double, Lavg As Double, T_hot_out As Double, T_cold_out As Double, _
                                Optional ExcludeLatentHeat As Boolean = True)
' This routine solves for an output temperature that satisfies an assumed heat flux Qactual.  Qactual is a fraction of the maximum heat flow
' T_hot - hot input temperature (K)
' T_cold - cold input temperature (K)
' m_hot - hot input mass flow
' m_cold - cold input mass flow
' S_hot - hot input salinity
' S_cold - cold input salinity
' mMD - membrane distillation mass flow
' cpc - cold input flow specific heat
' cph - hot input flow specific heat
' Lavg - Latent heat of vaporization at average of hot and cold flows
' T_hot_out - estimated hot output temperature
' T_cold_out - estimated cold output temperature

Dim Qactual As Double
Dim Tactual As Double
Dim diff As Double



Dim T_ref As Double

T_ref = mdlConstants.glbNISTReferenceTemperature

Qactual = mdlConstants.glbMaxHeatFlowFractionForInitialCondition * MaximumHeatFlow(T_hot, T_cold, m_hot, m_cold, mMD, cpc, cph, Lavg, ExcludeLatentHeat)

' First guess Temperature (should be pretty close)
Tactual = T_hot - Qactual / (m_hot * cph)
' Avoid calculating this constant term over and over again
diff = (m_hot * cph * (T_hot - T_ref) - Qactual) / (m_hot - mMD)

' Solve the nonlinear equation by successive substitution
T_hot_out = SuccessiveSubSolutionForTemperature(diff, S_hot, Tactual, T_ref)

'Now do the cold side.
' First guess (should not be too far off)
Tactual = T_cold + Qactual / (m_cold * cpc)
' Avoid calculating this constant term over and over again
diff = (m_cold * cpc * (T_cold - T_ref) + Qactual) / (m_cold + mMD)

' Solve the nonlinear equation by successive substition
T_cold_out = SuccessiveSubSolutionForTemperature(diff, S_cold, Tactual, T_ref)

If T_cold_out < T_cold Or T_hot_out > T_hot Then
   T_cold_out = mdlError.ReturnError("mdlTransferCoefficient.EstimateOutputTemperatures: The cold output temperature is lower" & _
                        " than the cold input temperature OR the hot output temperature is greater than the hot input temperature." & _
                        " Either way that is a violation of the 2nd law of thermodynamics that is unacceptable.", , True)
End If

End Sub

Private Function SuccessiveSubSolutionForTemperature(diff As Double, S As Double, Tactual As Double, Tref As Double) As Double

    Dim Iter As Long
    Dim Tactual_m1 As Double
    Dim cp_actual As Double
    Const MaxIter = 25
    
    Iter = 1
    
    Do
        ' cp evaluation is expensive!
        cp_actual = mdlProperties.SeaWaterSpecificHeat(Tactual, S)
        Tactual_m1 = Tactual
        Tactual = diff / cp_actual + Tref
        Iter = Iter + 1
    Loop Until mdlMath.AbsDbl(Tactual - Tactual_m1) < 0.0001 Or Iter > MaxIter
    
    If Iter > MaxIter Then
       SuccessiveSubSolutionForTemperature = mdlError.ReturnError("mdlTransferCoefficient.SuccessiveSubSolutionForTemperature: The output temperature was not found by" & _
                            " successive substitution. A new algorithm is needed or there is a bug!", , True)
    Else
       SuccessiveSubSolutionForTemperature = Tactual
    End If

End Function
