Attribute VB_Name = "mdlProperties"
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

' References
'
' Lindberg, Michael R. 2006. Mechanical Engineering Reference Manual for the PE Exam 12th Edition. Professional Publishing Incorporated
'        Belmont, California (www.ppi2pass.com) ISBN 978-1-59126-049-3.

' Tsilingiris, P.T. 2008. "Thermophysical and transport properties of humid air at temperature range between 0 and 100 degrees Celcius."
'        Energy Conversion and Management 49: 1098-1110
'

Public Function CompressibilityFactorHumidAir(Temperature As Double) As Double
   'Tsilingiris, P.T, 2008 equation 14
Dim T As Double
Dim A As Double
Dim B As Double
Dim Psv As Double
Const C1 = 0.000000007
Const C2 = -0.00000000147184
Const C3 = 1734.29
Const K1 = 1.04E-15
Const K2 = -3.35297E-18
Const K3 = 3645.09

T = Temperature

If T < 273.15 Or T > 373.15 Then
   CompressibilityFactorHumidAir = mdlError.ReturnError("mdlProperties.CompressibilityFactorHumidAir: The temperature supplied " _
                                   & CStr(T) & " Kelvin is out side the valid range of 273.15 to 373.15K", , True)
Else
   A = C1 + C2 * Exp(C3 / T)
   B = K1 + K2 * Exp(K3 / T)
   Psv = mdlProperties.SaturatedPressurePureWater(T)
   CompressibilityFactorHumidAir = 1 + A * Psv + B * Psv ^ 2
End If

End Function

Public Function AirWaterSaturatedMixtureDensity(Temperature As Double, Pressure As Double)
' Tsilingiris, 2008 - equation 11

' This function was tested against 2 points of a psychometric chart and produced similar results.

Dim xv As Double
Dim zm As Double 'compressibility of humid air
Dim Pv As Double
Dim Mratio As Double
Dim T As Double
Dim P0 As Double

T = Temperature
P0 = Pressure

If T < 273.15 Or T > 373.15 Then
   AirWaterSaturatedMixtureDensity = mdlError.ReturnError("mdlProperties.AirWaterSaturatedMixtureDensity: The temperature supplied " _
                                   & CStr(T) & " Kelvin is out side the valid range of 273.15 to 373.15K", , True)
Else
   zm = mdlProperties.CompressibilityFactorHumidAir(T)
   xv = mdlProperties.WaterVaporMoleFractionSaturatedAir(T, P0)
   ' This is for saturated only!
   Pv = mdlProperties.SaturatedPressurePureWater(T)
   Mratio = mdlConstants.glbWaterMolecularWeight / mdlConstants.glbAirMolecularWeight
   ' Pv is Psv since these are saturated conditions.
   AirWaterSaturatedMixtureDensity = (1 / zm) * P0 * glbAirMolecularWeight / (glbGasConstant * T) * (1 - xv * (1 - Mratio))
End If

End Function

Public Function SaturatedThermalConductivityWaterVapor(Temperature As Double)
'Temperature is in Kelvin output is in W/m-K
' valid range from 274K absolute to 640K with errors less than +/- 0.01 W/m-K
' Code was verified for 2 points
' this is from NIST REFPROP database
If Temperature < 274 Or Temperature > 640 Then
    SaturatedThermalConductivityWaterVapor = mdlError.ReturnError("mdlProperties.SaturatedThermalConductivityWaterVapor: Invalid Temperature of " _
             & CStr(Temperature) & " which must be between 274 and 640 Pascal" & _
             " applied to this polynomial relationship.")
    
Else
    Dim T As Double
    Dim A1 As Double
    Dim A2 As Double
    Dim A3 As Double
    Dim a4 As Double
    Dim a5 As Double
    Dim a6 As Double
    Dim a7 As Double
    Dim a8 As Double
    Dim B1 As Double
    Dim B2 As Double
    Dim B3 As Double
    Dim b4 As Double
    Dim b5 As Double
    Dim b6 As Double
    Dim b7 As Double
    Dim b8 As Double
    Dim C1 As Double
    Dim C2 As Double
    Dim C3 As Double
    Dim c4 As Double
    Dim c5 As Double
    Dim c6 As Double
    Dim c7 As Double
    Dim c8 As Double
    ' 8th order Guassian from cftool in Matlab
    A1 = 5524882.48
    B1 = 724.160538
    C1 = 23.8949549
    A2 = 19671.3517
    B2 = 747.550996
    C2 = 44.7708448
    A3 = 548.918199
    B3 = 732.73607
    C3 = 65.5242359
    a4 = 43.3996146
    b4 = 692.216494
    c4 = 89.0685851
    a5 = 55.321835
    b5 = 720.458221
    c5 = 190.681792
    a6 = 0.276753984
    b6 = 535.011009
    c6 = 0.401644615
    b7 = 559.721538
    c7 = 2.22044605E-14
    a8 = 32.6976229
    b8 = 653.935915
    c8 = 461.041588

    T = Temperature
    ' MS_EDIT
    SaturatedThermalConductivityWaterVapor = (A1 * Exp(-((T - B1) / C1) ^ 2) + A2 * Exp(-((T - B2) / C2) ^ 2) + _
              A3 * Exp(-((T - B3) / C3) ^ 2) + a4 * Exp(-((T - b4) / c4) ^ 2) + _
              a5 * Exp(-((T - b5) / c5) ^ 2) + a6 * Exp(-((T - b6) / c6) ^ 2) + _
              a7 * Exp(-((T - b7) / c7) ^ 2) + a8 * Exp(-((T - b8) / c8) ^ 2)) / 1000

End If

End Function

Public Function SaturatedTemperaturePureWater(Pressure As Double)
'Pressure is in pascal output is in Kelvin
' valid range from 611Pa absolute to 22090000Pa
' this is from pages 723 - 724 of  Moran and Shapiro, "Fundamentals of Engineering Thermodynamics" Third Edition
'  the data has been entered in a table in excel in the "WaterProperties.xlsx" workbook and curve fit
If Pressure < 611 Or Pressure > 22090000 Then
    SaturatedTemperaturePureWater = mdlError.ReturnError("mdlProperties.SaturatedTemperaturePureWater: Invalid Pressure of " _
             & CStr(Pressure) & " which must be between 611 and 22090000 Pascal" & _
             " applied to this polynomial relationship.")
    
Else

    Dim coef() As Double
    
    ReDim coef(0 To 6)
    coef(6) = -0.0001138013
    coef(5) = 0.008618524
    coef(4) = -0.2472544
    coef(3) = 3.660487
    coef(2) = -28.95103
    coef(1) = 129.2319
    coef(0) = 2.133888
    
    SaturatedTemperaturePureWater = mdlMath.Polynomial(coef, Log(Pressure))

End If

End Function

Public Function SaturatedLiquidToVaporEnthalpyPureWater(Temperature As Double)

'Pressure is in pascal output is in Kelvin
' valid range from 275K to 600K
' This is from the NIST REFPROP database.
If Temperature < 270 Or Temperature > 600 Then
     
    SaturatedLiquidToVaporEnthalpyPureWater = mdlError.ReturnError("mdlProperties.SaturatedLiquidToVaporEnthalpyPureWater: Invalid Temperature of " _
             & CStr(Temperature) & " which must be between 275 and 600 Kelvin" & _
             " applied to this polynomial relationship.")
    
Else

    Dim coef() As Double
    
    ReDim coef(0 To 9)
    
    coef(9) = -2.50024929E-15
    coef(8) = 9.94077592E-12
    coef(7) = -1.74156156E-08
    coef(6) = 0.000017640385
    coef(5) = -0.0113818593
    coef(4) = 4.85010833
    coef(3) = -1364.72674
    coef(2) = 244481.292
    coef(1) = -25302161#
    coef(0) = 1155384540#
    
    SaturatedLiquidToVaporEnthalpyPureWater = mdlMath.Polynomial(coef, Temperature)

End If
End Function


Public Function SaturatedVaporEnthalpyPureWater(Temperature As Double)

' Output is in J/kg
' valid range from 275K to 600K
' This is from the NIST REFPROP database.
If Temperature < 275 Or Temperature > 600 Then
     
    SaturatedVaporEnthalpyPureWater = mdlError.ReturnError("mdlProperties.SaturatedVaporEnthalpyPureWater: Invalid Temperature of " _
             & CStr(Temperature) & " which must be between 275 and 600 Kelvin" & _
             " applied to this polynomial relationship.")
    
Else

    Dim coef() As Double
    
    ReDim coef(0 To 9)
    
    coef(9) = -1.50094973E-15
    coef(8) = 5.96598746E-12
    coef(7) = -1.04491551E-08
    coef(6) = 0.0000105809713
    coef(5) = -0.00682487784
    coef(4) = 2.90724215
    coef(3) = -817.719281
    coef(2) = 146423.801
    coef(1) = -15143123.2
    coef(0) = 691373990#
    
    SaturatedVaporEnthalpyPureWater = mdlMath.Polynomial(coef, Temperature)

End If
End Function

Public Function DryAirEnthalpy(Temperature As Double, Pressure As Double) As Double

' Input temperature in Kelvin
' Input Pressure in Pascal

'FROM NIST REFPROP
'Lemmon, E.W., R.T Jacobsen, S.G. Penoncello, and D.G. Friend. 2000. "Thermodynamic Properties of Air and Mixtures of Nitrogen, Argon, and Oxygen from 60 to 2000 K at Pressures to 2000 MPa." J Phys. Chem. Ref. Data, 29(3):331-385.
'Linear model Poly31:
'     f(x, Y) = p00 + p10 * x + p01 * Y + p20 * x ^ 2 + p11 * x * Y + p30 * x ^ 3 + p21 * x ^ 2 * Y
'Coefficients (with 95% confidence bounds):
'       p00 =   1.234e+05  (1.234e+05, 1.235e+05)
'       p10 =        1028  (1028, 1029)
'       p01 =    -0.01078  (-0.01086, -0.0107)
'       p20 =     -0.1013  (-0.1024, -0.1001)
'       p11 =   4.069e-05  (4.021e-05, 4.116e-05)
'       p30 =   0.0001375  (0.0001364, 0.0001386)
'       p21 =  -4.133e-08  (-4.201e-08, -4.065e-08)
'
'Goodness of fit:
'  SSE: 1.694e+04
'  R-square: 1
'  Adjusted R-square: 1
'  RMSE: 3.444
Const p00 = 123400#      '(1.234e+05, 1.235e+05)
Const p10 = 1028         '(1028, 1029)
Const p01 = -0.01078     '(-0.01086, -0.0107)
Const p20 = -0.1013      '(-0.1024, -0.1001)
Const p11 = 0.00004069   '(4.021e-05, 4.116e-05)
Const p30 = 0.0001375    '(0.0001364, 0.0001386)
Const p21 = -0.00000004133 '(-4.201e-08, -4.065e-08)


If Temperature < 250 Or Temperature > 450 Then
    DryAirEnthalpy = mdlError.ReturnError("mdlProperties.DryAirEnthalpy: Invalid Temperature of " _
             & CStr(Temperature) & " which must be between 250 and 450 Kelvin" & _
             " applied to this polynomial relationship.")
ElseIf Pressure < 1 Or Pressure > 250000 Then
    DryAirEnthalpy = mdlError.ReturnError("mdlProperties.DryAirEnthalpy: Invalid Pressure of " _
             & CStr(Pressure) & " which must be between 1 and 250,000 Pascals" & _
             " applied to this polynomial relationship.")
Else

    DryAirEnthalpy = p00 + p10 * Temperature + p01 * Pressure + p20 * Temperature ^ 2 _
                   + p11 * Temperature * Pressure + p30 * Temperature ^ 3 + p21 * Temperature ^ 2 * Pressure

End If
End Function


Public Function SaturatedLiquidEnthalpyPureWater(Temperature As Double)
'Pressure is in pascal output is in J/kg
' valid range from 275K to 600K
' This is from the NIST REFPROP database.
If Temperature < 275 Or Temperature > 600 Then
     
    SaturatedLiquidEnthalpyPureWater = mdlError.ReturnError("mdlProperties.SaturatedLiquidEnthalpyPureWater: Invalid Temperature of " _
             & CStr(Temperature) & " which must be between 275 and 600 Kelvin" & _
             " applied to this polynomial relationship.")
    
Else

    Dim coef() As Double
    
    ReDim coef(0 To 9)
    
    coef(9) = 9.99299563E-16
    coef(8) = -3.97478846E-12
    coef(7) = 6.96646051E-09
    coef(6) = -0.00000705941371
    coef(5) = 0.00455698146
    coef(4) = -1.94286618
    coef(3) = 547.007459
    coef(2) = -98057.4908
    coef(1) = 10159037.8
    coef(0) = -464010550#
    
    SaturatedLiquidEnthalpyPureWater = mdlMath.Polynomial(coef, Temperature)

End If

End Function


Public Function SaturatedPressurePureWater(Temperature As Double)
'Pressure is in pascal output is in Kelvin
' valid range from 275K to 475K
' This is from the NIST REFPROP database.
' see ./MaterialProp/SaturationTableRefProp.xlsx
If Temperature < 275 Or Temperature > 475 Then
     
    SaturatedPressurePureWater = mdlError.ReturnError("mdlProperties.SaturatedPressurePureWater: Invalid Temperature of " _
             & CStr(Temperature) & " which must be between 275 and 475 Kelvin" & _
             " applied to this polynomial relationship.", , True)
    
Else

    Dim coef() As Double
    
    ReDim coef(0 To 6)
    
    coef(6) = -7.08535E-15
    coef(5) = 1.85055E-11
    coef(4) = -0.0000000206119
    coef(3) = 0.0000126862
    coef(2) = -0.00464543
    coef(1) = 1.0007
    coef(0) = -89.3226
    
    'The polynomial fit is for Log(Pressure)
    SaturatedPressurePureWater = Exp(mdlMath.Polynomial(coef, Temperature))

End If

End Function

Public Function DensityOfSaturatedSteamPureWater(Temperature As Double)

DensityOfSaturatedSteamPureWater = mdlProperties.DensitySaturatedLiquidPureWater(Temperature) - mdlProperties.DensityChangeOnCondensationPureWater(Temperature)

End Function

Public Function DensitySaturatedLiquidPureWater(Temperature As Double)
'Temperature is in Kelvin, Output is in kg/m3
' valid range from 275K absolute to 475K.
' The data was fit to a 6th order polynomial from the NIST REFPROP database

' see excel spreadsheet ./MaterialProp/SaturationTableRefProp.xlsx for data table and fit characteristics

If Temperature < 275 Or Temperature > 475 Then
    
    DensitySaturatedLiquidPureWater = mdlError.ReturnError("mdlProperties.DensitySaturatedLiquidPureWater:" & _
             " Invalid temperature of " & CStr(Temperature) & " which must be between 275 and 475 Kelvin" & _
             " applied to this polynomial relationship")
    
Else

    Dim coef() As Double
    
    ReDim coef(0 To 6)
    coef(6) = -1.1255E-12
    coef(5) = 0.00000000269192
    coef(4) = -0.00000269372
    coef(3) = 0.00144435
    coef(2) = -0.440045
    coef(1) = 72.0883
    coef(0) = -3925.19
    
    DensitySaturatedLiquidPureWater = mdlMath.Polynomial(coef, Temperature)

End If

End Function

Public Function DensityChangeOnCondensationPureWater(Temperature As Double)
'Temperature is in Kelvin, Output is in kg/m3
' valid range from 275K absolute to 475K.
' The data was fit to a 6th order polynomial from the NIST REFPROP database

' see excel spreadsheet ./MaterialProp/SaturationTableRefProp.xlsx for data table and fit characteristics

If Temperature < 275 Or Temperature > 475 Then
    
    DensityChangeOnCondensationPureWater = mdlError.ReturnError("mdlProperties.DensityChangeOnCondensationPureWater:" & _
             " Invalid temperature of " & CStr(Temperature) & " which must be between 275 and 475 Kelvin" & _
             " applied to this polynomial relationship")
    
Else

    Dim coef() As Double
    
    ReDim coef(0 To 6)
    coef(6) = -1.14453E-12
    coef(5) = 0.00000000272609
    coef(4) = -0.00000272368
    coef(3) = 0.00145942
    coef(2) = -0.444379
    coef(1) = 72.7468
    coef(0) = -3966.18
    
    DensityChangeOnCondensationPureWater = mdlMath.Polynomial(coef, Temperature)

End If

End Function

Public Function LatentHeatOfPureWater(Temperature As Double)
'Temperature is in Kelvin, Output is in J/kg
' valid range from 275K absolute to 475K.
' The data was fit to a 6th order polynomial from the NIST REFPROP database
' see excel spreadsheet ./MaterialProp/SaturationTableRefProp.xlsx for data table and fit characteristics

If Temperature < 275 Or Temperature > 475 Then
    
    LatentHeatOfPureWater = mdlError.ReturnError("mdlProperties.LatentHeatOfPureWater:" & _
             " Invalid temperature of " & CStr(Temperature) & " which must be between 275 and 475 Kelvin" & _
             " applied to this polynomial relationship")
    
Else

    Dim coef() As Double
    
    ReDim coef(0 To 6)
    coef(6) = 0.000000000392177
    coef(5) = -0.000000897818
    coef(4) = 0.000826231
    coef(3) = -0.406878
    coef(2) = 114.29
    coef(1) = -19679.4
    coef(0) = 4244030#
    
    LatentHeatOfPureWater = mdlMath.Polynomial(coef, Temperature)

End If

End Function

Public Function MolarWeightOfWaterNaClMixture(Salinity As Double) As Double

If Salinity > 1 Or Salinity < 0 Then ' We have an invalid input!
    MolarWeightOfWaterNaClMixture = mdlError.ReturnError("mdlProperties.MolarWeightOfWaterNaClMixture: Salinity must be between 0 and 1!")
Else
    ' Only H2O and NaCl are considered here.  All other chemical species are neglected.
    MolarWeightOfWaterNaClMixture = mdlConstants.glbWaterMolecularWeight * (1 - Salinity) + mdlConstants.glbSodiumChlorideMolecularWeight * Salinity
End If

End Function

Public Function SuperHeatedSteamEnthalpy(Pressure As Double, Temperature As Double, Optional RefName As String = "NIST REFPROP")

Dim coef() As Double
Dim i As Long
Dim ValidNames() As String
ReDim ValidNames(1)
ValidNames(0) = "MoranoAndShapiro"
ValidNames(1) = "NIST REFPROP"

Select Case RefName
    Case ValidNames(0)

        ' Maximum Error within valid range is +/-100 J/kg from the original table points
        ' Polynomial curve fit From steam tables in "Fundamentals of Engineering Thermodynamics 3rd Edition" by Michael J. Moran and Howard N. Shapiro
        
        If (Pressure < 5000 Or Pressure > 101300) Then
            err.Source = "mdlProperties.SuperHeatedSteamEnthalpy: Pressure " & Pressure & " is outside of the valid range of 5000 to 101300 Pascal!"
            SuperHeatedSteamEnthalpy = mdlConstants.glbINVALID_VALUE
        ElseIf (Temperature < 309 Or Temperature > 800) Then
            err.Source = "mdlProperties.SuperHeatedSteamEnthalpy: Temperature " & Temperature & " is outside of the valid range of 309 to 800 Kelvin!"
            SuperHeatedSteamEnthalpy = mdlConstants.glbINVALID_VALUE
        Else
            
            ReDim coef(0 To 5, 0 To 1)
            
            coef(0, 0) = 1944121.72
            coef(0, 1) = -2.98075361
            coef(1, 0) = 2414.6914
            coef(1, 1) = 0.0184717984
            coef(2, 0) = -2.39214029
            coef(2, 1) = -0.0000436313053
            coef(3, 0) = 0.00476951216
            coef(3, 1) = 0.000000046001353
            coef(4, 0) = -0.00000422879777
            coef(4, 1) = -1.81641008E-11
            coef(5, 0) = 1.50159966E-09
            coef(5, 1) = 0#
            
            SuperHeatedSteamEnthalpy = coef(0, 0) + coef(1, 0) * Temperature + _
                                       coef(0, 1) * Pressure + coef(2, 0) * Temperature ^ 2 + _
                                       coef(1, 1) * Temperature * Pressure + _
                                       coef(3, 0) * Temperature ^ 3 + _
                                       coef(2, 1) * Temperature ^ 2 * Pressure + _
                                       coef(4, 0) * Temperature ^ 4 + _
                                       coef(3, 1) * Temperature ^ 3 * Pressure + _
                                       coef(5, 0) * Temperature ^ 5 + _
                                       coef(4, 1) * Temperature ^ 4 * Pressure
        End If
    Case ValidNames(1)
        ' More extensive fit from NIST REFPROP database - error +/- 200 J/kg for range in 2.5e6 to 3.5e6 J/kg
        Dim Tsat As Double
        
        Tsat = SaturatedTemperaturePureWater(Pressure)
        
        If Temperature < Tsat - mdlConstants.glbSatErr Then
            err.Source = "mdlProperties.SuperHeatedSteamEnthalpy: The temperature input (" & Temperature & _
                         ") for super-heated steam properties is below the saturation temperature (" & Tsat & _
                         ") for the pressure input (" & Pressure & ")! The requested point is therefore for sub-cooled liquid water!"
            SuperHeatedSteamEnthalpy = mdlConstants.glbINVALID_VALUE
        ElseIf Temperature > 773 Then
            err.Source = "mdlProperties.SuperHeatedSteamEnthalpy: The temperature input (" & Temperature & _
                         ") is beyond the upper limit of 773Kelvin!"
            SuperHeatedSteamEnthalpy = mdlConstants.glbINVALID_VALUE
        ElseIf Pressure < 612.5 Or Pressure > 200000 Then
            err.Source = "mdlProperties.SuperHeatedSteamEnthalpy: The pressure input (" & Pressure & _
                         ") is outside the valid range of 612.5 To 200000Pascal!"
            SuperHeatedSteamEnthalpy = mdlConstants.glbINVALID_VALUE
        Else
            ReDim coef(0 To 5, 0 To 4)
            coef(0, 0) = 1834188.60120496
            coef(0, 1) = 82075.6110008664
            coef(0, 2) = -18161.2732800223
            coef(0, 3) = 1773.6708090033
            coef(0, 4) = -80.1118264605804
            coef(1, 0) = 2016.75126183812
            coef(1, 1) = 14.7650257939712
            coef(1, 2) = 5.41996295050809
            coef(1, 3) = -0.388125077781118
            coef(1, 4) = 0.101069850533474
            coef(2, 0) = -0.939265956316152
            coef(2, 1) = -0.102299840248471
            coef(2, 2) = -4.02998563879419E-03
            coef(2, 3) = -2.43029357561529E-03
            coef(2, 4) = 0
            coef(3, 0) = 2.37140679980587E-03
            coef(3, 1) = 1.6725252718778E-04
            coef(3, 2) = 3.62614142283201E-05
            coef(3, 3) = 0
            coef(3, 4) = 0
            coef(4, 0) = -2.47039531790377E-06
            coef(4, 1) = -3.23917895379505E-07
            coef(4, 2) = 0
            coef(4, 3) = 0
            coef(4, 4) = 0
            coef(5, 1) = 1.68805059870015E-09
            
            'Output in J/kg
            SuperHeatedSteamEnthalpy = Polynomial2D(coef, Temperature, Log(Pressure))
        End If
    Case Else
        SuperHeatedSteamEnthalpy = mdlError.NameError("mdlProperties.SuperHeatedSteamEnthalpy", _
                                                    "RefName", _
                                                    ValidNames)
End Select
        
End Function

Public Function SuperHeatedSteamSpecificVolume(Pressure As Double, Temperature As Double, Optional RefName As String = "NIST REFPROP")

Dim coef() As Double
Dim i As Long
Dim logv As Double
Dim ValidNames() As String
ReDim ValidNames(0)
ValidNames(0) = "NIST REFPROP"

Select Case RefName
    Case ValidNames(0)
        ' More extensive fit from NIST REFPROP database - error +/- 4e-4 log(v) for range in 0 to 6 J/kg
        Dim Tsat As Double
        
        Tsat = SaturatedTemperaturePureWater(Pressure)
        
        If Temperature < Tsat - mdlConstants.glbSatErr Then
            err.Source = "mdlProperties.SuperHeatedSteamSpecificVolume: The temperature input (" & Temperature & _
                         ") for super-heated steam properties is below the saturation temperature (" & Tsat & _
                         ") for the pressure input (" & Pressure & ")! The requested point is therefore for sub-cooled liquid water!"
            SuperHeatedSteamSpecificVolume = mdlConstants.glbINVALID_VALUE
        ElseIf Temperature > 773 Then
            err.Source = "mdlProperties.SuperHeatedSteamSpecificVolume: The temperature input (" & Temperature & _
                         ") is beyond the upper limit of 773Kelvin!"
            SuperHeatedSteamSpecificVolume = mdlConstants.glbINVALID_VALUE
        ElseIf Pressure < 612.5 Or Pressure > 200000 Then
            err.Source = "mdlProperties.SuperHeatedSteamSpecificVolume: The pressure input (" & Pressure & _
                         ") is outside the valid range of 612.5 To 200000Pascal!"
            SuperHeatedSteamSpecificVolume = mdlConstants.glbINVALID_VALUE
        Else
            ReDim coef(0 To 5, 0 To 5)
            coef(0, 0) = 9.93759891
            coef(0, 1) = -0.982717763
            coef(0, 2) = -0.0020247233
            coef(0, 3) = -0.000071854836
            coef(0, 4) = 0.0000269118337
            coef(0, 5) = -0.00000279786692
            coef(1, 0) = 0.0108334028
            coef(1, 1) = -0.0000476501648
            coef(1, 2) = 0.00000972176031
            coef(1, 3) = -0.000000878917588
            coef(1, 4) = 0.000000139718104
            coef(2, 0) = -0.0000224101835
            coef(2, 1) = 3.24681855E-09
            coef(2, 2) = -6.04810058E-10
            coef(2, 3) = -3.12388686E-09
            coef(3, 0) = 3.02500457E-08
            coef(3, 1) = 8.0731363E-12
            coef(3, 2) = 4.39926061E-11
            coef(4, 0) = -2.24276965E-11
            coef(4, 1) = -3.07001069E-13
            coef(5, 0) = 7.7455332E-15
            
            'Output in logarithm of v
            logv = Polynomial2D(coef, Temperature, Log(Pressure))
            'output in m3/kg
            SuperHeatedSteamSpecificVolume = Exp(logv)
        End If
    Case Else
        SuperHeatedSteamSpecificVolume = mdlError.NameError("mdlProperties.SuperHeatedSteamSpecificVolume", _
                                                    "RefName", _
                                                    ValidNames)
End Select
        
End Function

Public Function SubCooledWaterSpecificHeat(Pressure As Double, Temperature As Double, _
                                           Optional RefName As String = "NIST REFPROP")

Dim coef() As Double
Dim i As Long
Dim ValidNames() As String
ReDim ValidNames(0)
ValidNames(0) = "NIST REFPROP"

Select Case RefName
    Case ValidNames(0)
        ' More extensive fit from NIST REFPROP database - error +/- 1 J/(kg*K) for range in 4180 to 4240 J/(kg*K)
        Dim Tsat As Double
        
        Tsat = SaturatedTemperaturePureWater(Pressure)
        
        If Temperature > Tsat + mdlConstants.glbSatErr Then
            err.Source = "mdlProperties.SubCooledWaterSpecificHeat: The temperature input (" & Temperature & _
                         ") for subcooled water properties is above the saturation temperature (" & Tsat & _
                         ") for the pressure input (" & Pressure & ")! The requested point is therefore for super-heated vapor!"
            SubCooledWaterSpecificHeat = mdlConstants.glbINVALID_VALUE
        ElseIf Temperature < 274 Then
            err.Source = "mdlProperties.SubCooledWaterSpecificHeat: The temperature input (" & Temperature & _
                         ") is beyond the lower limit of 274Kelvin!"
            SubCooledWaterSpecificHeat = mdlConstants.glbINVALID_VALUE
        ElseIf Pressure < 700 Or Pressure > 200000 Then
            err.Source = "mdlProperties.SubCooledWaterSpecificHeat: The pressure input (" & Pressure & _
                         ") is outside the valid range of 700 To 200000Pascal!"
            SubCooledWaterSpecificHeat = mdlConstants.glbINVALID_VALUE
        Else
            ReDim coef(0 To 5, 0 To 2)
            coef(0, 0) = 133228.645
            coef(0, 1) = -1765.05818
            coef(0, 2) = 56.7024555
            coef(1, 0) = -1822.14992
            coef(1, 1) = 19.011658
            coef(1, 2) = -0.582695346
            coef(2, 0) = 10.2646815
            coef(2, 1) = -0.0712535996
            coef(2, 2) = 0.00199192998
            coef(3, 0) = -0.0288937643
            coef(3, 1) = 0.000103015355
            coef(3, 2) = -0.00000226675226
            coef(4, 0) = 0.0000407654089
            coef(4, 1) = -3.78344779E-08
            coef(5, 0) = -2.31665625E-08
            
            'Output in J/(kg*K)
            SubCooledWaterSpecificHeat = Polynomial2D(coef, Temperature, Log(Pressure))
        End If
    Case Else
        SubCooledWaterSpecificHeat = mdlError.NameError("mdlProperties.SubCooledWaterSpecificHeat", _
                                                    "RefName", _
                                                    ValidNames)
End Select
        
End Function

Public Function SubCooledWaterThermalConductivity(Pressure As Double, Temperature As Double, _
                                           Optional RefName As String = "NIST REFPROP")

Dim coef() As Double
Dim i As Long
Dim ValidNames() As String
ReDim ValidNames(0)
ValidNames(0) = "NIST REFPROP"

Select Case RefName
    Case ValidNames(0)
        ' More extensive fit from NIST REFPROP database - error +/- 1e-4 W/(m*K) for range 0.557 to 0.682 in  W/(m*K)
        Dim Tsat As Double
        
        Tsat = SaturatedTemperaturePureWater(Pressure)
        
        If Temperature > Tsat + mdlConstants.glbSatErr Then
            err.Source = "mdlProperties.SubCooledWaterThermalConductivity: The temperature input (" & Temperature & _
                         ") for subcooled water properties is above the saturation temperature (" & Tsat & _
                         ") for the pressure input (" & Pressure & ")! The requested point is therefore for super-heated vapor!"
            SubCooledWaterThermalConductivity = mdlConstants.glbINVALID_VALUE
        ElseIf Temperature < 274 Then
            err.Source = "mdlProperties.SubCooledWaterThermalConductivity: The temperature input (" & Temperature & _
                         ") is beyond the lower limit of 274Kelvin!"
            SubCooledWaterThermalConductivity = mdlConstants.glbINVALID_VALUE
        ElseIf Pressure < 700 Or Pressure > 200000 Then
            err.Source = "mdlProperties.SubCooledWaterThermalConductivity: The pressure input (" & Pressure & _
                         ") is outside the valid range of 700 To 200000Pascal!"
            SubCooledWaterThermalConductivity = mdlConstants.glbINVALID_VALUE
        Else
            ReDim coef(0 To 5, 0 To 2)
            coef(0, 0) = -20.3185297
            coef(0, 1) = 0.266386556
            coef(0, 2) = -0.00898190624
            coef(1, 0) = 0.282476534
            coef(1, 1) = -0.00284528607
            coef(1, 2) = 0.0000922877018
            coef(2, 0) = -0.0015499239
            coef(2, 1) = 0.0000105072589
            coef(2, 2) = -0.000000315671557
            coef(3, 0) = 0.00000431027359
            coef(3, 1) = -1.46909101E-08
            coef(3, 2) = 3.59754445E-10
            coef(4, 0) = -6.05928747E-09
            coef(4, 1) = 4.70091094E-12
            coef(5, 0) = 3.44919593E-12
            
            'Output in W/(m*K)
            SubCooledWaterThermalConductivity = Polynomial2D(coef, Temperature, Log(Pressure))
        End If
    Case Else
        SubCooledWaterThermalConductivity = mdlError.NameError("mdlProperties.SubCooledWaterThermalConductivity", _
                                                    "RefName", _
                                                    ValidNames)
End Select
        
End Function

Public Function SubCooledWaterDensity(Pressure As Double, Temperature As Double, _
                                           Optional RefName As String = "NIST REFPROP")

Dim coef() As Double
Dim i As Long
Dim ValidNames() As String
ReDim ValidNames(0)
ValidNames(0) = "NIST REFPROP"

Select Case RefName
    Case ValidNames(0)
        ' More extensive fit from NIST REFPROP database - error +.01/-.03 kg/m3 for range in 940 to 1000 kg/m3
        Dim Tsat As Double
        
        Tsat = SaturatedTemperaturePureWater(Pressure)
        
        If Temperature > Tsat + mdlConstants.glbSatErr Then
            err.Source = "mdlProperties.SubCooledWaterDensity: The temperature input (" & Temperature & _
                         ") for subcooled water properties is above the saturation temperature (" & Tsat & _
                         ") for the pressure input (" & Pressure & ")! The requested point is therefore for super-heated vapor!"
            SubCooledWaterDensity = mdlConstants.glbINVALID_VALUE
        ElseIf Temperature < 274 Then
            err.Source = "mdlProperties.SubCooledWaterDensity: The temperature input (" & Temperature & _
                         ") is beyond the lower limit of 274Kelvin!"
            SubCooledWaterDensity = mdlConstants.glbINVALID_VALUE
        ElseIf Pressure < 700 Or Pressure > 200000 Then
            err.Source = "mdlProperties.SubCooledWaterDensity: The pressure input (" & Pressure & _
                         ") is outside the valid range of 700 To 200000Pascal!"
            SubCooledWaterDensity = mdlConstants.glbINVALID_VALUE
        Else
            ReDim coef(0 To 5, 0 To 3)
            coef(0, 0) = -4130.69705
            coef(0, 1) = 56.852855
            coef(0, 2) = -2.1773214
            coef(0, 3) = 0.000825950059
            coef(1, 0) = 68.9418879
            coef(1, 1) = -0.593594772
            coef(1, 2) = 0.022608859
            coef(1, 3) = -0.0000249341667
            coef(2, 0) = -0.369951899
            coef(2, 1) = 0.00208463872
            coef(2, 2) = -0.0000764129176
            coef(2, 3) = 9.83714798E-08
            coef(3, 0) = 0.00100008202
            coef(3, 1) = -0.0000025888428
            coef(3, 2) = 0.000000081941603
            coef(4, 0) = -0.00000137338486
            coef(4, 1) = 4.99440472E-10
            coef(5, 0) = 7.65796484E-10
            
            'Output in kg/m3
            SubCooledWaterDensity = Polynomial2D(coef, Temperature, Log(Pressure))
        End If
    Case Else
        SubCooledWaterDensity = mdlError.NameError("mdlProperties.SubCooledWaterDensity", _
                                                    "RefName", _
                                                    ValidNames)
End Select
        
End Function

Public Function SubCooledWaterPrandtl(Pressure As Double, Temperature As Double, _
                                           Optional RefName As String = "NIST REFPROP")

Dim coef() As Double
Dim i As Long
Dim ValidNames() As String
ReDim ValidNames(0)
ValidNames(0) = "NIST REFPROP"

Select Case RefName
    Case ValidNames(0)
        ' More extensive fit from NIST REFPROP database - error +.1/-.05 (unitless) for range in 2 to 12 unitless
        ' THIS ERROR IS LARGER THAN MANY AND AMOUNTS TO 2.5% in the worst case!
        Dim Tsat As Double
        
        Tsat = SaturatedTemperaturePureWater(Pressure)
        
        If Temperature > Tsat + mdlConstants.glbSatErr Then
            err.Source = "mdlProperties.SubCooledWaterPrandtl: The temperature input (" & Temperature & _
                         ") for subcooled water properties is above the saturation temperature (" & Tsat & _
                         ") for the pressure input (" & Pressure & ")! The requested point is therefore for super-heated vapor!"
            SubCooledWaterPrandtl = mdlConstants.glbINVALID_VALUE
        ElseIf Temperature < 274 Then
            err.Source = "mdlProperties.SubCooledWaterPrandtl: The temperature input (" & Temperature & _
                         ") is beyond the lower limit of 274Kelvin!"
            SubCooledWaterPrandtl = mdlConstants.glbINVALID_VALUE
        ElseIf Pressure < 700 Or Pressure > 200000 Then
            err.Source = "mdlProperties.SubCooledWaterPrandtl: The pressure input (" & Pressure & _
                         ") is outside the valid range of 700 To 200000Pascal!"
            SubCooledWaterPrandtl = mdlConstants.glbINVALID_VALUE
        Else
            ReDim coef(0 To 5, 0 To 4)
            coef(0, 0) = 15185.5094
            coef(0, 1) = -248.355311
            coef(0, 2) = 5.53988968
            coef(0, 3) = 0.069258761
            coef(0, 4) = 0.0151021293
            coef(1, 0) = -212.225315
            coef(1, 1) = 2.83616269
            coef(1, 2) = -0.0654089042
            coef(1, 3) = -0.00256249396
            coef(1, 4) = -0.0000536998038
            coef(2, 0) = 1.18279954
            coef(2, 1) = -0.0113389217
            coef(2, 2) = 0.00036019001
            coef(2, 3) = 0.0000082332671
            coef(3, 0) = -0.00329181692
            coef(3, 1) = 0.0000153687835
            coef(3, 2) = -0.000000702588259
            coef(4, 0) = 0.00000461567095
            coef(4, 1) = 9.05754999E-10
            coef(5, 0) = -2.66775266E-09
            
            'Output in unitless
            SubCooledWaterPrandtl = Polynomial2D(coef, Temperature, Log(Pressure))
        End If
    Case Else
        SubCooledWaterPrandtl = mdlError.NameError("mdlProperties.SubCooledWaterPrandtl", _
                                                    "RefName", _
                                                    ValidNames)
End Select
        
End Function

Public Function SubCooledWaterViscosity(Pressure As Double, Temperature As Double, _
                                           Optional RefName As String = "NIST REFPROP")

Dim coef() As Double
Dim i As Long
Dim ValidNames() As String
ReDim ValidNames(0)
ValidNames(0) = "NIST REFPROP"

Select Case RefName
    Case ValidNames(0)
        ' More extensive fit from NIST REFPROP database - error +1e-5/-1e-5 (Pa-s) for range in 5e-4 to 15e-4 Pa-s
        ' THIS ERROR IS LARGER THAN MANY AND AMOUNTS TO 2% in the worst case!
        Dim Tsat As Double
        
        Tsat = SaturatedTemperaturePureWater(Pressure)
        
        If Temperature > Tsat + mdlConstants.glbSatErr Then
            err.Source = "mdlProperties.SubCooledWaterViscosity: The temperature input (" & Temperature & _
                         ") for subcooled water properties is above the saturation temperature (" & Tsat & _
                         ") for the pressure input (" & Pressure & ")! The requested point is therefore for super-heated vapor!"
            SubCooledWaterViscosity = mdlConstants.glbINVALID_VALUE
        ElseIf Temperature < 274 Then
            err.Source = "mdlProperties.SubCooledWaterViscosity: The temperature input (" & Temperature & _
                         ") is beyond the lower limit of 274Kelvin!"
            SubCooledWaterViscosity = mdlConstants.glbINVALID_VALUE
        ElseIf Pressure < 700 Or Pressure > 200000 Then
            err.Source = "mdlProperties.SubCooledWaterViscosity: The pressure input (" & Pressure & _
                         ") is outside the valid range of 700 To 200000Pascal!"
            SubCooledWaterViscosity = mdlConstants.glbINVALID_VALUE
        Else
            ReDim coef(0 To 5, 0 To 1)
            coef(0, 0) = 1.5460851
            coef(0, 1) = -0.017739743
            coef(1, 0) = -0.0217286988
            coef(1, 1) = 0.000226784225
            coef(2, 0) = 0.000121718473
            coef(2, 1) = -0.00000108328113
            coef(3, 0) = -0.000000338925226
            coef(3, 1) = 2.29123193E-09
            coef(4, 0) = 4.68111877E-10
            coef(4, 1) = -1.81034509E-12
            coef(5, 0) = -2.55970701E-13
            
            'Output in Pa-s
            SubCooledWaterViscosity = Polynomial2D(coef, Temperature, Log(Pressure))
        End If
    Case Else
        SubCooledWaterViscosity = mdlError.NameError("mdlProperties.SubCooledWaterViscosity", _
                                                    "RefName", _
                                                    ValidNames)
End Select
        
End Function

Public Function SeaWaterDensity(Temperature As Double, _
                                  Salinity As Double, _
                                  Optional RefName As String = "Sun_et_al", Optional ThrowError As Boolean = True) As Double

Dim T As Double
Dim S As Double
Dim A() As Double
Dim B() As Double
Dim ValidNames() As String
Dim i As Long

ReDim ValidNames(0)
ValidNames(0) = "Sun_et_al"

Select Case RefName
'You must enter every case here or the function will be inconsistent!
    Case ValidNames(0)
    ' This relation comes from Sharquwy et. al. (2010) reference 28 on page 357.
    ' Mostafa H. Sharqawy, John H. Lienhard V, Syed M. Zubair, "Thermophysical properties
    ' of seawater: a review of existing correlations and data," Desalination and Water Treatment
    ' Vol 16:pp354-380. April 2010.  www.deswater.com

    ' [28] H.Sun , R.Feistel, M.Koch And a.Markoe, New Equations
    'for density, entropy, heat capacity, and potential temperature
    'of a saline thermal fl uid, Deep-Sea Research, I 55 (2008)
    '1304 –1310
                                            
    ' Salinity as kg/kg (fraction) valid range 0 to 0.16
    ' Temperature as Kelvin (convert back to Celcius) valid range 0 < T < 180C
    ' accurancy w/r to IAPWS-2008 formulation +/-0.1%
        T = Temperature - mdlConstants.glbCelciusToKelvinOffset
        S = Salinity
        
        If T < 0 Or T > 180 Then
            SeaWaterDensity = mdlError.ReturnError("mdlProperties.SeaWaterDensity: Temperature " & T & "Celcius is outside the valid range of 0 to 180!", _
                                , ThrowError, True)
        ElseIf S < 0 Or S > 0.16 Then
            SeaWaterDensity = mdlError.ReturnError("mdlProperties.SeaWaterDensity: Salinity " & S & "kg/kg is outside the valid range of 0 to 0.16!", _
                                                    , ThrowError, True)
        Else
            ReDim A(1 To 5)
            ReDim B(1 To 5)
            A(1) = 999.9
            A(2) = 0.02034
            A(3) = -0.006162
            A(4) = 0.00002261
            A(5) = -0.00000004657
            B(1) = 802#
            B(2) = -2.001
            B(3) = 0.01677
            B(4) = -0.0000306
            B(5) = -0.00001613
            SeaWaterDensity = A(1) + A(2) * T + A(3) * T ^ 2 + A(4) * T ^ 3 + _
                                          A(5) * T ^ 4 + B(1) * S + B(2) * S * T + _
                                          B(3) * S * T ^ 2 + B(4) * S * T ^ 3 + _
                                          B(5) * S ^ 2 * T ^ 2
                                
        End If
    Case Else
        SeaWaterDensity = mdlError.NameError("mdlProperties.SeaWaterDensity", _
                                                    "RefName", _
                                                    ValidNames)
End Select

End Function

Public Function SeaWaterSpecificEnthalpy(Temperature As Double, _
                                  Salinity As Double, _
                                  Pressure As Double, _
                                  Optional RefName As String = "Nayar_et_al") As Double
' In Pascals
Dim T As Double
Dim S As Double
Dim P As Double
Dim ValidNames() As String
Dim i As Long

ReDim ValidNames(0)
ValidNames(0) = "Nayar_et_al"

Select Case RefName
'You must enter every case here or the function will be inconsistent!
    Case ValidNames(0)
    ' Nayar, K. G., Mostafa H. Sharqawy, Leonardo D. Banchik, John H. Lienhard V. "Thermophysical properties of seawater:
    ' A review and new correlations that include pressure dpendence." Desalination Vol 390 (2016):1-24.
                                            
    ' Salinity as kg/kg (fraction) valid range 0 to 0.16
    ' Temperature as Kelvin (NO NEED FOR CONVERSION - EVEN THOUGH THE RELATIONSHIP IS
    ' SUPPOSEDLY IN DegC Kelvin ACTUALLY WORKS) valid range 0 < T < 180C
    ' 0 - 20C is an extrapolation using activity coefficients of pure water in sea water to derive vapor pressure.
    ' estimated error in this extrapolation is +/-0.91% while error elsewhere is 0.26%
        T = Temperature - mdlConstants.glbCelciusToKelvinOffset
        S = Salinity * 1000
        P = Pressure / 1000000#
        
        If T < 10 Or T > 120 Then
            SeaWaterSpecificEnthalpy = mdlError.ReturnError("mdlProperties.SeaWaterSpecificEnthalpy: Temperature " & T & "Celcius is outside the valid range of 10 to 120Celcius!")
        ElseIf S < 0 Or S > 120 Then
            SeaWaterSpecificEnthalpy = mdlError.ReturnError("mdlProperties.SeaWaterSpecificEnthalpy: Salinity " & S & "g/kg is outside the valid range of 0 to 160g/kg!")
        ElseIf P < 0 Or P > 12 Then
            SeaWaterSpecificEnthalpy = mdlError.ReturnError("mdlProperties.SeaWaterSpecificEnthalpy: Pressure " & P & "MPa is outside the valid range of 0 to 12MPa!")
        Else
            Dim P0 As Double
            Dim hswP0 As Double
            Dim Skgkg As Double
            Dim Hw As Double
            Dim B() As Double
            Dim A() As Double
            ReDim A(1 To 8)
            ReDim B(1 To 10)
            
            B(1) = -23482.5
            B(2) = 315183#
            B(3) = 2802690#
            B(4) = -14460600#
            B(5) = 7826.07
            B(6) = -44.1733
            B(7) = 0.21394
            B(8) = -19910.8
            B(9) = 27784.6
            B(10) = 97.2801
            
            A(1) = 996.7767
            A(2) = -3.2406
            A(3) = 0.0127
            A(4) = -0.000047723
            A(5) = -1.1748
            A(6) = 0.01169
            A(7) = -0.000026185
            A(8) = 0.000000070661
            
            If T <= 100 Then
                P0 = 0.101
            Else
                P0 = mdlProperties.SeaWaterVaporPressure(Temperature, Salinity) / 1000000#
            End If
            Skgkg = S / 1000
                
            
            Hw = 141.335 + 4202.07 * T - 0.535 * T ^ 2 + 0.004 * T ^ 3
            hswP0 = Hw - Skgkg * (B(1) + B(2) * Skgkg + B(3) * Skgkg ^ 2 + B(4) * Skgkg ^ 3 + _
                          B(5) * T + B(6) * T ^ 2 + B(7) * T ^ 3 + B(8) * Skgkg * T + B(9) * Skgkg ^ 2 * T + _
                          B(10) * Skgkg * T ^ 2)
                          
            SeaWaterSpecificEnthalpy = hswP0 + (P - P0) * (A(1) + A(2) * T + A(3) * T ^ 2 + A(4) * T ^ 3 + _
                                        S * (A(5) + A(6) * T + A(7) * T ^ 2 + A(8) * T ^ 3))
                
            
            
            
                                
        End If
    Case Else
        SeaWaterSpecificEnthalpy = mdlError.NameError("mdlProperties.SeaWaterSpecificEnthalpy", _
                                                    "RefName", _
                                                    ValidNames)
End Select

End Function

Public Function SeaWaterVaporPressure(Temperature As Double, _
                                  Salinity As Double, _
                                  Optional RefName As String = "Nayar_et_al") As Double
' In Pascals
Dim T As Double
Dim S As Double
Dim ValidNames() As String
Dim i As Long

ReDim ValidNames(0)
ValidNames(0) = "Nayar_et_al"

Select Case RefName
'You must enter every case here or the function will be inconsistent!
    Case ValidNames(0)
    ' Nayar, K. G., Mostafa H. Sharqawy, Leonardo D. Banchik, John H. Lienhard V. "Thermophysical properties of seawater:
    ' A review and new correlations that include pressure dpendence." Desalination Vol 390 (2016):1-24.
                                            
    ' Salinity as kg/kg (fraction) valid range 0 to 0.16
    ' Temperature as Kelvin (NO NEED FOR CONVERSION - EVEN THOUGH THE RELATIONSHIP IS
    ' SUPPOSEDLY IN DegC Kelvin ACTUALLY WORKS) valid range 0 < T < 180C
    ' 0 - 20C is an extrapolation using activity coefficients of pure water in sea water to derive vapor pressure.
    ' estimated error in this extrapolation is +/-0.91% while error elsewhere is 0.26%
        T = Temperature '- mdlConstants.glbCelciusToKelvinOffset
        S = Salinity * 1000
        
        If T < 273.15 Or T > 453.15 Then
            SeaWaterVaporPressure = mdlError.ReturnError("mdlProperties.SeaWaterVaporPressure: Temperature " & T & "Kelvin is outside the valid range of 273.15 to 453.15!")
        ElseIf S < 0 Or S > 160 Then
            SeaWaterVaporPressure = mdlError.ReturnError("mdlProperties.SeaWaterVaporPressure: Salinity " & S & "g/kg is outside the valid range of 0 to 160g/kg!")
        Else
            Dim Pvw As Double ' Vapor pressure of pure water
            
            Pvw = Exp(-5800 / T + 1.3915 - 0.04864 * T + 0.000041765 * T ^ 2 - 0.000000014452 * T ^ 3 + 6.546 * Log(T))
            
            SeaWaterVaporPressure = Pvw * Exp(-0.00045818 * S - 0.0000020443 * S ^ 2)
                                
        End If
    Case Else
        SeaWaterVaporPressure = mdlError.NameError("mdlProperties.SeaWaterVaporPressure", _
                                                    "RefName", _
                                                    ValidNames)
End Select

End Function

Public Function SoluteDiffusivityNaClSolution(Temperature As Double, _
                                  Salinity As Double, _
                                  Optional RefName As String = "ChiamAndSarbatly") As Double
'In m2/s
'          CHIAM, 2016 PROVIDES A RELATIONSHIP BUT I CANNOT FIND THE CONDUCTANCE DATA
'          NEEDED.
Dim T As Double
Dim S As Double
Dim ValidNames() As String
Dim i As Long

ReDim ValidNames(0)
ValidNames(0) = "ChiamAndSarbatly"

Select Case RefName
'You must enter every case here or the function will be inconsistent!
    Case ValidNames(0)
    ' This relation comes from Chel-Ken Chiam and Rosalam Sarbatly, " Study of the rectangular
    ' cross flow falt-sheet membrane module for desalination by vacuum membrane distillation"
    ' Chemical Engineering And Processing 102 (2016) 169-185.

        ' The ranges of validity for this relationship are not indicated for this relationship
        ' It is listed as the Nernst-Haskell equation from reference 42 of Chiam and Sarbatly, 2016.
        ' THE NERNST-HASKELL EQUATION IS NOWN IMPLEMENTED
        
        T = Temperature
        S = Salinity
        If T < 273.15 Or T > 453.15 Then
            err.Source = "mdlProperties.SoluteDiffusivityNaClSolution: Temperature " & T & "Celcius is outside the valid range of 273.15 to 453.15Kelvin!"
            SoluteDiffusivityNaClSolution = mdlConstants.glbINVALID_VALUE
        ElseIf S < 0 Or S > 0.16 Then
            err.Source = "mdlProperties.SoluteDiffusivityNaClSolution: Salinity " & S & "kg/kg is outside the valid range of 0 to .16kg/kg!"
            SoluteDiffusivityNaClSolution = mdlConstants.glbINVALID_VALUE
        Else
            Dim nplus As Double
            Dim nminus As Double
            Dim lambda_minus As Double
            Dim lambda_plus As Double
            Dim mu As Double
            
            nplus = 1 'Na+
            nminus = 1 'Cl-
            ' The lambda values comes from https://en.wikipedia.org/wiki/Conductivity_%28electrolytic%29
            lambda_minus = 0.007634   'Siemen * meters^2/mole
            lambda_plus = 0.005011 ' Siemen * meters^2/mole    1 Siemen = second^3*Ampere^2/(kilogram * meters^2)
            
            mu = mdlProperties.SeaWaterViscosity(T, S)
            
            ' value of 0.12e-9 obtained from table from Isidro Martinez "Mass diffusivity data.pdf"
            ' 300K listed as the operational temperature
            ' THIS NUMBER SEEMS OFF
            'SoluteDiffusivityNaClSolution = 0.00000000012
            ' 37C = 310.15K
            
            ' http://oto2.wustl.edu/cochlea/model/diffcoef.htm
            ' Handbook of Chemistry And Physics, CRC press
            'SoluteDiffusivityNaClSolution = 0.00000000199
            SoluteDiffusivityNaClSolution = 0.0000000008928 * 298.15 * (1 / nplus + 1 / nminus) * T / _
                                           ((1 / lambda_plus + 1 / lambda_minus) * 334000 * mu)
            
            
            
            
            
        End If
    Case Else
        SoluteDiffusivityNaClSolution = mdlError.NameError("mdlProperties.SoluteDiffusivityNaClSolution", _
                                                    "RefName", _
                                                    ValidNames)
End Select

End Function



Public Function SeaWaterViscosity(Temperature As Double, _
                                  Salinity As Double, _
                                  Optional RefName As String = "Isdale_et_al") As Double
'Viscosity in Pa-s
Dim t68 As Double
Dim Sp As Double
Dim ValidNames() As String
Dim i As Long

ReDim ValidNames(0)
ValidNames(0) = "Isdale_et_al"

Dim muw As Double

Select Case RefName
'You must enter every case here or the function will be inconsistent!
    Case ValidNames(0)
    ' This relation comes from Sharquwy et. al. (2010) reference 66 on page 364.
    ' Mostafa H. Sharqawy, John H. Lienhard V, Syed M. Zubair, "Thermophysical properties
    ' of seawater: a review of existing correlations and data," Desalination and Water Treatment
    ' Vol 16:pp354-380. April 2010.  www.deswater.com

    '[66] J.D. Isdale, C.M. Spence and J.S. Tudhope, Physical properties
    '   of sea water solutions: viscosity, Desalination, 10(4) (1972)
    '   319–328.
                                            
    ' Salinity as kg/kg (fraction) valid range 0 to 0.150kg/kg
    ' Temperature as Kelvin (convert back to Celcius) valid range 10 < T < 180C
    ' accurancy w/r to IAPWS-2008 formulation +/-1%
    
    'Validated against a single point by dlvilla 7/14/2016
        t68 = Temperature - mdlConstants.glbCelciusToKelvinOffset
        Sp = Salinity * 1000
        If t68 < 10 Or t68 > 180 Then
            err.Source = "mdlProperties.SeaWaterViscosity: Temperature " & t68 & "Celcius is outside the valid range of 10 to 180Celcius!"
            SeaWaterViscosity = mdlConstants.glbINVALID_VALUE
        ElseIf Sp < 0 Or Sp > 150 Then
            err.Source = "mdlProperties.SeaWaterViscosity: Salinity " & Sp & "kg/kg is outside the valid range of 0 to 150g/kg!"
            SeaWaterViscosity = mdlConstants.glbINVALID_VALUE
        Else
            Dim A As Double
            Dim B As Double
            
            A = 0.001474 + 0.000015 * t68 - 0.00000003927 * t68 ^ 2
            B = 0.00001073 - 0.000000085 * t68 + 0.000000000223 * t68 ^ 2
            
            'Be careful!! you have to consider the typo error correction at the end of the report to get this calculation of water viscosity correct
            muw = Exp(-10.7019 + (604.129 / (139.18 + t68)))
            
            SeaWaterViscosity = muw * (1 + A * Sp + B * Sp ^ 2)
            
        End If
    Case Else
        SeaWaterViscosity = mdlError.NameError("mdlProperties.SeaWaterViscosity", _
                                                    "RefName", _
                                                    ValidNames)
End Select

End Function

Public Function SeaWaterThermalConductivity(Temperature As Double, _
                                  Salinity As Double, _
                                  Optional RefName As String = "JamiesonAndTudhope") As Double
'Viscosity in Pa-s, Assumes pressure is near 1Atm
Dim t68 As Double
Dim Sp As Double

Dim muw As Double
Dim ValidNames() As String
Dim i As Long

ReDim ValidNames(0)
ValidNames(0) = "JamiesonAndTudhope"

Select Case RefName
'You must enter every case here or the function will be inconsistent!
    Case ValidNames(0)
    ' This relation comes from Sharquwy et. al. (2010) reference 58 on page 362.
    ' Mostafa H. Sharqawy, John H. Lienhard V, Syed M. Zubair, "Thermophysical properties
    ' of seawater: a review of existing correlations and data," Desalination and Water Treatment
    ' Vol 16:pp354-380. April 2010.  www.deswater.com

    '[58] D.T. Jamieson,and J.S. Tudhope, Physical properties of sea
        'water solutions – Thermal Conductivity, Desalination, 8
        '(1970) 393–401.
                                            
    ' Salinity as kg/kg (fraction) valid range 0 to 0.160kg/kg
    ' Temperature as Kelvin (convert back to Celcius) valid range 0 < T < 180C
    ' accurancy w/r to IAPWS-2008 formulation +/-3%
    
    ' Correct Implmentation checked by Daniel Villa 7/14/2016
        t68 = Temperature - mdlConstants.glbCelciusToKelvinOffset
        Sp = Salinity * 1000
        If t68 < 0 Or t68 > 180 Then
            err.Source = "mdlProperties.SeaWaterThermalConductivity: Temperature " & t68 & "Celcius is outside the valid range of 0 to 180Celcius!"
            SeaWaterThermalConductivity = mdlConstants.glbINVALID_VALUE
        ElseIf Sp < 0 Or Sp > 160 Then
            err.Source = "mdlProperties.SeaWaterThermalConductivity: Salinity " & Sp & "g/kg is outside the valid range of 0 to 160g/kg!"
            SeaWaterThermalConductivity = mdlConstants.glbINVALID_VALUE
        Else
        'Must divide by 1000 to get W/m-K! because the original relation outputs milliWatt/(m*K)
            SeaWaterThermalConductivity = 10 ^ (Log(240 + 0.0002 * Sp) / Log(10) + _
            0.434 * (2.3 - (343.5 + 0.037 * Sp) / (t68 + 273.15)) * (1 - (t68 + 273.15) / (647 + 0.03 * Sp)) ^ 0.333) / 1000
        End If
    Case Else
        SeaWaterThermalConductivity = mdlError.NameError("mdlProperties.SeaWaterThermalConductivity", _
                                                    "RefName", _
                                                    ValidNames)
End Select

End Function

'STOPPED DEVELOPMENT OF THIS FUNCTION BECAUSE A MORE RECENT SOURCE WAS FOUND
'Public Function Air_Water_Vapor_Thermal_Conductivity(Temperature As Double) As Double
'' This function uses the relationship for 1950-LindsayAndBromley-ThermalConductivityOfGasMixtures-IndustrialAndEngineeringChemistry.pdf
'' it only requires two elements
'' k_mix = k1 / (1 + A12 * (x2/x1)) + k2 / (1 + A21 * (x1/x2))
'' where
'' A12 = 0.25 * ( 1 + ((mu1/mu2)*(M2 / M1)^a * (1 + (S1 / T) / ( 1 + (S2 / T))) ^ 0.5)^2 * ((M1 + M2)/(2*M2))^b * (1 + S12/T) / (1 + S1/T)
''
''  mu = gas viscosity, S = Sutherland constant = 1.5 T_boil. S12 = 0.733 * sqrt(S1*S2) (used this instead of sqrt(S1*S2) because water vapor is strongly polar).
'
'  Dim mu1 'air viscosity
'  Dim mu2 'water viscosity
'
'  mu1 = DryAirViscosity(Temperature) ' Pressure dependence is insignificant!
'  mu2 = SeaWaterViscosity(Temperature, 0) ' Do we need a water vapor viscosity?
'
'End Function

Public Function WaterVaporMoleFractionSaturatedAir(Temperature As Double, Pressure As Double) As Double
' This function has been tested based on a psychometric chart's evaluation.
Dim f As Double
Dim Psv As Double 'Saturated pressure of water
Dim Pv As Double 'vapor pressure of water at Temperature

Psv = mdlProperties.SaturatedPressurePureWater(Temperature)

' This function does the necessary checks for errors so that none are needed at this level.
' f represents departure from ideal gas behavior.
f = f_enhancementfactor(Pressure, Temperature, Psv)

WaterVaporMoleFractionSaturatedAir = f * Psv / Pressure

End Function

Public Function SteamViscosity(Temperature As Double) As Double 'kg/(m*s)
' This is from 2008-Tsilingiris equation 41 but the equation was clearly off by a factor of 10!.  This was double confirmed
' by comparing the data to Lindberg-2006, page A-100 and data from the Engineering toolbox website:

' www.engineeringtoolbox.com/steam-viscosity-d_770.html

' The comparisons are available in an excel spreadsheet in the repository that this code is contained in:

'\MaterialProp\NIST_Matlab\ViscosityOfSteam_Comparison_04242017.xlsx

Dim T As Double

T = Temperature - mdlConstants.glbCelciusToKelvinOffset

If T < 0 Or T > 120 Then
   SteamViscosity = mdlError.ReturnError("mdlProperties.SteamViscosity: The requested temperature (" _
                    & CStr(T) & " degC) is outside the valid range of 0 to 120degC!", , True)
Else  ' The factor of ten difference in the coefficients is a correction. These values are 1/10th of the Tsilingiris coefficients!
   SteamViscosity = 0.000001 * (8.058131868 + 0.04000549451 * T)
End If

End Function

Public Function AirWaterSaturatedMixtureThermalConductivity(Temperature As Double, Pressure As Double) As Double
' This is from equation 28 of 2008, Tsilingiris.

Dim ka As Double
Dim kv As Double
Dim xv As Double
Dim xa As Double
Dim mu_a As Double
Dim mu_v As Double
Dim Psi_av As Double
Dim Psi_va As Double
Dim Mratio As Double
Dim mu_ratio As Double

ka = DryAirThermalConductivityAt1Atm(Temperature)  ' mdlCharlieProp.AirThermalCond(Temperature)
kv = WaterVaporThermalConductivity(Temperature)
xv = WaterVaporMoleFractionSaturatedAir(Temperature, Pressure)
xa = 1 - xv
mu_a = mdlProperties.DryAirViscosity(Temperature, Pressure)
mu_v = mdlProperties.SteamViscosity(Temperature) 'This is not significantly sensitive to Pressure

Mratio = mdlConstants.glbWaterMolecularWeight / mdlConstants.glbAirMolecularWeight
mu_ratio = mu_v / mu_a

Psi_av = 2 ^ (-1.5) * (1 + 1 / Mratio) ^ -0.5 * (1 + (mu_ratio) ^ 0.5 * (Mratio) ^ 0.25) ^ 2
Psi_va = 2 ^ (-1.5) * (1 + Mratio) ^ -0.5 * (1 + (1 / mu_ratio) ^ 0.5 * (1 / Mratio) ^ 0.25) ^ 2

AirWaterSaturatedMixtureThermalConductivity = xa * ka / (xa + xv * Psi_av) + xv * kv / (xv + xa * Psi_va)

End Function

Public Function WaterVaporThermalConductivity(Temperature As Double) As Double 'W/(m*K)
' This is from 2008-Tsilingiris-Thermophysical And Transport Properties of humid air at temperature range between 0 and 100C" Energy conversion and management. pp 1098-1110
' page 1103.
' !@#$ This function needs testing! - shown to be in reasonable agreement with internet table value in
'  http://www.engineeringtoolbox.com/thermal-conductivity-d_429.html
'
Const KV0 = 17.61758242
Const KV1 = 0.05558941059
Const KV2 = 0.0001663336663

Dim T As Double

T = Temperature - mdlConstants.glbCelciusToKelvinOffset

If T < 0 Or T > 120 Then
   WaterVaporThermalConductivity = mdlError.ReturnError("mdlProperties.WaterVaporThermalConductivity: Requested temperature (" & CStr(T) & _
                                   " deg C) is outside the valid range of 0 to 120 deg C!", , True)
Else
   WaterVaporThermalConductivity = (KV0 + KV1 * T + KV2 * T ^ 2) / 1000
End If

End Function

Private Function f_enhancementfactor(Pressure As Double, Temperature As Double, Psv As Double) As Double
' From 2008-Tsilingiris page 1100. This factor is very near to 1.0 and the expressions that use can neglect it and only produce a 0.5% error.
Dim xi1 As Double
Dim xi2 As Double
'Psv is the saturated pressure of water

Const A0 = 0.000353624
Const A1 = 0.0000293228
Const A2 = 0.000000261474
Const A3 = 0.00000000857538

Const B0 = -10.7588
Const B1 = 0.0632529
Const B2 = -0.000253591
Const B3 = 0.000000633784

Dim T As Double

T = Temperature - mdlConstants.glbCelciusToKelvinOffset

If T < 0 Or T > 100 Then
   f_enhancementfactor = mdlError.ReturnError("mdlProperties.f_enhancementfactor: The temperature requested (" & CStr(T) & " degC) is outside the valid range of 0 to 100C")
ElseIf Pressure < 0 Then
   f_enhancementfactor = mdlError.ReturnError("mdlProperties.f_enhancementfactor: The pressure requested (" & CStr(Pressure) & " Pa) is negative!")
Else

   xi1 = A0 + A1 * T + A2 * T ^ 2 + A3 * T ^ 3
   xi2 = Exp(B0 + B1 * T + B2 * T ^ 2 + B3 * T ^ 3)
   f_enhancementfactor = Exp(xi1 * (1 - Psv / Pressure) + xi2 * (Psv / Pressure - 1))
   
End If

End Function

Public Function DryAirThermalDiffusivityAt1Atm(Temperature As Double) As Double  'm2/s

' NIST REFPROP database.
'LITERATURE REFERENCE
'Lemmon, E.W., Jacobsen, R.T, Penoncello, S.G., and Friend, D.G.,
'"Thermodynamic Properties of Air and Mixtures of Nitrogen, Argon, and
'Oxygen from 60 to 2000 K at Pressures to 2000 MPa,"
'J. Phys. Chem. Ref. Data, 29(3):331-385, 2000.
'
'In the range from the solidification point to 873 K at pressures to 70
'MPa, the estimated uncertainty of density values calculated with the
'equation of state is 0.1%.  The estimated uncertainty of calculated
'speed of sound values is 0.2% and that for calculated heat capacities is
'1%.  At temperatures above 873 K and 70 MPa, the estimated uncertainty
'of calculated density values is 0.5% increasing to 1.0% at 2000 K and
'2000 MPa.
   
   'MATLAB FIT cftool 2016a
'Linear model Poly3:
'     f(x) = p1 * x ^ 3 + p2 * x ^ 2 + p3 * x + p4
'Coefficients (with 95% confidence bounds):
'       p1 =  -1.325e-09  (-1.327e-09, -1.322e-09)
'       p2 =   2.805e-06  (2.802e-06, 2.807e-06)
'       p3 =   3.812e-05  (3.737e-05, 3.887e-05)
'       p4 =   -0.005343  (-0.00542, -0.005266)
'
'Goodness of fit:
'  SSE: 8.337e-09
'  R-square: 1
'  Adjusted R-square: 1
'  RMSE: 5.81e-06
   '
   ' Valid Range is 200K to 450K
   
   If Temperature < 200 Or Temperature > 450 Then
        DryAirThermalDiffusivityAt1Atm = mdlError.ReturnError("mdlProperties.DryAirThermalDiffusivityAt1Atm: The requested temperature (" & CStr(Temperature) & _
                 ") is outside the valid range of 200K to 450K!", , True)
   Else ' We are within the valid bounds
        DryAirThermalDiffusivityAt1Atm = (-0.000000001325 * Temperature ^ 3 + 0.000002805 * Temperature ^ 2 + 0.00003812 * Temperature - 0.005343) * _
                                         mdlConstants.glbcm2ToMeter2
   End If

End Function

Public Function DryAirKinematicViscosityAt1Atm(Temperature As Double) As Double 'm2/s

' NIST REFPROP database.
'LITERATURE REFERENCE
'Lemmon, E.W., Jacobsen, R.T, Penoncello, S.G., and Friend, D.G.,
'"Thermodynamic Properties of Air and Mixtures of Nitrogen, Argon, and
'Oxygen from 60 to 2000 K at Pressures to 2000 MPa,"
'J. Phys. Chem. Ref. Data, 29(3):331-385, 2000.
'
'In the range from the solidification point to 873 K at pressures to 70
'MPa, the estimated uncertainty of density values calculated with the
'equation of state is 0.1%.  The estimated uncertainty of calculated
'speed of sound values is 0.2% and that for calculated heat capacities is
'1%.  At temperatures above 873 K and 70 MPa, the estimated uncertainty
'of calculated density values is 0.5% increasing to 1.0% at 2000 K and
'2000 MPa.
   
   'MATLAB FIT cftool 2016a
'Linear model Poly3:
'     f(x) = p1 * x ^ 3 + p2 * x ^ 2 + p3 * x + p4
'Coefficients (with 95% confidence bounds):
'       p1 =  -5.447e-10  (-5.499e-10, -5.395e-10)
'       p2 =   1.574e-06  (1.569e-06, 1.579e-06)
'       p3 =   0.0001381  (0.0001365, 0.0001397)
'       p4 =     -0.0109  (-0.01106, -0.01073)
'
'Goodness of fit:
'  SSE: 3.836e-08
'  R-square: 1
'  Adjusted R-square: 1
'  RMSE: 1.246e-05
   '
   ' Valid Range is 200K to 450K
   
   If Temperature < 200 Or Temperature > 450 Then
        DryAirKinematicViscosityAt1Atm = mdlError.ReturnError("mdlProperties.DryAirKinematicViscosityAt1Atm: The requested temperature (" & CStr(Temperature) & _
                 ") is outside the valid range of 200K to 450K!", , True)
   Else ' We are within the valid bounds
        DryAirKinematicViscosityAt1Atm = (-0.0000000005447 * Temperature ^ 3 + 0.000001574 * Temperature ^ 2 + 0.0001381 * Temperature - 0.0109) * _
                                         mdlConstants.glbcm2ToMeter2
   End If

End Function

Public Function DryAirDensityAt1Atm(Temperature As Double) As Double 'kg/m3

' NIST REFPROP database.
'LITERATURE REFERENCE
'Lemmon, E.W., Jacobsen, R.T, Penoncello, S.G., and Friend, D.G.,
'"Thermodynamic Properties of Air and Mixtures of Nitrogen, Argon, and
'Oxygen from 60 to 2000 K at Pressures to 2000 MPa,"
'J. Phys. Chem. Ref. Data, 29(3):331-385, 2000.
'
'In the range from the solidification point to 873 K at pressures to 70
'MPa, the estimated uncertainty of density values calculated with the
'equation of state is 0.1%.  The estimated uncertainty of calculated
'speed of sound values is 0.2% and that for calculated heat capacities is
'1%.  At temperatures above 873 K and 70 MPa, the estimated uncertainty
'of calculated density values is 0.5% increasing to 1.0% at 2000 K and
'2000 MPa.
   
   'MATLAB FIT cftool 2016a
'Linear model Poly6:
'     f(x) = p1*x^6 + p2*x^5 + p3*x^4 + p4*x^3 + p5*x^2 +
'                    p6*x + p7
'Coefficients (with 95% confidence bounds):
      Const p1 = 1.354E-15    '(1.301e-15, 1.406e-15)
      Const p2 = -0.000000000003047 '(-3.149e-12, -2.945e-12)
      Const p3 = 0.000000002902 '(2.82e-09, 2.984e-09)
      Const p4 = -0.000001515 '(-1.55e-06, -1.48e-06)
      Const p5 = 0.0004678    '(0.0004596, 0.0004759)
      Const p6 = -0.08529     '(-0.08629, -0.08429)
      Const p7 = 8.483        '(8.432, 8.533)
'
'Goodness of fit:
'  SSE: 2.425e-07
'  R-square: 1
'  Adjusted R-square: 1
'  RMSE: 3.153e-05


   '
   ' Valid Range is 200K to 450K
   
   If Temperature < 200 Or Temperature > 450 Then
        DryAirDensityAt1Atm = mdlError.ReturnError("mdlProperties.DryAirDensityAt1Atm: The requested temperature (" & CStr(Temperature) & _
                 ") is outside the valid range of 200K to 450K!", , True)
   Else ' We are within the valid bounds
        DryAirDensityAt1Atm = p1 * Temperature ^ 6 + p2 * Temperature ^ 5 + p3 * Temperature ^ 4 + p4 * Temperature ^ 3 + p5 * Temperature ^ 2 + _
                    p6 * Temperature + p7
   End If

End Function

Public Function DryAirPrandtlAt1Atm(Temperature As Double) As Double 'kg/m3

' NIST REFPROP database.
'LITERATURE REFERENCE
'Lemmon, E.W., Jacobsen, R.T, Penoncello, S.G., and Friend, D.G.,
'"Thermodynamic Properties of Air and Mixtures of Nitrogen, Argon, and
'Oxygen from 60 to 2000 K at Pressures to 2000 MPa,"
'J. Phys. Chem. Ref. Data, 29(3):331-385, 2000.
'
'In the range from the solidification point to 873 K at pressures to 70
'MPa, the estimated uncertainty of density values calculated with the
'equation of state is 0.1%.  The estimated uncertainty of calculated
'speed of sound values is 0.2% and that for calculated heat capacities is
'1%.  At temperatures above 873 K and 70 MPa, the estimated uncertainty
'of calculated density values is 0.5% increasing to 1.0% at 2000 K and
'2000 MPa.
   
   'MATLAB FIT cftool 2016a
'Linear model Poly6:
'     f(x) = p1*x^6 + p2*x^5 + p3*x^4 + p4*x^3 + p5*x^2 +
'                    p6*x + p7
'Coefficients (with 95% confidence bounds):
       Const p1 = 4.464E-17    '(4.017e-17, 4.91e-17)
       Const p2 = -9.756E-14   '(-1.063e-13, -8.885e-14)
       Const p3 = 0.00000000008822 '(8.123e-11, 9.521e-11)
       Const p4 = -0.0000000426 '(-4.555e-08, -3.965e-08)
       Const p5 = 0.00001211   '(1.142e-05, 1.28e-05)
       Const p6 = -0.002118    '(-0.002203, -0.002032)
       Const p7 = 0.8927       '(0.8884, 0.897)
'
'Goodness of fit:
'  SSE: 1.767e-09
'  R-square: 1
'  Adjusted R-square: 1
'  RMSE: 2.691e-06
   '
   ' Valid Range is 200K to 450K
   
   If Temperature < 200 Or Temperature > 450 Then
        DryAirPrandtlAt1Atm = mdlError.ReturnError("mdlProperties.DryAirPrandtlAt1Atm: The requested temperature (" & CStr(Temperature) & _
                 ") is outside the valid range of 200K to 450K!", , True)
   Else ' We are within the valid bounds
        DryAirPrandtlAt1Atm = p1 * Temperature ^ 6 + p2 * Temperature ^ 5 + p3 * Temperature ^ 4 + p4 * Temperature ^ 3 + p5 * Temperature ^ 2 + _
                    p6 * Temperature + p7
   End If

End Function

Public Function DryAirThermalConductivityAt1Atm(Temperature As Double) As Double 'W/(m*K)

' NIST REFPROP database.
'LITERATURE REFERENCE
'Lemmon, E.W., Jacobsen, R.T, Penoncello, S.G., and Friend, D.G.,
'"Thermodynamic Properties of Air and Mixtures of Nitrogen, Argon, and
'Oxygen from 60 to 2000 K at Pressures to 2000 MPa,"
'J. Phys. Chem. Ref. Data, 29(3):331-385, 2000.
'
'In the range from the solidification point to 873 K at pressures to 70
'MPa, the estimated uncertainty of density values calculated with the
'equation of state is 0.1%.  The estimated uncertainty of calculated
'speed of sound values is 0.2% and that for calculated heat capacities is
'1%.  At temperatures above 873 K and 70 MPa, the estimated uncertainty
'of calculated density values is 0.5% increasing to 1.0% at 2000 K and
'2000 MPa.
   
   'Excel fit R2 = 1.000 errors from data < 0.01% over range of 200K to 450K
'
'Goodness of fit:
'  SSE: 1.767e-09
'  R-square: 1
'  Adjusted R-square: 1
'  RMSE: 2.691e-06
   '
   ' Valid Range is 200K to 450K
   
       Const p1 = 0.00000000004349
       Const p2 = -0.00000007976
       Const p3 = 0.0001104
       Const p4 = -0.0007323

   
   If Temperature < 200 Or Temperature > 450 Then
        DryAirThermalConductivityAt1Atm = mdlError.ReturnError("mdlProperties.DryAirThermalConductivityAt1Atm: The requested temperature (" & CStr(Temperature) & _
                 ") is outside the valid range of 200K to 450K!", , True)
   Else ' We are within the valid bounds
        DryAirThermalConductivityAt1Atm = p1 * Temperature ^ 3 + p2 * Temperature ^ 2 + p3 * Temperature + p4
   End If

End Function

Public Function DryAirViscosity(Temperature As Double, Optional Pressure As Double = 101325) As Double

   ' NIST REFPROP database.
'LITERATURE REFERENCE
'Lemmon, E.W., Jacobsen, R.T, Penoncello, S.G., and Friend, D.G.,
'"Thermodynamic Properties of Air and Mixtures of Nitrogen, Argon, and
'Oxygen from 60 to 2000 K at Pressures to 2000 MPa,"
'J. Phys. Chem. Ref. Data, 29(3):331-385, 2000.
'
'In the range from the solidification point to 873 K at pressures to 70
'MPa, the estimated uncertainty of density values calculated with the
'equation of state is 0.1%.  The estimated uncertainty of calculated
'speed of sound values is 0.2% and that for calculated heat capacities is
'1%.  At temperatures above 873 K and 70 MPa, the estimated uncertainty
'of calculated density values is 0.5% increasing to 1.0% at 2000 K and
'2000 MPa.
   
   'MATLAB FIT cftool 2016a
'Linear model Poly11:
'     f(X, Y) = p00 + p10 * X + p01 * Y
'Coefficients (with 95% confidence bounds):
'       p00 =   4.812e-06  (4.791e-06, 4.834e-06)
'       p10 =   4.573e-08  (4.567e-08, 4.58e-08)
'       p01 =   1.355e-13  (1.256e-13, 1.454e-13)
'
'Goodness of fit:
'  SSE: 1.017e-12
'  R-square: 0.9996
'  Adjusted R-square: 0.9996
'  RMSE: 3.415e-08
   '
   ' This is a nearly exact fit (not sure what the NIST REFPROP Errors are) to the NIST REFPROP constitutive relationship
   ' max error is 0.4156% from the NIST fit. Not sure what the NIST Fit errors are. Pressure dependence is very week.
   
   If Temperature < 280 Or Temperature > 400 Then
        DryAirViscosity = mdlError.ReturnError("mdlProperties.DryAirViscosity: The requested temperature (" & CStr(Temperature) & _
                 ") is outside the valid range of 280K to 400K!", , True)
   ElseIf Pressure < 0.1 Or Pressure > 1000000# Then
        DryAirViscosity = mdlError.ReturnError("mdlProperties.DryAirViscosity: The requested pressure (" & CStr(Pressure) & _
                 ") is outside the valid range of 0.1Pa to 1,000,000Pa!", , True)
   Else ' We are within the valid bounds
        DryAirViscosity = 0.000004812152422 + 4.57335956E-08 * Temperature + 1.354851553E-13 * Pressure
   End If

End Function

Public Function SeaWaterSpecificHeat(Temperature As Double, _
                                  Salinity As Double, _
                                  Optional RefName As String = "Jamieson_et_al", Optional ThrowError As Boolean = True) As Double 'J/(kg*K)
'Viscosity in Pa-s
Dim t68 As Double
Dim Sp As Double

Dim muw As Double
Dim ValidNames() As String
Dim i As Long

ReDim ValidNames(0)
ValidNames(0) = "Jamieson_et_al"

Select Case RefName
'You must enter every case here or the function will be inconsistent!
    Case ValidNames(0)
    ' This relation comes from Sharquwy et. al. (2010) reference 46 on page 360.
    ' Mostafa H. Sharqawy, John H. Lienhard V, Syed M. Zubair, "Thermophysical properties
    ' of seawater: a review of existing correlations and data," Desalination and Water Treatment
    ' Vol 16:pp354-380. April 2010.  www.deswater.com

    '[46] D.T. Jamieson, J.S. Tudhope, R. Morris and G. Cartwright,
        'Physical properties of sea water solutions: heat capacity,
        'Desalination, 7(1) (1969) 23–30.
                                            
    ' Salinity as kg/kg (fraction) valid range 0 to 0.180kg/kg
    ' Temperature as Kelvin (convert back to Celcius) valid range 273.15 < T < 453.15K
    ' accurancy w/r to IAPWS-2008 formulation +/-.28%
    
    ' Correct Implmentation checked by Daniel Villa 7/14/2016
        t68 = Temperature
        Sp = Salinity * 1000
        If t68 < 273.15 Or t68 > 453.15 Then
            err.Source = "mdlProperties.SeaWaterSpecificHeat: Temperature " & t68 & "Kelvin is outside the valid range of 273.15 to 453.15Kelvin!"
            SeaWaterSpecificHeat = mdlConstants.glbINVALID_VALUE
        ElseIf Sp < 0 Or Sp > 180 Then
            err.Source = "mdlProperties.SeaWaterSpecificHeat: Salinity " & Sp & "g/kg is outside the valid range of 0 to 180g/kg!"
            SeaWaterSpecificHeat = mdlConstants.glbINVALID_VALUE
        Else
            Dim A As Double
            Dim B As Double
            Dim c As Double
            Dim d As Double
            A = 5.328 - 0.0976 * Sp + 0.000404 * Sp ^ 2
            B = -0.006913 + 0.0007351 * Sp - 0.00000315 * Sp ^ 2
            c = 0.0000096 - 0.000001927 * Sp + 0.00000000823 * Sp ^ 2
            d = 0.0000000025 + 0.000000001666 * Sp - 0.000000000007125 * Sp ^ 2
            
            ' Originally this was in KJ/(kg *K) but we need J/(kg *K)
            SeaWaterSpecificHeat = 1000 * (A + B * t68 + c * t68 ^ 2 + d * t68 ^ 3)
        End If
    Case Else
        SeaWaterSpecificHeat = mdlError.NameError("mdlProperties.SeaWaterSpecificHeat", _
                                                    "RefName", _
                                                    ValidNames)
End Select

End Function
