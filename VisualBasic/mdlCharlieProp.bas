Attribute VB_Name = "mdlCharlieProp"
'
'    Primary Author Charles Morrow
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



'Option Explicit
'
'' WARNING! MAny of these function do not have valid range error brackets! They are purely emperical fits
'' that will return invalid values outside their ranges of validities.  Some do have valid range checks.
'
'Function PvWaterNIST(Tk As Double) As Double 'in Pascal
'    'Returns vapor pressure of water (Pascal) as a function of water temperature
'    '(Kelvin).  Based on Nist Antoine relationship
'    'P= 10^(A-(B//(T+C)))
'    'http://webbook.nist.gov/cgi/cbook.cgi?ID=C7732185&Units=SI&Mask=4#Thermo-Phase
'    '    379    -   573 3.55959 643.748     -198.043
'    '    273    -   303 5.40221 1838.675    -31.737
'    '    304    -   333 5.20389 1733.926    -39.485
'    '    334    -   363 5.0768  1659.793    -45.854
'    '    344    -   373 5.08354 1663.125    -45.622
'    '    293    -   343 6.20963 2354.731     7.559
'    '    255.9  -   373 4.6543  1435.264    -64.848
'    Dim dA As Double
'    Dim dB As Double
'    Dim dC As Double
'
'    If Tk < 273 Or Tk >= 373.15 Then
'        PvWaterNIST = mdlError.ReturnError("mdlCharlieProp.PvWaterNIST: Temperature = " & CStr(Tk) & " is outside of the valid range of 273 to 373.15.", , True)
'    ElseIf Tk >= 273 And Tk < 303 Then
'        dA = 5.40221
'        dB = 1838.675
'        dC = -31.737
'    ElseIf Tk >= 303 And Tk < 333 Then
'        dA = 5.20389
'        dB = 1733.926
'        dC = -39.485
'    ElseIf Tk >= 333 And Tk < 353 Then
'        dA = 5.0768
'        dB = 1659.793
'        dC = -45.854
'    Else ' If Tk >= 353 And Tk < 373.15 Then
'        dA = 5.08354
'        dB = 1663.125
'        dC = -45.622
'    End If
'
'    PvWaterNIST = (100000) * 10 ^ (dA - (dB / (Tk + dC)))
'
'End Function
'
'Function WaterDen(Tk As Double) As Double 'kg/m³
'    'verified 6/5/2013
'
'    Const c0 As Double = 770.2259596
'    Const C1 As Double = 1.781189646
'    Const C2 As Double = -0.003424621
'
'    WaterDen = c0 + Tk * (C1 + Tk * C2)
'
'
'End Function
'
'Function WaterThermCond(Tk As Double) As Double 'W/m-K
'    'verified 6/5/2013
'    Const c0 As Double = -754.6891919
'    Const C1 As Double = 7.46199899
'    Const C2 As Double = -0.009702525
'
'    WaterThermCond = c0 + Tk * (C1 + Tk * C2)
'
'End Function
'
'Function WaterViscosity(Tk As Double) As Double 'Pa - s
'    'verified 6/5/2013
'
'    Const c0 As Double = 93274.77988
'    Const C1 As Double = -786.1876882
'    Const C2 As Double = 2.229287121
'    Const C3 As Double = -0.002118194
'
'    WaterViscosity = c0 + Tk * (C1 + Tk * (C2 + C3 * Tk))
'
'End Function
'
'Function WaterThermDiff(Tk As Double) As Double 'm²/s
'
'    Const c0 As Double = -0.001782749
'    Const C1 As Double = 0.0000171719
'    Const C2 As Double = -0.000000021149
'
'    WaterThermDiff = c0 + Tk * (C1 + Tk * C2)
'
'End Function
'
'Function WaterPrandtl(Tk As Double) As Double
'    'verified 6/5/2013
'
'    Const c0 As Double = 3306.723488
'    Const C1 As Double = -37.9870686
'    Const C2 As Double = 0.164421383
'    Const C3 As Double = -0.000317249
'    Const c4 As Double = 0.000000230005
'
'    WaterPrandtl = c0 + Tk * (C1 + Tk * (C2 + Tk * (C3 + Tk * c4)))
'
'End Function
'
'Function SteamThermalCond(Tk As Double) As Double 'W/m-K
'    'verified 6/5/2013
'    Const c0 As Double = 0.000000294346
'    Const C1 As Double = -0.000110274
'    Const C2 As Double = 0.025256635
'    Const Sey As Double = 0.00000954254
'
'    SteamThermalCond = c0 + Tk * (C1 + Tk * C2)
'
'End Function
'
'Function WaterLatentHeat(Tk As Double) As Double 'J/kg
'    'Verified 6/5/2013
'    'Input Tk (K)
'    Dim c0 As Double
'    Dim C1 As Double
'    Dim C2 As Double
'    Dim C3 As Double
'    Dim c4 As Double
'    Dim R2 As Double
''    Dim Sey As Double
''    Dim dX As Double
''    Dim Out As Variant
'
'    If Tk >= 280 And Tk <= 400 Then
'    '-0.013305406    11.60926996 -5747.812833    3476011.365
'         c4 = 0
'         C3 = -0.013305406
'         C2 = 11.60926996
'         C1 = -5747.812833
'         c0 = 3476011.365
''         R2 = 0.999999868
''         Sey = 35.68889627
'    ElseIf Tk > 400 And Tk <= 550 Then
'    '-0.031565779    34.52660533 -15379.86136    4830994.114
'         c4 = 0
'         C3 = -0.031565779
'         C2 = 34.52660533
'         C1 = -15379.86136
'         c0 = 4830994.114
''         R2 = 0.9999998
''         Sey = 192.1058841
'    ElseIf Tk > 550 And Tk <= 640 Then
'    '-0.011651787    26.71797499 -22996.0012 8797191.216 -1259590943
'         c4 = -0.011651787
'         C3 = 26.71797499
'         C2 = -22996.0012
'         C1 = 8797191.216
'         c0 = -1259590943
''         R2 = 0.999894906
''         Sey = 3595.041474
'    Else
'        WaterLatentHeat = mdlError.ReturnError("mdlCharlieProp.WaterLatentHeat: Temperature = " & CStr(Tk) & " is outside the valid range of 280 to 640 Kelvin!", , True)
'    End If
'
'    WaterLatentHeat = c0 + Tk * (C1 + Tk * (C2 + Tk * (C3 + Tk * c4)))
'
'End Function
'
'Function WaterdLdT(Tk As Double) As Double
'    'Verified 6/5/2013
'    'Input Tk (K)
'    Dim c0 As Double
'    Dim C1 As Double
'    Dim C2 As Double
'    Dim C3 As Double
'    Dim c4 As Double
'
'    If Tk >= 280 And Tk <= 400 Then
'    '-0.013305406    11.60926996 -5747.812833    3476011.365
'         c4 = 0
'         C3 = -0.013305406
'         C2 = 11.60926996
'         C1 = -5747.812833
'         c0 = 3476011.365
''         R2 = 0.999999868
''         Sey = 35.68889627
'    ElseIf Tk > 400 And Tk <= 550 Then
'    '-0.031565779    34.52660533 -15379.86136    4830994.114
'         c4 = 0
'         C3 = -0.031565779
'         C2 = 34.52660533
'         C1 = -15379.86136
'         c0 = 4830994.114
''         R2 = 0.9999998
''         Sey = 192.1058841
'    ElseIf Tk > 550 And Tk <= 640 Then
'    '-0.011651787    26.71797499 -22996.0012 8797191.216 -1259590943
'         c4 = -0.011651787
'         C3 = 26.71797499
'         C2 = -22996.0012
'         C1 = 8797191.216
'         c0 = -1259590943
''         R2 = 0.999894906
''         Sey = 3595.041474
'    Else
'        WaterdLdT = mdlError.ReturnError("mdlCharlieProp.WaterdLdT: Temperature = " & CStr(Tk) & " is outside the valid range of 280 to 640 Kelvin!", , True)
'    End If
'
'    WaterdLdT = C1 + Tk * (2 * C2 + Tk * (3 * C3 + Tk * (4 * c4)))
'
'End Function
'
'Function WaterEnthalpy(Tk As Double) As Double 'J/kg
'    'Input Tk (K)
'    'Verified 6/5/2013
''    3.36801E-05 -0.050150239    28.38379477 -2997.579618    -457662.9716
'
'    Const c0 As Double = -457662.9716
'    Const C1 As Double = -2997.579618
'    Const C2 As Double = 28.38379477
'    Const C3 As Double = -0.050150239
'    Const c4 As Double = 0.0000336801
'
'    WaterEnthalpy = c0 + Tk * (C1 + Tk * (C2 + Tk * (C3 + Tk * c4)))
'
'
'End Function
'
'Function WaterSpHeat(Tk As Double) As Double 'J/kg
'    'Input Tk (K)
'    'Verified 9/5/2013
'    ' Slimmed down for faster evaluation by dlvilla 4/12/2017
'    Dim H As Double
'
'    H = WaterEnthalpy(Tk)
'    WaterSpHeat = H / (Tk - 273.15)
'
'End Function
'
'Function AirThermalCond(Tk As Double) As Double 'W/m/K
'    'Input Tk (K)
'    '-2.40763E-08    8.54928E-05 0.002699045
'
'
'    Const c0 As Double = 0.002699045
'    Const C1 As Double = 0.0000854928
'    Const C2 As Double = -0.0000000240763
'    'Const Sey As Double = 0.0000797253
'
'    AirThermalCond = c0 + Tk * (C1 + Tk * C2)
'
'End Function
