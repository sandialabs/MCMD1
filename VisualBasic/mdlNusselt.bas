Attribute VB_Name = "mdlNusselt"
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
Option Base 0

Public Function FlatSheetLaminarFlow(Reynolds As Double, _
                                                  Prandtl As Double, _
                                                  HydraulicDiameter As Double, _
                                                  Length As Double, _
                                                  RelationShipToUse As String, _
                                                  Optional PrandtlInterface As Double) As Double



Dim Nu As Double
Dim ValidNames() As String
ReDim ValidNames(2)
ValidNames(0) = "Gryta_et_al_Table1No11"
ValidNames(1) = "Gryta_et_al_Table1No8"
ValidNames(2) = "Gryta_et_al_Table1No6"

If Reynolds > glbMaxReynolds Then
    err.Source = "mdlNusselt.FlatSheetLaminarFlow: Reynolds number (" & _
              Reynolds & ") must be less than " & glbMaxReynolds & " for laminar flow relationships!"
    FlatSheetLaminarFlow = mdlConstants.glbINVALID_VALUE
Else
    Select Case RelationShipToUse
        Case ValidNames(0)
            'L. Martínez, F.J. Florido-Díaz, Theoretical and experimental studies on
            'desalination using membrane distillation, Desalination 139 (2001) 373–379.
            'http://refhub.elsevier.com/S0255-2701(16)30014-9/sbref0160
            '
            ' Length of 55mm is listed as the length used for the development of this
            ' relationship. The accuracy for smaller divisions in this model
            ' is unknown.  As a result the length is fixed to 0.055 to avoid making heat trasfer rate
            ' length dependent.
            If Reynolds > 2100 Then
                FlatSheetLaminarFlow = mdlError.ReturnError("mdlNusselt.FlatSheetLaminarFlow: The Relationship in  1997-Gryta Table 1 No 11. must have " & _
                                        " a Reynold's number less than 2100 to be valid!", , True)
            Else
                Nu = 1.86 * (Reynolds * Prandtl * HydraulicDiameter / 0.055) ^ (1 / 3)
            End If
        Case ValidNames(1)
            ' This relationship was found in
            ' Hitsov, I, T Maere, K De Sitter, C. Dotremont, I. Nopens. "Modelling approaches in
            ' membrane distillation: A critical review" Separation and Purification Technology Vol 142
            ' (2015): 48-64. Whose 6th reference is used for a Nu equation for laminar flow.
            '[6] M. Gryta, M. Tomaszewska, M.A. Morawski, Membrane distillation with
            '    laminar flow, Sep. Purif. Technol. 11 (1997) 93–101.
            ' http://refhub.elsevier.com/S1383-5866(14)00771-0/h0030
            If IsMissing(PrandtlInterface) Then
                FlatSheetLaminarFlow = mdlError.ReturnError("mdlNusselt.FlatSheetLaminarFlow: When Using " & _
                                                            ValidNames(1) & " the optional argument ""PrandtlInterface""" & _
                                                            " must be included.", , True)
            Else
                If Reynolds <= 1000 Then
                   FlatSheetLaminarFlow = mdlError.ReturnError("mdlNusselt.FlatSheetLaminarFlow: The Relationship in Hitsov (presented by 1997-Gryta) must have " & _
                                        " a Reynold's number greater than 1000", , True)
                Else
                   Nu = 0.097 * Reynolds ^ 0.73 * Prandtl ^ 0.13 * (Prandtl / PrandtlInterface) ^ 0.25
                End If
            End If
        
        Case ValidNames(2)
            ' Table 1 No 6: M. Gryta, M. Tomaszewska, M.A. Morawski, Membrane distillation with
            '    laminar flow, Sep. Purif. Technol. 11 (1997) 93–101.
            ' http://refhub.elsevier.com/S1383-5866(14)00771-0/h0030
                If Reynolds < 150 Or Reynolds > 3500 Then
                    FlatSheetLaminarFlow = mdlError.ReturnError("mdlNusselt.FlatSheetLaminarFlow: The Relationship in Gryta,1997 Table 1 Equation No. 6 is only valid " & _
                               " for a Reynolds number range of 150 to 3500.  The reynolds number for the requested flow is " & CStr(Reynolds), , True)
                Else
                    Nu = 0.298 * Reynolds ^ 0.646 * Prandtl ^ 0.316
                End If
        Case ValidNames(3)
            ' Appendix A:
        Case Else
            FlatSheetLaminarFlow = mdlError.NameError("mdlNusselt.FlatSheetLaminarFlow", _
                                                                   "RelationshipToUse", ValidNames)
    End Select
    FlatSheetLaminarFlow = Nu
End If

End Function

Public Function FlatSheetTurbulentFlow(Reynolds As Double, _
                                                  Prandtl As Double, _
                                                  RelationShipToUse As String, _
                                                  Thickness As Double, _
                                                  Spacer As clsSpacer, _
                                                  Optional PrandtlInterface As Double) As Double

' YOU CANNOT USE Spacer For ANY PROPERTY THAT IS CHANGING. I GOT BURNED BECAUSE THE SPACER CLASS CHANGES HOT AND COLD SPACERS SIMULTANEOUSLY
' IF YOU CHANGE THE THICKNESS HAS TO BE VARIED WITH MULTILAYER RUNS SO THAT 1/2 the hot or 1/2 the cold thickness is used assymetrically.

' This comes from 2013, Alsaadi et. al. "Modeling of air gap membrane distillation process - a theoretical and experimental study"
' Zhang et al: [28]
' http://refhub.elsevier.com/S0376-7388(13)00470-5/sbref23

Dim Nu As Double
Dim ValidNames() As String
Dim kS As Double
ReDim ValidNames(0)
ValidNames(0) = "Alsaadi_et_al"


If Reynolds > glbMaxTurbulentReynolds Then
    err.Source = "mdlNusselt.FlatSheetTurbulentFlow: Reynolds number (" & _
              Reynolds & ") must be less than " & glbMaxTurbulentReynolds & " for turbulent flow relationships!"
    FlatSheetTurbulentFlow = mdlConstants.glbINVALID_VALUE
Else
    Select Case RelationShipToUse
        Case ValidNames(0)
            ' This never changes, it would be best to store this in a class so that it doesn't have to get calculated over and over again.
            kS = 1.904 * (Spacer.FilamentDiameter / Thickness) ^ -0.039 * Spacer.Porosity ^ 0.75 * _
                  (Sin(Spacer.FilamentIntersectAngle * glbDegreesToRadians / 2)) ^ 0.086
            Nu = 0.029 * kS * Reynolds ^ 0.8 * Prandtl ^ 0.33
        Case Else
            FlatSheetTurbulentFlow = mdlError.NameError("mdlNusselt.FlatSheetTurbulentFlow", _
                                                                   "RelationshipToUse", ValidNames)
    End Select
    FlatSheetTurbulentFlow = Nu
End If

End Function

' These functions are older

Public Function NusseltNumberRectangularLaminar(Lz As Double, L As Double) As Double

' Constant nusselt number for laminar flow and constant wall heat flux

' Fit comes from London "Heat Transfer 2nd ed.", Capstone Publishing Corporation 2000 page 465.
Dim Lstar As Double

Lstar = Lz / L

If Lstar < 0 Or Lstar > 1 Then
    ErrorTalk 1, MsgBoxInc
Else
' Polynomial fit contained in RectangularChannelFanningFrictionFactor.xlsx
    NusseltNumberRectangularLaminar = 3.10458 * Lstar ^ 6 + 4.78884 * Lstar ^ 5 - _
                                      22.8042 * Lstar ^ 4 + 15.0881 * Lstar ^ 3 + _
                                      8.46907 * Lstar ^ 2 - 13.1395 * Lstar + _
                                      7.99477
End If



End Function

Public Function NusseltNumberTurbulent(Re As Double, Pr As Double) As Double

    Dim f As Double
    
        f = (1.58 * Log(Re) - 3.28) ^ (-2)
        NusseltNumberTurbulent = (f / 2) * Re * Pr / (1.07 + 12.7 * ((f / 2) ^ (1 / 2)) * (((Pr ^ (2 / 3)) - 1)))

End Function

Public Function NusseltNumberNaturalConvection(Ra As Double, Pr As Double) As Double
    
    If Ra < 0 Or Ra > 100000000000000# Then 'Invalid
       NusseltNumberNaturalConvection = mdlError.ReturnError("mdlNusselt.NusseltNumberNaturalConvection: The valid range for the Rayleigh " & _
                                                             "number for these empirical relationships is between zero and less than 10^14!", _
                                                             , True)
    ElseIf Ra < 1000000000# Then 'Laminar equation 9-18 of London page 569
       NusseltNumberNaturalConvection = 0.68 + (0.67 * Ra ^ 0.25 / (1 + (0.492 / Pr) ^ (9 / 16)) ^ (4 / 9))
    Else 'Turbulent equation 9-17 of London page 569
       NusseltNumberNaturalConvection = (0.825 + ((0.387 * Ra ^ (1 / 6)) / (1 + (0.492 / Pr) ^ (9 / 16)) ^ (8 / 27))) ^ 2
    End If

End Function


