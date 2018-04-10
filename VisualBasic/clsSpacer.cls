VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSpacer"
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

Private SpacerName As String
Private Thick As Double
Private Por As Double 'Porosity
Private FilDiam As Double 'Filament Diameter
Private FilIntersectAngle As Double 'Filament intersection angle
Private ErrMsg As String
Private ErrMsgEnd As String
Private ThermCond As Double

Property Get ErrorMessage() As String
    If Len(ErrMsg) = 0 Then
        ErrorMessage = ""
    Else
        ErrorMessage = ErrMsg & ErrMsgEnd
    End If
End Property

Property Get ErrorMessageEnd() As String
    ErrMessageEnd = ErrMsgEnd
End Property

Property Get ThermalConductivity() As Double
    ThermalConductivity = ThermCond
End Property

Property Let ThermalConductivity(Val As Double)
   ThermCond = AssignValueIfInLimits(Val, ErrMsg, "thermal conductivity", glbINVALID_VALUE, "", glbThermConductLowerLimit, glbThermConductUpperLimit)
End Property

Private Sub Class_Initialize()
   'Defaults on instantiation
    ErrMsgEnd = ""
    ErrMsg = ""
    SpacerName = "Default"
    Thick = 1
    Por = glbINVALID_VALUE
End Sub

Property Let Name(NameStr As String)
    SpacerName = NameStr
    ErrMsgEnd = " Check material """ & NameStr & """ in the ""MembranesAndFoils"" sheet"
End Property
Property Get Name() As String
    Name = SpacerName
End Property

Property Let Thickness(Val As Double) 'Thickness is input as milimeters
    Thick = AssignValueIfInLimits(Val, ErrMsg, "spacer thickness", glbINVALID_VALUE, "mm", glbThickLowerLimit, glbThickUpperLimit)
End Property
Property Get Thickness() As Double 'Thickness is always output in meters ' ADDED 8/24/2017 - thickness multiplier is needed if the channel is being modeled as symmetric!
    Thickness = glbMiliMeterToMeter * Thick
End Property

Property Let Porosity(Val As Double)
    Por = AssignValueIfInLimits(Val, ErrMsg, "porosity", glbINVALID_VALUE, "", glbPorosityLowerLimit, glbPorosityUpperLimit)
End Property
Property Get Porosity() As Double
    Porosity = Por
End Property

Property Let FilamentDiameter(Val As Double) 'Filament Diameter is input as milimeters
    FilDiam = AssignValueIfInLimits(Val, ErrMsg, "filament diameter", glbINVALID_VALUE, "mm", glbThickLowerLimit, glbThickUpperLimit)
End Property
Property Get FilamentDiameter() As Double 'Filament Diameter is always output in meters
    FilamentDiameter = glbMiliMeterToMeter * FilDiam
End Property

Property Let FilamentIntersectAngle(Val As Double) '
    FilIntersectAngle = AssignValueIfInLimits(Val, ErrMsg, "filament intersect angle", glbINVALID_VALUE, "degrees", glbFilamentIntersectAngleLowerLimit, glbFilamentIntersectAngleUpperLimit)
End Property
Property Get FilamentIntersectAngle() As Double '
    FilamentIntersectAngle = FilIntersectAngle
End Property
