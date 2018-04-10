VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWaterStream"
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

Private PrPH As Double
Private Temp As Double
Private MassFl As Double
Private VolFl As Double 'Volumetric flow
Private Pres As Double
Private Salin As Double
Private Cond As Double
Private PrName As String

Private ErrMsg As String
Private ErrMsgEnd As String

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

Private Sub Class_Initialize()
'Default Water Stream Properties
PrName = "Default Water Stream"
Temp = 323.15 'Kelvin
MassFl = 1    'kg/s
Pres = 120000   'Pa
Salin = 0.000001 'Fraction
Cond = 3000      'Not sure what to put in for this value.  Currently not in use
End Sub

Property Let Name(Str As String)
    PrName = Str
    ErrMsgEnd = " Check water stream """ & Str & """ in the ""WaterStreams"" sheet"
End Property
Property Get Name() As String
    Name = PrName
End Property

'  input temperature
Property Let Temperature(Val As Double)
    Temp = AssignValueIfInLimits(Val, ErrMsg, "temperature", glbINVALID_VALUE, "", glbWaterTempInLowerLimit, glbWaterTempInUpperLimit)
    
    ' Now the mass flow has to change (assume volumetric flow stays constant)
    MassFl = VolFl * Me.Density
    
End Property

Property Get Temperature() As Double
    Temperature = Temp + glbCelciusToKelvinOffset
End Property

'  input volumetric flow rate
Property Let VolumeFlow(Val As Double)
   VolFl = glbLPMToMeter3PerSecond * AssignValueIfInLimits(Val, ErrMsg, "volumetric flow rate", glbINVALID_VALUE, "", glbWaterVolumeFlowLowerLimit, glbWaterVolumeFlowUpperLimit)
   ' Recalculate mass flow as a result
   MassFl = VolFl * Me.Density
End Property

'  input mass flow rate
Property Let MassFlow(Val As Double)
    
    MassFl = AssignValueIfInLimits(Val, ErrMsg, "mass flow rate", glbINVALID_VALUE, "", glbWaterMassFlowLowerLimit, glbWaterMassFlowUpperLimit)
    ' Recalculate volumetric flow as a result (assume temperature stays constant)
    VolFl = MassFl / Me.Density
    
End Property
Property Get MassFlow() As Double
    MassFlow = MassFl
End Property

Property Get Density() As Double
    Density = mdlProperties.SeaWaterDensity(Me.Temperature, Me.Salinity)
End Property

'  input pressure
Property Let Pressure(Val As Double)
    Pres = AssignValueIfInLimits(Val, ErrMsg, "pressure", glbINVALID_VALUE, "", glbWaterMassFlowLowerLimit, glbWaterMassFlowUpperLimit)
End Property
Property Get Pressure() As Double
    Pressure = glbBarToPascal * Pres
End Property

' input PH
Property Let PH(Val As Double)
    PrPH = AssignValueIfInLimits(Val, ErrMsg, "PH", glbINVALID_VALUE, "", glbWaterPHLowerLimit, glbWaterPHUpperLimit)
End Property
Property Get PH() As Double
    PH = PrPH
End Property

' input salinity
Property Let Salinity(Val As Double)
    Salin = AssignValueIfInLimits(Val, ErrMsg, "Salinity", glbINVALID_VALUE, "", glbWaterSalinityLowerLimit, glbWaterSalinityUpperLimit)
    ' No the mass flow changes (assume volume flow stays constant)
    MassFl = VolFl * Me.Density
End Property
Property Get Salinity() As Double
    Salinity = glbGramPerKilogramToFraction * Salin
End Property

' input conductance
Property Let Conductance(Val As Double)
    Cond = AssignValueIfInLimits(Val, ErrMsg, "Conductance", glbINVALID_VALUE, "", glbWaterConductanceLowerLimit, glbWaterConductanceUpperLimit)
End Property
Property Get Conductance() As Double
    Salinity = glbMicroSiemenToSiemen * Cond
End Property
