VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMaterial"
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

Private ThermCond As Double
Private MatName As String
Private Thick As Double
Private UseDirectInputForMassTransferCoeff As Boolean
Private DirectMassTransferCoeff As Double
Private Pore As Double
Private MeanPoreRad As Double
Private ErrMsg As String
Private ErrMsgEnd As String
Private MatType As String
Private ReqPropInd As Long

Property Get ErrorMessage() As String
    If Len(ErrMsg) = 0 Then
        ErrorMessage = ""
    Else
        ErrorMessage = ErrMsg & ErrMsgEnd
    End If
End Property

Property Get ErrorMessageEnd() As String
    ErrorMessageEnd = ErrMsgEnd
End Property

Property Let MaterialType(Str As String)
    Dim Var As Variant
    Dim Bool As Boolean
    Dim ValidNames() As String

    ReDim Var(0)
    Bool = ArrayMod_shared.GetColumn(Range("RangeMaterialTypeNames").Value, Var, 1)
    CStrArr Var, ValidNames

    ReqPropInd = StringModule_shared.FindFirstMatchInStringArray(Str, ValidNames)
    
    If ReqPropInd <> -1 Then
       MatType = Str
    Else
       Var = mdlError.ReturnError("clsMaterial.MaterialType: For material " & Me.Name & _
       " The requested materials type " & Str & " is not a valid material type!", , True)
    End If
End Property

Function SetProperty(PropName As String, Optional DblVal As Double, Optional StrVal As String, Optional BoolVal As Boolean)
    Select Case PropName
        Case "ThermalConductivity"
           Me.ThermalConductivity = DblVal
        Case "DirectMassTransferCoefficient"
           Me.DirectMassTransferCoefficient = DblVal
        Case "MaterialType"
           Me.MaterialType = StrVal
        Case "MeanPoreRadius"
           Me.MeanPoreRadius = DblVal
        Case "Porosity"
           Me.Porosity = DblVal
        Case "Thickness"
           Me.Thickness = DblVal
        Case "UseDirectInputForMassTransferCoefficient"
           Me.UseDirectInputForMassTransferCoefficient = BoolVal
        Case Else
            Dim Dum As Double
            Dum = mdlError.ReturnError("clsMaterial.SetProperty: The requested property name " & PropName & " does not exist for clsMaterial or needs to be added to the GetProperty function.", , True)
    End Select
End Function

Property Get RequiredPropertyIndex() As Long
    RequiredPropertyIndex = ReqPropInd
End Property

Property Get MaterialType() As String
     MaterialType = MatType
End Property

Private Sub Class_Initialize()
   'Defaults on instantiation
    ErrMsgEnd = ""
    ErrMsg = ""
    ThermCond = 1
    MatName = "Default"
    Thick = 1
    UseDirectInputForMassTransferCoeff = True
    DirectMassTransferCoeff = 0.0000005
    Pore = glbINVALID_VALUE
    MeanPoreRad = glbINVALID_VALUE
End Sub

Property Let ThermalConductivity(Val As Double)
    ThermCond = AssignValueIfInLimits(Val, ErrMsg, "number of frames", glbINVALID_VALUE, "W/m/K", glbThermConductLowerLimit, glbThermConductUpperLimit)
End Property
Property Get ThermalConductivity() As Double
    ThermalConductivity = ThermCond
End Property

Property Let Name(NameStr As String)
    MatName = NameStr
    ErrMsgEnd = " Check material """ & NameStr & """ in the ""MembranesAndFoils"" sheet"
End Property
Property Get Name() As String
    Name = MatName
End Property

Property Let Thickness(Val As Double) 'Thickness is input as milimeters
    Thick = AssignValueIfInLimits(Val, ErrMsg, "material thickness", glbINVALID_VALUE, "mm", glbThickLowerLimit, glbThickUpperLimit)
End Property
Property Get Thickness() As Double 'Thickness is always output in meters
    Thickness = glbMiliMeterToMeter * Thick
End Property

Property Let UseDirectInputForMassTransferCoefficient(Bool As Boolean)
    UseDirectInputForMassTransferCoeff = Bool
End Property
Property Get UseDirectInputForMassTransferCoefficient() As Boolean
    UseDirectInputForMassTransferCoefficient = UseDirectInputForMassTransferCoeff
End Property

Property Let DirectMassTransferCoefficient(Val As Double)
    DirectMassTransferCoeff = AssignValueIfInLimits(Val, ErrMsg, "direct mass transfer coefficient", glbINVALID_VALUE, "kg/s/m2/Pa", _
                                                    glbDirectMassTransferCoeffLowerLimit, glbDirectMassTransferCoeffUpperLimit)
End Property
Property Get DirectMassTransferCoefficient() As Double
    DirectMassTransferCoefficient = DirectMassTransferCoeff
End Property

Property Let Porosity(Val As Double)
    Pore = AssignValueIfInLimits(Val, ErrMsg, "porosity", glbINVALID_VALUE, "", glbPorosityLowerLimit, glbPorosityUpperLimit)
End Property
Property Get Porosity() As Double
    Porosity = Pore
End Property

Property Let MeanPoreRadius(Val As Double) 'input as milimeters output as meters
    MeanPoreRad = AssignValueIfInLimits(Val, ErrMsg, "mean pore radius", glbINVALID_VALUE, "kg/s/m2/Pa", glbMeanPoreRadLowerLimit, Thick)
End Property
Property Get MeanPoreRadius() As Double ' output as meters
    MeanPoreRadius = glbMiliMeterToMeter * MeanPoreRad
End Property




