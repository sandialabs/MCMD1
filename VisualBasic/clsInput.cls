VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInput"
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

Private ErrMsg As String
Private Mat As Collection ' private materials collection
Private Spa As Collection ' Spacers
Private H2OStreams As Collection '

Private FoilMat As String
Private MembMat As String
Private AirGapSpa As String
Private HotSpa As String
Private ColdSpa As String
Private HotWS As String
Private ColdWS As String
Private MDType As String
Private NLayers As Long
Private AirGapPres As Double ' In bar - the simulation needs Pascals.
Private ExtInsulMat As String

Private Horizontal_Length As Double
Private Vertical_Length As Double
Private ExtTemperature As Double
Private NumExpExtSides As Long
Private IncludeExtLoss As Boolean
Private ExtLossIsHotSide As Boolean

Private Sub ReadFromInputPage(Optional ErrorOccured As Boolean = False)
   ErrorOccured = False
   FoilMat = Sheet4.Combo_FoilMaterial.Value
   MembMat = Sheet4.Combo_MembraneMaterial.Value
   HotWS = Sheet4.Combo_HotWaterStream.Value
   HotSpa = Sheet4.Combo_HotSpacer.Value
   ColdWS = Sheet4.Combo_ColdWaterStreams.Value
   ColdSpa = Sheet4.Combo_ColdSpacer.Value
   AirGapSpa = Sheet4.Combo_AirGapSpacer.Value
   MDType = Sheet4.Combo_MD_Type
   
   
   
   On Error GoTo ErrorHappened
   
   Dim HWS As clsWaterStream
   
   Set HWS = Me.WaterStreams(Me.HotWaterStream)
   
  AirGapPres = AssignValueIfInLimits(Range("AirGapPressure"), ErrMsg, "Air Gap Pressure", mdlConstants.glbINVALID_VALUE, "bar", _
                                        mdlConstants.glbAirPressureLowerLimit, 2 * HWS.Pressure)
  If Len(ErrMsg) <> 0 Then
     ErrMsg = ErrMsg & vbCrLf & vbCrLf & "The air gap pressure cannot exceed 2 * the hot water stream pressure! Please reduce the air gap pressure or " & _
               " increase the hot water stream (" & HWS.Name & " pressure."
  End If
   
  Horizontal_Length = AssignValueIfInLimits(Range("Horizontal_Length"), ErrMsg, "Horizontal Length", mdlConstants.glbINVALID_VALUE, "meters", _
                                            mdlConstants.glbPlateDimensionsLowerLimit, mdlConstants.glbPlateDimensionsUpperLimit)
  Vertical_Length = AssignValueIfInLimits(Range("Vertical_Length"), ErrMsg, "Horizontal Length", mdlConstants.glbINVALID_VALUE, "meters", _
                                            mdlConstants.glbPlateDimensionsLowerLimit, mdlConstants.glbPlateDimensionsUpperLimit)
  NLayers = AssignValueIfInLimits(Range("NumberOfLayers"), ErrMsg, "Number of layers", glbINVALID_VALUE, "", _
                                            mdlConstants.glbNumberOfLayersLowerLimit, mdlConstants.glbNumberOfLayersUpperLimit)
  IncludeExtLoss = Sheet4.CheckBox_IncludeAmbient
  
  If IncludeExtLoss Then
     ExtTemperature = AssignValueIfInLimits(Range("AmbientTemperatureRange"), ErrMsg, "External Ambient Temperature", glbINVALID_VALUE, "", _
                                            mdlConstants.glbAmbientTemperatureLowerLimit, mdlConstants.glbAmbientTemperatureUpperLimit)
     ExtInsulMat = Sheet4.Combo_ExternalMaterial
  
     If Sheet4.OptionButton1 Then
       NumExpExtSides = 1
     Else
       NumExpExtSides = 2
     End If
     
     ExtLossIsHotSide = Sheet4.CheckBox_HotSideExposed
  End If
  
Exit Sub
ErrorHappened:
   ErrorOccured = True
End Sub

Property Get ExternalLossIsHotSide() As Boolean
    ExternalLossIsHotSide = ExtLossIsHotSide
End Property

Property Get ExternalInsulationMaterial() As clsMaterial
    Set ExternalInsulationMaterial = Materials(ExtInsulMat)
End Property

Property Get ExternalAmbientTemperature() As Double
    ExternalAmbientTemperature = ExtTemperature + mdlConstants.glbCelciusToKelvinOffset
End Property

Property Get IncludeExternalLosses() As Boolean
    IncludeExternalLosses = IncludeExtLoss
End Property

Property Get NumberOfExposedSides() As Long
    NumberOfExposedSides = NumExpExtSides
End Property

Property Get ColdSpacer() As String
   ColdSpacer = ColdSpa
End Property

Property Get HotSpacer() As String
   HotSpacer = HotSpa
End Property

Property Get AirGapSpacer() As String
   AirGapSpacer = AirGapSpa
End Property

Property Get AirGapPressure() As Double
   AirGapPressure = mdlConstants.glbBarToPascal * AirGapPres
End Property

Property Get FoilMaterial() As String
   FoilMaterial = FoilMat
End Property

Property Get MembraneMaterial() As String
   MembraneMaterial = MembMat
End Property

Property Get HotWaterStream() As String
   HotWaterStream = HotWS
End Property

Property Get ColdWaterStream() As String
   ColdWaterStream = ColdWS
End Property

Property Get MembraneDistillationType() As String
   MembraneDistillationType = MDType
End Property

Property Let HorizontalLength(Val As Double)
   Horizontal_Length = AssignValueIfInLimits(Val, ErrMsg, "Horizontal Length", mdlConstants.glbINVALID_VALUE, "meters", _
                                            mdlConstants.glbPlateDimensionsLowerLimit, mdlConstants.glbPlateDimensionsUpperLimit)
End Property

Property Get HorizontalLength() As Double
    HorizontalLength = Horizontal_Length
End Property



Property Let VerticalLength(Val As Double)
   Vertical_Length = AssignValueIfInLimits(Val, ErrMsg, "Horizontal Length", mdlConstants.glbINVALID_VALUE, "meters", _
                                            mdlConstants.glbPlateDimensionsLowerLimit, mdlConstants.glbPlateDimensionsUpperLimit)
End Property
 
Property Get VerticalLength() As Double
    VerticalLength = Vertical_Length
End Property

Property Get NumberOfLayers() As Long
   NumberOfLayers = NLayers
End Property

Property Get ErrorMessage() As String
    ErrorMessage = ErrMsg
End Property

Public Property Get WaterStreams() As Collection
    Set WaterStreams = H2OStreams
End Property

Public Property Get Materials() As Collection
      Set Materials = Mat
End Property

Public Property Get MembraneMaterialObj() As clsMaterial
    Set MembraneMaterialObj = Mat.Item(MembMat)
End Property

Public Property Get Spacers() As Collection
      Set Spacers = Spa
End Property

Private Sub ReadMembraneAndFoilProperties(Wksh As Worksheet, ErrMsg As String)

' Nearly the same as ReadMembraneAndFoilProperties - same concept for spacers...If only VBA was more flexible with its class structures.

Dim CR As Range 'CR = current range
Dim TempMat As clsMaterial
Dim Var As Variant
Dim Ind As Long
Dim ErrMsgEnd As String
Dim NumBlank As Long
Dim ReqMatProp() As Long
Dim ReadRequiredProperties As Boolean
Dim i As Long
Dim MatPropDataTypes As Variant
Dim MatPropNames As Variant
Dim MatInd As Long
Dim Bool As Boolean


Set CR = Wksh.Range("Read_Start_Point_Materials")
Ind = 1

NumBlank = 0

MatPropDataTypes = Range("RangeMaterialPropertyTypes").Value
MatPropNames = Range("RangeclsMaterialPropertyNames").Value

Set Mat = New Collection

Do While NumBlank < glbNumberOfLinesInExcelToSearch

    Var = CR.Offset(Ind, 0).Value
    
    Set TempMat = New clsMaterial
    
    TempMat.Name = CStr(Var)
    
    If Len(TempMat.Name) = 0 Then
        NumBlank = NumBlank + 1
    Else
        On Error GoTo TypeError
        
        Var = CR.Offset(Ind, 1).Value
        TempMat.MaterialType = CStr(Var)
        ' NO error procedure is needed here because a direct error is called
        ' Figure out which properties are required
        ReDim Var(0)
        Bool = ArrayMod_shared.GetColumn(Range("RangeRequiredMaterialProperties").Value, Var, 1)
        CLngArr Split(Var(TempMat.RequiredPropertyIndex), ","), ReqMatProp, ReadRequiredProperties
        
        If Not ReadRequiredProperties Then
           Var = mdlError.ReturnError("clsInput.ReadMembraneAndFoilProperties: The required properties in the ""RangeRequiredMaterialProperties"" " & _
            "named range are not properly separated by commas or are not of type integer! This needs fixing before the code can run!", , True)
        End If
        
        ' Read all required properties
        For i = LBound(ReqMatProp) To UBound(ReqMatProp)
            MatInd = ReqMatProp(i)
            Var = CR.Offset(Ind, MatInd).Value
            If IsEmpty(Var) Then
               Var = mdlError.ReturnError("clsInput.ReadMembraneAndFoilProperties: A required value for material """ & TempMat.Name & _
                      """ property """ & MatPropNames(MatInd, 1) & """ has no entry! Correct this on the ""MembranesAndFoils"" worksheet.", , True)
            End If
            Select Case MatPropDataTypes(MatInd, 1)
               Case "Double"
                   TempMat.SetProperty CStr(MatPropNames(MatInd, 1)), CDbl(Var)
               Case "String"
                   TempMat.SetProperty CStr(MatPropNames(MatInd, 1)), , CStr(Var)
               Case "Boolean"
                   TempMat.SetProperty CStr(MatPropNames(MatInd, 1)), , , CBool(Var)
               Case Else
                  Var = mdlError.ReturnError("clsInput.ReadMembraneAndFoilProperties: The requested data types available are" & _
                    """Double"", ""String"", and ""Boolean."" " & MatPropDataTypes(MatInd) & " was requested. Please correct " & _
                    " the corresponding entry in the ""RangeMaterialPropertyTypes"" named range on the Configurations worksheet.", , True)
            End Select
            ' Now see if the assignment worked.
            If Len(TempMat.ErrorMessage) <> 0 Then
                GoTo InputError
            End If
        Next i
        
        'Now we have a couple of custom functions for membrane properties
        If TempMat.MaterialType = "Membrane" Then
            If TempMat.UseDirectInputForMassTransferCoefficient Then
                Var = CR.Offset(Ind, 5).Value
                If IsEmpty(Var) Then
                   Var = mdlError.ReturnError("clsInput.ReadMembraneAndFoilProperties: A required value for material """ & TempMat.Name & _
                          """ property ""DirectMassTransferCoefficient"" has no entry! Correct this on the ""MembranesAndFoils"" worksheet.", , True)
                End If
                TempMat.DirectMassTransferCoefficient = CDbl(Var)
                If Len(TempMat.ErrorMessage) <> 0 Then
                    GoTo InputError
                End If
            End If
            
            ' These values will be used IF they are included even if a direct input mass transfer coefficient is provided.
            Var = CR.Offset(Ind, 6).Value
            If Not IsEmpty(Var) Then
                TempMat.Porosity = CDbl(Var)
                If Len(TempMat.ErrorMessage) <> 0 Then
                    GoTo InputError
                End If
            Else
                If Not TempMat.UseDirectInputForMassTransferCoefficient Then
                    ErrMsg = "clsInput.ReadMembraneAndFoilProperties: Material " & TempMat.Name & _
                    ": The porosity must be input if a direct mass transfer coefficient is not supplied. Correct this on the ""MembranesAndFoils"" worksheet."
                    GoTo EndOfSub
                End If
            End If
            
            Var = CR.Offset(Ind, 7).Value
            If Not IsEmpty(Var) Then
               TempMat.MeanPoreRadius = CDbl(Var)
               If Len(TempMat.ErrorMessage) <> 0 Then
                  GoTo InputError
               End If
            Else
               If Not TempMat.UseDirectInputForMassTransferCoefficient Then
                    ErrMsg = "clsInput.ReadMembraneAndFoilProperties: Material " & TempMat.Name & _
                    ": The mean pore radius must be input if a direct mass transfer coefficient is not supplied. Correct this on the ""MembranesAndFoils"" worksheet."
                    GoTo EndOfSub
               End If
            End If
        End If
        
        ' In the future this is where new properties can be added CR.Offset(Ind, 8).Value
        
        On Error GoTo 0
        
        If Len(TempMat.ErrorMessage) = 0 Then
            Mat.Add TempMat, TempMat.Name
        Else
InputError:
            ErrMsg = TempMat.ErrorMessage
            GoTo EndOfSub
        End If
    End If
    
    Ind = Ind + 1
    
Loop
EndOfSub:
    If Len(ErrMsg) <> 0 Then
       Var = mdlError.ReturnError(ErrMsg, , True)
    End If
Exit Sub

TypeError:
    ErrMsg = "Invalid data type in one of the 7 material properties inputs!" & TempMat.ErrorMessageEnd
GoTo EndOfSub
End Sub


Private Sub ReadSpacerProperties(Wksh As Worksheet, ErrMsg As String)

Dim CR As Range 'CR = current range
Dim TempSpa As clsSpacer
Dim Var As Variant
Dim Ind As Long
Dim ErrMsgEnd As String
Dim NumBlank As Long

Set CR = Wksh.Range("Start_Point_Spacers")
Ind = 1

NumBlank = 0

Set Spa = New Collection

Do While NumBlank < glbNumberOfLinesInExcelToSearch

    Var = CR.Offset(Ind, 0).Value
    
    Set TempSpa = New clsSpacer
    
    TempSpa.Name = CStr(Var)
    
    If Len(TempSpa.Name) = 0 Then
        NumBlank = NumBlank + 1
    Else
        On Error GoTo TypeError
        
        'The first three properties are essential
        Var = CR.Offset(Ind, 1).Value
        TempSpa.Thickness = CDbl(Var)
        If Len(TempSpa.ErrorMessage) <> 0 Then
            GoTo InputError
        End If
        Var = CR.Offset(Ind, 2).Value
        TempSpa.FilamentDiameter = CDbl(Var)
        If Len(TempSpa.ErrorMessage) <> 0 Then
            GoTo InputError
        End If
        
        Var = CR.Offset(Ind, 3).Value
        TempSpa.FilamentIntersectAngle = CDbl(Var)
        If Len(TempSpa.ErrorMessage) <> 0 Then
            GoTo InputError
        End If
        
        Var = CR.Offset(Ind, 4).Value
        TempSpa.Porosity = CDbl(Var)
        If Len(TempSpa.ErrorMessage) <> 0 Then
            GoTo InputError
        End If
        
        Var = CR.Offset(Ind, 5).Value
        TempSpa.ThermalConductivity = CDbl(Var)
        If Len(TempSpa.ErrorMessage) <> 0 Then
           GoTo InputError
        End If
        
        On Error GoTo 0
        
        If Len(TempSpa.ErrorMessage) = 0 Then
            Spa.Add TempSpa, TempSpa.Name
        Else
InputError:
            ErrMsg = TempSpa.ErrorMessage
            GoTo EndOfSub
        End If
    End If
    
    Ind = Ind + 1
    
Loop
EndOfSub:
    
Exit Sub

TypeError:
    ErrMsg = "Invalid data type in one of the 4 spacer properties inputs!" & TempSpa.ErrorMessageEnd
GoTo EndOfSub
End Sub

Public Sub DebugInputs()

If glbEvaluateEquationsAtEnd Then
    
Dim FoilMatObj As clsMaterial
Dim MembMatObj As clsMaterial
Dim AirGapSpaObj As clsSpacer
Dim HotSpaObj As clsSpacer
Dim ColdSpaObj As clsSpacer
Dim HotWSObj As clsWaterStream
Dim ColdWSObj As clsWaterStream
Dim DebugString As String

Set FoilMatObj = Me.Materials(Me.FoilMaterial)
Set MembMatObj = Me.Materials(Me.MembraneMaterial)
Set AirGapSpaObj = Me.Spacers(Me.AirGapSpacer)
Set HotSpaObj = Me.Spacers(Me.HotSpacer)
Set ColdSpaObj = Me.Spacers(Me.ColdSpacer)
Set HotWSObj = Me.WaterStreams(Me.HotWaterStream)
Set ColdWSObj = Me.WaterStreams(Me.ColdWaterStream)

DebugString = DebugString & "INPUTS:" & vbCrLf
DebugString = DebugString & FormatDebugColumns("feed_hot.t:    Hot inlet temperature (C):  " & HotWSObj.Temperature - mdlConstants.glbCelciusToKelvinOffset & vbCrLf)
DebugString = DebugString & FormatDebugColumns("feed_cold.t:   Cold inlet temperature (C): " & ColdWSObj.Temperature - mdlConstants.glbCelciusToKelvinOffset & vbCrLf)
DebugString = DebugString & FormatDebugColumns("feed_hot.w:    Hot Inlet Salinity (kg/kg): " & HotWSObj.Salinity & vbCrLf)
DebugString = DebugString & FormatDebugColumns("feed_cold.w:   Cold Inlet Salinity (kg/kg):    " & ColdWSObj.Salinity & vbCrLf)
DebugString = DebugString & FormatDebugColumns("feed_hot.m:    Hot Inlet Flow rate (kg/s):    " & HotWSObj.MassFlow & vbCrLf)
DebugString = DebugString & FormatDebugColumns("feed_cold.m:   Cold inlet flow rate (kg/s):   " & ColdWSObj.MassFlow & vbCrLf)
DebugString = DebugString & FormatDebugColumns("Kp:    Thermal Conductivity of Membrane Material (W/m/K):  " & MembMatObj.ThermalConductivity & vbCrLf)
DebugString = DebugString & FormatDebugColumns("epsilon:   Porosity of Membrane Material:  " & MembMatObj.Porosity & vbCrLf)
DebugString = DebugString & FormatDebugColumns("r: Membrane Mean Pore Radius (m):  " & MembMatObj.MeanPoreRadius & vbCrLf)
DebugString = DebugString & FormatDebugColumns("delta_h:   Hot Stream Flow Thickness (m):  " & HotSpaObj.Thickness & vbCrLf)
DebugString = DebugString & FormatDebugColumns("delta_c:   Cold Stream Flow Thickness (m): " & ColdSpaObj.Thickness & vbCrLf)
DebugString = DebugString & FormatDebugColumns("delta_m:   Membrane Thickness (m): " & MembMatObj.Thickness & vbCrLf)
DebugString = DebugString & FormatDebugColumns("Kb:    Boltzmann's Constant (J/K): " & mdlConstants.glbBoltzmann & vbCrLf)
DebugString = DebugString & FormatDebugColumns("R: Universal Gas Constant (J/mol/K):   " & mdlConstants.glbGasConstant & vbCrLf)
DebugString = DebugString & FormatDebugColumns("sigma: Collision diameter of water vapor (m):  " & mdlConstants.glbWaterVaporCollisionDiameter & vbCrLf)
DebugString = DebugString & FormatDebugColumns("M: Molecular weight of water (kg/mol): " & mdlConstants.glbWaterMolecularWeight & vbCrLf)
DebugString = DebugString & FormatDebugColumns("t_ref: Celcius to Kelvin offset factor (K):    " & mdlConstants.glbCelciusToKelvinOffset & vbCrLf)
DebugString = DebugString & FormatDebugColumns("Es_h:    Hot Spacer Porosity:    " & HotSpaObj.Porosity & vbCrLf)
DebugString = DebugString & FormatDebugColumns("Es_c:  Cold Spacer Porosity:   " & ColdSpaObj.Porosity & vbCrLf)
DebugString = DebugString & FormatDebugColumns("Dfs_h: Hot Spacer Filament Diameter (m):   " & HotSpaObj.FilamentDiameter & vbCrLf)
DebugString = DebugString & FormatDebugColumns("Dfs_c: Cold Spacer Filament Diameter (m):  " & ColdSpaObj.FilamentDiameter & vbCrLf)
DebugString = DebugString & FormatDebugColumns("delta_s_h:   Hot Spacer Thickness (m) (same as hot flow thickness):   " & HotSpaObj.Thickness & vbCrLf)
DebugString = DebugString & FormatDebugColumns("delta_s_c:   Cold Spacer Thickness (m) (same as cold flow thickness):   " & ColdSpaObj.Thickness & vbCrLf)
DebugString = DebugString & FormatDebugColumns("theta_h:   Hot Spacer filament intersection angle (deg):   " & HotSpaObj.FilamentIntersectAngle & vbCrLf)
DebugString = DebugString & FormatDebugColumns("theta_c:   Cold Spacer filament intersection angle (deg):   " & ColdSpaObj.FilamentIntersectAngle & vbCrLf)
DebugString = DebugString & FormatDebugColumns("n_layers:  Number of layers: " & Me.NumberOfLayers & vbCrLf)

Debug.Print DebugString
mdlValidation.WriteDebugInfoToFile DebugString

End If


End Sub

Private Sub ReadWaterStreamProperties(Wksh As Worksheet, ErrMsg As String)
' Very similar to ReadMembraneAndFoilProperties - if VBA had higher level functionality these could be one function

Dim CR As Range 'CR = current range
Dim TempH2O As clsWaterStream
Dim Var As Variant
Dim Ind As Long
Dim ErrMsgEnd As String
Dim NumBlank As Long

Set CR = Wksh.Range("Read_Start_Point_WaterStreams")
Ind = 1

NumBlank = 0

Set H2OStreams = New Collection

Do While NumBlank < glbNumberOfLinesInExcelToSearch

    Var = CR.Offset(Ind, 0).Value
    
    Set TempH2O = New clsWaterStream
    
    TempH2O.Name = CStr(Var)
    
    If Len(TempH2O.Name) = 0 Then
        NumBlank = NumBlank + 1
    Else
        On Error GoTo TypeError
        
        'The first three properties are essential
        Var = CR.Offset(Ind, 1).Value
        TempH2O.Temperature = CDbl(Var)
        If Len(TempH2O.ErrorMessage) <> 0 Then
            GoTo InputError
        End If
        Var = CR.Offset(Ind, 2).Value
        TempH2O.VolumeFlow = CDbl(Var)
        If Len(TempH2O.ErrorMessage) <> 0 Then
            GoTo InputError
        End If
        Var = CR.Offset(Ind, 3).Value
        TempH2O.Pressure = CDbl(Var)
        If Len(TempH2O.ErrorMessage) <> 0 Then
            GoTo InputError
        End If
        Var = CR.Offset(Ind, 4).Value
        TempH2O.Salinity = CDbl(Var)
        If Len(TempH2O.ErrorMessage) <> 0 Then
            GoTo InputError
        End If
        Var = CR.Offset(Ind, 5).Value
        TempH2O.PH = CDbl(Var)
        If Len(TempH2O.ErrorMessage) <> 0 Then
            GoTo InputError
        End If
        Var = CR.Offset(Ind, 6).Value
        TempH2O.Conductance = CDbl(Var)
        If Len(TempH2O.ErrorMessage) <> 0 Then
            GoTo InputError
        End If
        
        ' Add more properties as needed.
        
        On Error GoTo 0
        
        If Len(TempH2O.ErrorMessage) = 0 Then
           H2OStreams.Add TempH2O, TempH2O.Name
        Else
InputError:
            ErrMsg = TempH2O.ErrorMessage
            GoTo EndOfSub
        End If
    End If
    
    Ind = Ind + 1
    
Loop
EndOfSub:
    
Exit Sub

TypeError:
    ErrMsg = "Invalid data type in one of the 6 material properties inputs!" & TempH2O.ErrorMessageEnd
GoTo EndOfSub
End Sub

Private Sub Class_Initialize()
    Dim Dum As Double
    Set Mat = New Collection
    Set H2OStreams = New Collection
    
    ReadMembraneAndFoilProperties ThisWorkbook.Worksheets("MembranesAndFoils"), ErrMsg
    
    If Len(ErrMsg) = 0 Then
        ReadWaterStreamProperties ThisWorkbook.Worksheets("WaterStreams"), ErrMsg
        If Len(ErrMsg) <> 0 Then
            Dum = mdlError.ReturnError(ErrMsg, , True)
        Else
            ReadSpacerProperties ThisWorkbook.Worksheets("Spacers"), ErrMsg
            
            If Len(ErrMsg) <> 0 Then
               Dum = mdlError.ReturnError(ErrMsg, , True)
            End If
            
            ReadFromInputPage
            If Len(ErrMsg) <> 0 Then
               Dum = mdlError.ReturnError(ErrMsg, , True)
            End If
            
        End If
    Else
        Dum = mdlError.ReturnError(ErrMsg)
    End If
    
End Sub
