Attribute VB_Name = "mdlControls"
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



Public Sub PopulateInputComboBoxes()

Dim Inp As clsInput
Dim Wksh As Worksheet
Dim i As Long
Dim WS As clsWaterStream
Dim Dum As Double



Set Inp = New clsInput

If Len(Inp.ErrorMessage) <> 0 Then
   MsgBox "The errors in the input listed after this message must be fixed before the model can be run properly!"
   Dum = mdlError.ReturnError("mdlControls.PopulateComboBox: " & Inp.ErrorMessage, , True)
End If
Set Wksh = ThisWorkbook.Worksheets("Input")

PopulateComboBoxWithWaterStreams Sheet4.Combo_ColdWaterStreams, Inp.WaterStreams
PopulateComboBoxWithWaterStreams Sheet4.Combo_HotWaterStream, Inp.WaterStreams

PopulateComboBoxWithSpacers Sheet4.Combo_ColdSpacer, Inp.Spacers
PopulateComboBoxWithSpacers Sheet4.Combo_HotSpacer, Inp.Spacers
PopulateComboBoxWithSpacers Sheet4.Combo_AirGapSpacer, Inp.Spacers

PopulateComboBoxWithMaterials Sheet4.Combo_MembraneMaterial, Inp.Materials, "Membrane"
PopulateComboBoxWithMaterials Sheet4.Combo_FoilMaterial, Inp.Materials, "Foil"
PopulateComboBoxWithMaterials Sheet4.Combo_ExternalMaterial, Inp.Materials, "Foil"

PopulateMDTypeCombo

End Sub


Private Sub PopulateComboBoxWithSpacers(Cmb As ComboBox, SpaCol As Collection)

    Dim Spaobj As clsSpacer
    Dim Value As String
    
    Value = Cmb.Value
    
    Cmb.Clear
    
    For i = 1 To SpaCol.count
       Set Spaobj = SpaCol(i)
       Cmb.AddItem Spaobj.Name
    Next i
    
    On Error Resume Next
    Cmb.Value = Value
    If Cmb.ListIndex = -1 Then
       Cmb.Value = Cmb.List(0)
    End If

End Sub

Private Sub PopulateComboBoxWithWaterStreams(Cmb As ComboBox, WSCol As Collection)

    Dim WSobj As clsWaterStream
    Dim Value As String
    
    Value = Cmb.Value
    
    Cmb.Clear
    
    For i = 1 To WSCol.count
       Set WSobj = WSCol(i)
       Cmb.AddItem WSobj.Name
    Next i
    
    On Error Resume Next
    Cmb.Value = Value
    If Cmb.ListIndex = -1 Then
       Cmb.Value = Cmb.List(0)
    End If

End Sub

Private Sub PopulateComboBoxWithMaterials(Cmb As ComboBox, MatCol As Collection, MatType As String)

    Dim Matobj As clsMaterial
    Dim Value As String
    
    Value = Cmb.Value
    Cmb.Clear
    
    For i = 1 To MatCol.count
       Set Matobj = MatCol(i)
       If Matobj.MaterialType = MatType Then
           Cmb.AddItem Matobj.Name
       End If
    Next i
    
    'If it fails to set the old value, then leave it blank
    On Error Resume Next
    Cmb.Value = Value
    If Cmb.ListIndex = -1 Then
       Cmb.Value = Cmb.List(0)
    End If

End Sub

Public Sub PopulateMDTypeCombo()

    Dim i As Long
    Dim Value As String
    
    mdlConstants.GlobalArrays
    
    Value = Sheet4.Combo_MD_Type.Value
    
    Sheet4.Combo_MD_Type.Clear
    For i = LBound(mdlConstants.glbMembraneDistillationTypes) To UBound(mdlConstants.glbMembraneDistillationTypes)
        Sheet4.Combo_MD_Type.AddItem mdlConstants.glbMembraneDistillationTypes(i)
    Next i
    
    On Error Resume Next
    Sheet4.Combo_MD_Type.Value = Value
    If Cmb.ListIndex = -1 Then
       Cmb.Value = Cmb.List(0)
    End If

End Sub

Public Sub ExtraInputsForExternalOnOffSub()

If Sheet4.CheckBox_IncludeAmbient.Value Then
   Sheet4.Combo_ExternalMaterial.Visible = True
   Sheet4.Range("AmbientTemperatureRange").Font.ColorIndex = 1
   Sheet4.Range("AmbientTemperatureRange").BorderAround Weight:=xlThin, ColorIndex:=xlColorIndexAutomatic, Color:=1
   Sheet4.Range("ExtraInputsForExternalRange").Font.ColorIndex = 1
   Sheet4.OptionButton1.Visible = True
   Sheet4.OptionButton2.Visible = True
   If Sheet4.OptionButton1 Then
      Sheet4.CheckBox_HotSideExposed.Visible = True
   End If
Else
   Sheet4.Combo_ExternalMaterial.Visible = False
   Sheet4.Range("AmbientTemperatureRange").Font.ColorIndex = 2
   Sheet4.Range("AmbientTemperatureRange").Borders.ColorIndex = 2
   Sheet4.Range("ExtraInputsForExternalRange").Font.ColorIndex = 2
   Sheet4.OptionButton1.Visible = False
   Sheet4.OptionButton2.Visible = False
   Sheet4.CheckBox_HotSideExposed.Visible = False
End If
   
End Sub

Public Sub MakeHotSideOptionVisible()
    If Sheet4.OptionButton1 Then
       Sheet4.CheckBox_HotSideExposed.Visible = True
    ElseIf Sheet4.OptionButton2 Then
       Sheet4.CheckBox_HotSideExposed.Visible = False
    End If
End Sub



' Excel macro https://gist.github.com/steve-jansen/7589478
' to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
Public Sub ExportVisualBasicCode()
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    Dim fso As New FileSystemObject
    
    directory = ActiveWorkbook.path & "\VisualBasic"
    count = 0
    
    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If
    Set fso = Nothing
    
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select
            
                
        On Error Resume Next
        err.Clear
        
        path = directory & "\" & VBComponent.Name & extension
        Call VBComponent.Export(path)
        
        If err.Number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.Name & " to " & path, vbCritical)
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If

        On Error GoTo 0
    Next
    
    Application.StatusBar = "Successfully exported " & CStr(count) & " VBA files to " & directory
    Application.OnTime Now + TimeSerial(0, 0, 10), "ClearStatusBar"
End Sub

