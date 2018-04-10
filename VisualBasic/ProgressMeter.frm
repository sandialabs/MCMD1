VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressMeter 
   Caption         =   "Parameter Study Status"
   ClientHeight    =   2700
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   5244
   OleObjectBlob   =   "ProgressMeter.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Private Const TargetWidth As Long = 210 '!@#$ Manual Input

Dim TargFrac As Double
Dim NumSteps As Long
Dim CurFrac As Double
Dim StartFrac As Double
Dim IssueMessage As Boolean
Dim NoClose As Boolean




Private Sub CancelRunButton_Click()
    Dim Ans As Long
    
    If IssueMessage Then
        Ans = MsgBox("Canceling will erase all results.", vbOKCancel, "Cancel Run?")
    Else
        Ans = vbOK
    End If
    
    Select Case Ans
       Case vbOK
            Me.Label2.Caption = "Canceling ... Please Wait"
            Me.CancelRunButton.Tag = 1
       Case vbCancel
            'Nothing to do, continue the run.
       
    End Select
    
End Sub

Property Let TargetFraction(Value As Double)
     If Value > 1 Then
        TargFrac = 1
     ElseIf Value < CurFrac Then
        TargFrac = CurFrac
     Else
        TargFrac = Value
     End If
     
     StartFrac = CurFrac
End Property

Property Let DoNotAllowClose(Value As Boolean)
     NoClose = Value
End Property

Property Get DoNotAllowClose() As Boolean
     DoNotAllowClose = NoClose
End Property

Property Get TargetFraction() As Double
     TargetFraction = TargFrac
End Property

Property Let NumberOfSteps(Value As Long)
    If Value < 1 Then
        NumSteps = 1
    Else
        NumSteps = Value
    End If
End Property

Property Get NumberOfSteps() As Long
    NumberOfSteps = NumSteps
End Property


Private Sub UserForm_Initialize()
   IssueMessage = True

   Me.CancelRunButton.Tag = 0
   Me.Label1.width = 0
   Me.Label2.Caption = "Initializing ..."
   NumSteps = 1
   TargFrac = 1
   NoClose = True
   ' Place the progress ProgressMeter squarely in the middle of the application
   Me.Top = (2 * Application.Top - Application.height) / 2 + Me.height / 2
   Me.Left = (2 * Application.Left - Application.width) / 2 + Me.width / 2
   
   DoEvents
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
       If NoClose Then
            MsgBox "You must press ""CANCEL RUN"" to stop the run."
            Cancel = True
       Else
            IssueMessage = False
            CancelRunButton_Click
       End If
End Sub

Property Get CurrentFraction() As Double
    CurrentFraction = CurFrac
End Property

Public Function CancelPressed() As Boolean
    If Me.CancelRunButton.Tag = 1 Then
        CancelPressed = True
    Else
        CancelPressed = False
    End If
End Function

Public Sub Step(Message As String, Optional TargetFrac As Double = -1, Optional NumberSteps As Long = -1)

    Dim Ans As Long
    Dim PMFrac0 As Double
    Dim PMFrac1 As Double
    Dim TotalNumType As Long
    
    If TargetFrac <> -1 Then
       Me.TargetFraction = TargetFrac
    End If
    
    If NumberSteps <> -1 Then
       Me.NumberOfSteps = NumberSteps
    End If
    
    CurFrac = (CurFrac + (TargFrac - StartFrac) * (1 / NumSteps))
    
    If CurFrac > TargFrac Then
        CurFrac = TargFrac
    End If
    
    
    Me.Label1.width = _
         Round(TargetWidth * CurFrac)
         
    Me.Label2.Caption = Message
    
    Me.Repaint
    DoEvents
    
End Sub


