Attribute VB_Name = "mdlMain"
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
Option Base 1


Public Sub MultiConfigurationMembraneDistillationModel(Optional IncludeCustomInputSubRoutine As Boolean = False, _
                                                       Optional CustomSubroutineName As String = "", _
                                                       Optional SubCaseValue As Variant)
' Many different models like this can be configured using the clsSystemEquations class which is
' the master class for the entire simulation.
Dim Inp As clsInput
Dim RunSequenceOfLossConfigurations() As Long
Dim HotFlowMult() As Double
Dim ColdFlowMult() As Double
Dim i As Long
Dim SysEqns() As clsSystemEquations
Dim NumRun As Long
Dim Out As clsOutput
Dim Dum As Double

'Make sure that you have the proper workbook active
'Otherwise, the code is full of references and ranges that will fail!
ThisWorkbook.Activate
' Make sure that these troubleshooting mechanisms get reset just
' in case a run failed.
Range("IsFinalRunRange").Value = "FALSE"
Range("IsFinalNewtonIterationRange").Value = "FALSE"
mdlValidation.WriteDebugInfoToFile "", True ' For this first case, delete the debug file for a fresh run

' Get all of the input (this is redundantly stored in SysEq)
Set Inp = New clsInput
' Figure out how many runs have to be performed of the various layers (this is getting closer to what you would do for a Spiral Configuration that
' has lateral heat transfer but for now we can only solve for losses that are not added to other nodes.
If mdlConstants.glbUseIndependentLayers Then
    FigureOutRunSequenceIndependentLayers Inp, RunSequenceOfLossConfigurations, HotFlowMult, ColdFlowMult
Else
' THIS IS THE CORRECT METHOD FOR CSM LAYERS 3 membranes = 2 Layers!!!
    FigureOutRunSequenceSharingLayers Inp, RunSequenceOfLossConfigurations, HotFlowMult, ColdFlowMult
End If
' Perform all of the required runs
NumRun = UBound(RunSequenceOfLossConfigurations)
ReDim SysEqns(1 To NumRun)

For i = 1 To NumRun
   RunEngine RunSequenceOfLossConfigurations(i), SysEqns(i), HotFlowMult(i), _
             ColdFlowMult(i), glbEvaluateEquationsAtEnd, i, IncludeCustomInputSubRoutine, _
             CustomSubroutineName, SubCaseValue
   'SysEqns(i).ControlVolumePair(1).CalculationComplete = False
   'SysEqns(i).ControlVolumePair(1).CalculateControlVolumePair (SysEqns(i))
Next i

' Process the output
Set Out = New clsOutput

Out.CalculateOutput SysEqns, RunSequenceOfLossConfigurations

End Sub

'THIS IS THE STANDARD CODE COMPATIBLE WITH CSM's EXPERIMENTS!
Private Sub FigureOutRunSequenceSharingLayers(Inp As clsInput, RunSequenceOfLossConfigurations() As Long, HotFlowMult() As Double, ColdFlowMult() As Double)
'0 = No losses, 1 = Losses on the hot side, 2 = losses on the cold side, 3 Losses on both sides
' This function has to be synced with clsOutput.DetermineLayerMultiplicationFactor
' The first value in RunSequenceOfLossConfigurations is always the value for which temperatures will
' be read.
If Inp.IncludeExternalLosses Then
   If Inp.NumberOfLayers = 1 Then
        ReDim HotFlowMult(1)
        ReDim ColdFlowMult(1)
        HotFlowMult(1) = 1
        ColdFlowMult(1) = 1
   
        If Inp.NumberOfExposedSides = 2 Then
             ReDim RunSequenceOfLossConfigurations(1)
             RunSequenceOfLossConfigurations(1) = 3
        ElseIf Inp.NumberOfExposedSides = 1 Then
             ReDim RunSequenceOfLossConfigurations(1)
             
             If Inp.ExternalLossIsHotSide Then
                RunSequenceOfLossConfigurations(1) = 1
             Else
                RunSequenceOfLossConfigurations(1) = 2
             End If
        Else
             MsgBox "This should never happen!!!"
        End If
   ElseIf Inp.NumberOfLayers > 1 Then
       ReDim RunSequenceOfLossConfigurations(1 To 3)
       ReDim HotFlowMult(1 To 3)
       ReDim ColdFlowMult(1 To 3)
       HotFlowMult(1) = 0.5
       HotFlowMult(2) = 0.5
       HotFlowMult(3) = 1
       ColdFlowMult(1) = 0.5
       ColdFlowMult(2) = 1
       ColdFlowMult(3) = 0.5
       If Inp.NumberOfExposedSides = 1 Then

            If Inp.ExternalLossIsHotSide Then
                 RunSequenceOfLossConfigurations(1) = 0  'Symmetric
                 RunSequenceOfLossConfigurations(2) = 0  'Cold side no losses
                 RunSequenceOfLossConfigurations(3) = 1  'Hot side with losses

            Else
                 RunSequenceOfLossConfigurations(1) = 0 'Symmetric
                 RunSequenceOfLossConfigurations(2) = 2 'Cold side with losses
                 RunSequenceOfLossConfigurations(3) = 0 'Hot side no losses
            End If
        ElseIf Inp.NumberOfExposedSides = 2 Then ' We have to run three times! ugh!
        
            RunSequenceOfLossConfigurations(1) = 0
            RunSequenceOfLossConfigurations(2) = 1
            RunSequenceOfLossConfigurations(3) = 2

        Else 'ERROR
           MsgBox "This should never happen!!"
        End If
   Else
       Dim Dum As Double
       Dum = mdlError.ReturnError("mdlMain.FigureOutRunSequence: This should never happen! there is a flaw in the input structure.  There is a case that is not covered for the run sequence!", , True)
   End If
Else
   If Inp.NumberOfLayers = 1 Then
        ReDim RunSequenceOfLossConfigurations(1 To 1)
        RunSequenceOfLossConfigurations(1) = 0
        ReDim HotFlowMult(1 To 1)
        ReDim ColdFlowMult(1 To 1)
        HotFlowMult(1) = 1
        ColdFlowMult(1) = 1
   Else
        ReDim RunSequenceOfLossConfigurations(1 To 3)
        RunSequenceOfLossConfigurations(1) = 0
        RunSequenceOfLossConfigurations(2) = 0
        RunSequenceOfLossConfigurations(3) = 0
        ReDim HotFlowMult(1 To 3)
        ReDim ColdFlowMult(1 To 3)
        HotFlowMult(1) = 0.5
        HotFlowMult(2) = 0.5
        HotFlowMult(3) = 1
        ColdFlowMult(1) = 0.5
        ColdFlowMult(2) = 1
        ColdFlowMult(3) = 0.5
    End If
End If

End Sub


Private Sub FigureOutRunSequenceIndependentLayers(Inp As clsInput, RunSequenceOfLossConfigurations() As Long, HotFlowMult() As Double, ColdFlowMult() As Double)
'0 = No losses, 1 = Losses on the hot side, 2 = losses on the cold side, 3 Losses on both sides
' This function has to be synced with clsOutput.DetermineLayerMultiplicationFactor
' The first value in RunSequenceOfLossConfigurations is always the value for which temperatures will
' be read.
Dim i As Long

If Inp.IncludeExternalLosses Then
   If Inp.NumberOfLayers = 1 And Inp.NumberOfExposedSides = 2 Then
        ReDim RunSequenceOfLossConfigurations(1)
        RunSequenceOfLossConfigurations(1) = 3
   ElseIf Inp.NumberOfLayers = 1 And Inp.NumberOfExposedSides = 1 Then
        ReDim RunSequenceOfLossConfigurations(1)
        If Inp.ExternalLossIsHotSide Then
           RunSequenceOfLossConfigurations(1) = 1
        Else
           RunSequenceOfLossConfigurations(1) = 2
        End If
   ElseIf Inp.NumberOfLayers > 1 And Inp.NumberOfExposedSides = 1 Then
       ReDim RunSequenceOfLossConfigurations(1 To 2)
       If Inp.ExternalLossIsHotSide Then
            RunSequenceOfLossConfigurations(1) = 0
            RunSequenceOfLossConfigurations(2) = 1
       Else
            RunSequenceOfLossConfigurations(1) = 0
            RunSequenceOfLossConfigurations(2) = 2
       End If
   ElseIf Inp.NumberOfLayers = 2 And Inp.NumberOfExposedSides = 2 Then
        ReDim RunSequenceOfLossConfigurations(1 To 2)
        RunSequenceOfLossConfigurations(1) = 1
        RunSequenceOfLossConfigurations(2) = 2
   ElseIf Inp.NumberOfLayers > 2 And Inp.NumberOfExposedSides = 2 Then ' We have to run three times! ugh!
        ReDim RunSequenceOfLossConfigurations(1 To 3)
        RunSequenceOfLossConfigurations(1) = 0
        RunSequenceOfLossConfigurations(2) = 1
        RunSequenceOfLossConfigurations(3) = 2
   Else
       Dim Dum As Double
       Dum = mdlError.ReturnError("mdlMain.FigureOutRunSequence: This should never happen! there is a flaw in the input structure.  There is a case that is not covered for the run sequence!", , True)
   End If
Else
   ReDim RunSequenceOfLossConfigurations(1)
   RunSequenceOfLossConfigurations(1) = 0
End If

ReDim HotFlowMult(1 To UBound(RunSequenceOfLossConfigurations))
ReDim ColdFlowMult(1 To UBound(RunSequenceOfLossConfigurations))
For i = 1 To UBound(RunSequenceOfLossConfigurations)
   HotFlowMult(i) = 1
   ColdFlowMult(i) = 1
Next

End Sub


Public Sub RunEngine(LossesConfiguration As Long, SysEqn As clsSystemEquations, _
                     HotFlowMult As Double, ColdFlowMult As Double, _
                     Optional EvaluateConvergedSolutionAtEnd As Boolean = False, _
                     Optional RunNum As Long = 0, _
                     Optional IncludeCustomInputSubRoutine As Boolean = False, _
                     Optional CustomSubroutineName As String = "", _
                     Optional SubCaseValue As Variant)

Dim Xk() As Double ' Solution vector for overall system - if n is the number of nodes, then 1 To n = nodal temperatures,
'                                                         n+1 to 2n are nodal mass flows, and 2n+1 to 3n are nodal salinities
Dim XkMax() As Double ' maximum limits for the solution. any solution outside of these limits is invalid
Dim XkMin() As Double ' minimum limits for the solution. any solution outside of these limits is invalid
Dim GlobalAbsConvergCrit() As Double

Dim CV As clsControlVolumePair 'this is not used here
Dim Out As clsOutput
Dim Wksh As Worksheet
Dim DebugString As String

' Develop connectivity of the control volumes and node numbers

Set SysEqn = New clsSystemEquations
SysEqn.LossesConfiguration = LossesConfiguration
SysEqn.SetMassFlowMultipliers HotFlowMult, ColdFlowMult
' 01/02/2018 - This method now includes a custom subroutine to change the input through a parameter study!
SysEqn.InitlializeSystemEquations IncludeCustomInputSubRoutine, CustomSubroutineName, SubCaseValue


ReDim Xk(1 To SysEqn.NumberEquations)

'Assign the initial variables
SysEqn.GetVariables Xk

' Now write the inputs to the debug output file:
SysEqn.Inputs.DebugInputs

' ''''''''''''''''''''''''''''''''
' Solve the overall set of equations
' ''''''''''''''''''''''''''''''''

' Get maximum variable values
SysEqn.ReturnMaximumAndMinimumForVariables XkMax, XkMin

' Build the absolute convergence criteria.  These are necessary because there is
' numeric noise in the equations at different orders of magnitude for different physical entities
' (energy, mass flow, salinity) that cause
' the system to not converge in a relative sense even though errors are in the noise for the
' physical variable under consideration.
mdlConstants.GlobalEquationSetConvergenceCriterion SysEqn, GlobalAbsConvergCrit

' Newton's method works conjugate gradients just doesn't seem to want to work for the model solution.
NewtonsMethod SysEqn, Xk, CV, SysEqn, mdlConstants.glbNewtonConvergenceCriterion, _
              mdlConstants.glbMaxNewtonIterations, mdlConstants.glbDerivativeIncrement, _
              XkMax, XkMin, GlobalAbsConvergCrit

' SysEqn now has all of the solution results which can be processed by clsOutput.
If EvaluateConvergedSolutionAtEnd Then
    ' This is used to set break points at locations where you are interested
    
    Range("IsFinalRunRange").Value = "TRUE"
    
    
    DebugString = ""
    DebugString = DebugString & ""
    DebugString = DebugString & " **********************************************" & vbCrLf
    DebugString = DebugString & " ****              RUN NUMBER " & RunNum & "            ****" & vbCrLf
    DebugString = DebugString & " **********************************************" & vbCrLf
    Debug.Print DebugString
    mdlValidation.WriteDebugInfoToFile DebugString
    
    
    Dim Rk() As Double
    SysEqn.GetVariables Xk
    ' Reset the calculation complete flag for CVP(1)
    SysEqn.SetControlVolumePairsToNotCalulated
    SysEqn.EvaluateFunction Xk, SysEqn.ControlVolumePair(1), SysEqn, Rk
    
    Range("IsFinalRunRange").Value = "FALSE"
    Range("IsFinalNewtonIterationRange").Value = "FALSE"
End If
    
End Sub

