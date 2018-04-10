Attribute VB_Name = "mdlError"
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
        
        
Public Function NoError() As Boolean
    If Len(err.Source) <> 0 Then
        NoError = False
    Else
        NoError = True
    End If
End Function

Public Function NameError(Address As String, ArgName As String, ValidNames() As String)
         
        NameError = mdlConstants.glbINVALID_VALUE
         
        err.Source = Address & ": A valid name for the """ & ArgName & _
                         """ must be made.  Valid names are:"
        For i = 0 To UBound(ValidNames)
            err.Source = err.Source & vbCrLf & i & ". " & ValidNames(i)
        Next i

End Function

Public Function ReturnError(Optional PreMessage As String, Optional PostMessage As String, Optional IncludeMsgBox As Boolean = False, Optional EndOnError As Boolean = True)

        If Len(err.Source) = 0 Then
           If Not IsMissing(PreMessage) Then
              If Len(PreMessage) <> 0 Then
                  err.Source = PreMessage
              End If
           End If
           If Not IsMissing(PostMessage) Then
              If Len(PostMessage) <> 0 Then
                 err.Source = err.Source & vbCrLf & vbCrLf & PostMessage
              End If
           End If
        Else
           If Not IsMissing(PreMessage) Then
              If Len(PreMessage) <> 0 Then
                  err.Source = PreMessage & vbCrLf & vbCrLf & err.Source
              End If
           End If
           If Not IsMissing(PostMessage) Then
              If Len(PostMessage) <> 0 Then
                 err.Source = err.Source & vbCrLf & vbCrLf & PostMessage
              End If
           End If
        End If
        ' There is both a global and a local switch for whether to issue a message box
        If IncludeMsgBox And mdlConstants.glbIncludeMsgBox Then
           MsgBox err.Source, vbCritical, "Error"
           MsgBox "This model is trying to solve a set of nonlinear equations that may not be robust if you are exploring" & _
                  " a new realm of interest. The membrane distillation and math algorithms probably do not have an error unless you are doing new development." & _
                  " You are probably applying the model to a difficult or unphysical situation to solve. Sometimes increasing the " & _
                  "discretizations on the ""Configuration"" Sheet can help the model obtain a solution. Another possibility is " & _
                  "setting different values for the ""glbMaxHeatFlowFractionForInitialCondition"" or """ & _
                  "glbNewtonReductionFactor"" constants.  Be sure to note the values they have since changing them may cause other " & _
                  " solutions to stop working! Finally, you can set a break point at the line ""Out.HeatExchangerPerformanceCurves Func""" & _
                  " in mdlMath.NewtonsMethod, set ""glbTroubleshoot"" to ""True"" and watch the solution iterations in the ""Output"" sheet" & _
                  " this often makes it clear concerning what Newton's method is doing and why the solution is not converging."
           ' There is both a global and a local switch for ending the program
           If mdlConstants.glbEndOnError And EndOnError Then
              End
           End If
        End If
        
        'This is just here to allow the assignment of the invalid value with a single line when using this command.
        ReturnError = mdlConstants.glbINVALID_VALUE
        
End Function
