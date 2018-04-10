Attribute VB_Name = "mdlTest"
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

Public Sub TestQuadraticFit()

Dim x1 As Double
Dim x2 As Double
Dim x3 As Double
Dim y1 As Double
Dim y2 As Double
Dim y3 As Double
Dim A As Double
Dim B As Double
Dim c As Double
'
x1 = 1
x2 = 5
x3 = 3

y1 = 1
y2 = 1
y3 = 3

mdlMath.FitParabolaTo3Points x1, y1, x2, y2, x3, y3, A, B, c

End Sub

Public Sub Test_MCMD()

MultiConfigurationMembraneDistillationModel

End Sub

Public Sub Test_AdjustInterfaceTemperaturesIfInvalid()
    ' Case 1 - no changes
    Dim T() As Double
    ReDim T(1 To 4)
    
    T(1) = 1
    T(2) = 0.1
    T(3) = 0.15
    T(4) = 0
    
    AdjustInterfaceTemperaturesIfInvalid T(1), T(4), T(2), T(3)
End Sub





Public Sub TestNewtonsMethod()

Dim Test As clsTestNewtonMethod
Dim Xk() As Double
Dim gk() As Double
Dim CV As clsControlVolumePair
Dim SysEq As clsSystemEquations

ReDim Xk(1 To 3)

Xk(1) = 2
Xk(2) = 4
Xk(3) = 10

Set Test = New clsTestNewtonMethod

NewtonsMethod Test, Xk, CV, SysEq, 0.000001, 25, 0.001

Test.EvaluateFunction Xk, CV, SysEq, gk

End Sub

Public Sub TestGaussianElimination()

Dim A() As Double
Dim B() As Double
Dim Xi() As Double

ReDim A(1 To 3, 1 To 3)
ReDim B(1 To 3)

A(1, 1) = 2
A(1, 2) = 3
A(1, 3) = -2

A(2, 1) = -1
A(2, 2) = 5
A(2, 3) = 9

A(3, 1) = -9
A(3, 2) = 7
A(3, 3) = 8

B(1) = 3
B(2) = 13
B(3) = 6

GaussianElimination A, B, Xi

' now let's deliberately provide a rank 2 matrix
A(1, 1) = 2
A(1, 2) = 3
A(1, 3) = -2

A(2, 1) = 4
A(2, 2) = 6
A(2, 3) = -4

A(3, 1) = -9
A(3, 2) = 7
A(3, 3) = 8

B(1) = 3
B(2) = 6
B(3) = 6

GaussianElimination A, B, Xi

End Sub

Public Function Test_MatrixRowSort()

' Test Case 1
' Just generate a 5x5 example matrix - from matlab...
'xoriginal =
'
'     1     1     0     0     1
'     1     1     0     1     0
'     1     0     1     0     1
'     0     0     0     1     0
'     0     0     0     0     1
'
'Shuffled Version'
'
'x =
'
'     0     0     0     0     1
'     0     0     0     1     0
'     1     1     0     1     0
'     1     0     1     0     1
'     1     1     0     0     1
'
'This is the solution found
'
'xsol =
'
'     1     1     0     1     0
'     1     1     0     0     1
'     1     0     1     0     1
'     0     0     0     1     0
'     0     0     0     0     1
'
'Order of the rows:
'
'per =
'
'     5     4     1     3     2
'
Dim A() As Double
Dim Out() As Double
Dim B() As Double
Dim i As Long
Dim WRTP() As Long
Dim Per As Variant
ReDim A(1 To 5, 1 To 5)
A(1, 1) = 0
A(1, 2) = 0
A(1, 3) = 0
A(1, 4) = 0
A(1, 5) = 1
A(2, 1) = 0
A(2, 2) = 0
A(2, 3) = 0
A(2, 4) = 1
A(2, 5) = 0
A(3, 1) = 1
A(3, 2) = 1
A(3, 3) = 0
A(3, 4) = 1
A(3, 5) = 0
A(4, 1) = 1
A(4, 2) = 0
A(4, 3) = 1
A(4, 4) = 0
A(4, 5) = 1
A(5, 1) = 1
A(5, 2) = 1
A(5, 3) = 0
A(5, 4) = 0
A(5, 5) = 1
ReDim WRTP(1 To 5)
WRTP(1) = 1
WRTP(2) = 2
WRTP(3) = 3
WRTP(4) = 4
WRTP(5) = 5
Per = First_SortMatrixDiagonalNonZero(A, 0, WRTP)


Dim c() As Double
ReDim c(1 To 10, 1 To 10)
c(1, 1) = 0
c(1, 2) = 0
c(1, 3) = 0
c(1, 4) = 0
c(1, 5) = 0
c(1, 6) = 1
c(1, 7) = 0
c(1, 8) = 0
c(1, 9) = 0
c(1, 10) = 0

c(2, 1) = 0
c(2, 2) = 0
c(2, 3) = 1
c(2, 4) = 0
c(2, 5) = 0
c(2, 6) = 0
c(2, 7) = 0
c(2, 8) = 0
c(2, 9) = 0
c(2, 10) = 0

c(3, 1) = 0
c(3, 2) = 1
c(3, 3) = 0
c(3, 4) = 0
c(3, 5) = 1
c(3, 6) = 0
c(3, 7) = 0
c(3, 8) = 0
c(3, 9) = 1
c(3, 10) = 0

c(4, 1) = 0
c(4, 2) = 0
c(4, 3) = 0
c(4, 4) = 1
c(4, 5) = 1
c(4, 6) = 0
c(4, 7) = 0
c(4, 8) = 1
c(4, 9) = 0
c(4, 10) = 0

c(5, 1) = 0
c(5, 2) = 0
c(5, 3) = 1
c(5, 4) = 0
c(5, 5) = 0
c(5, 6) = 0
c(5, 7) = 1
c(5, 8) = 0
c(5, 9) = 0
c(5, 10) = 1

c(6, 1) = 1
c(6, 2) = 0
c(6, 3) = 0
c(6, 4) = 0
c(6, 5) = 1
c(6, 6) = 0
c(6, 7) = 1
c(6, 8) = 0
c(6, 9) = 0
c(6, 10) = 1

c(7, 1) = 0
c(7, 2) = 0
c(7, 3) = 1
c(7, 4) = 1
c(7, 5) = 0
c(7, 6) = 1
c(7, 7) = 1
c(7, 8) = 0
c(7, 9) = 0
c(7, 10) = 0

c(8, 1) = 1
c(8, 2) = 0
c(8, 3) = 1
c(8, 4) = 1
c(8, 5) = 0
c(8, 6) = 0
c(8, 7) = 0
c(8, 8) = 0
c(8, 9) = 1
c(8, 10) = 0

c(9, 1) = 0
c(9, 2) = 1
c(9, 3) = 1
c(9, 4) = 1
c(9, 5) = 1
c(9, 6) = 0
c(9, 7) = 0
c(9, 8) = 0
c(9, 9) = 0
c(9, 10) = 0

c(10, 1) = 1
c(10, 2) = 0
c(10, 3) = 1
c(10, 4) = 1
c(10, 5) = 0
c(10, 6) = 1
c(10, 7) = 1
c(10, 8) = 0
c(10, 9) = 1
c(10, 10) = 0

ReDim WRTP(1 To 10)
For i = 1 To 10
  WRTP(i) = i
Next

Per = First_SortMatrixDiagonalNonZero(c, 0, WRTP)
' Solution must be 6 3 2 8 7 10 4 1 5 9


CreateRandomArrayWithZerosAndScrambledNonZeroDiagonal 0.75, 100, 100, 1, 1, 10, -10, A
ReDim B(1 To 100)

For i = 1 To 100
   B(i) = i
Next

Per = SortMatrixDiagonalNonZero(A, Out, B)

For i = 1 To 100
   If Out(i, i) = 0 Then
      MsgBox "The Matrix Row Sort Failed!!!!"
   End If
Next

' Deliverately give it a function that does not have a solution

ReDim c(1 To 10, 1 To 10)
ReDim B(1 To 10)
c(1, 1) = 0
c(1, 2) = 0
c(1, 3) = 0
c(1, 4) = 0
c(1, 5) = 0
c(1, 6) = 1
c(1, 7) = 0
c(1, 8) = 0
c(1, 9) = 0
c(1, 10) = 0

c(2, 1) = 0
c(2, 2) = 0
c(2, 3) = 1
c(2, 4) = 0
c(2, 5) = 0
c(2, 6) = 0
c(2, 7) = 0
c(2, 8) = 0
c(2, 9) = 0
c(2, 10) = 0

c(3, 1) = 0
c(3, 2) = 1
c(3, 3) = 0
c(3, 4) = 0
c(3, 5) = 1
c(3, 6) = 0
c(3, 7) = 0
c(3, 8) = 0
c(3, 9) = 1
c(3, 10) = 0

c(4, 1) = 0
c(4, 2) = 0
c(4, 3) = 0
c(4, 4) = 1
c(4, 5) = 1
c(4, 6) = 0
c(4, 7) = 0
c(4, 8) = 1
c(4, 9) = 0
c(4, 10) = 0

c(5, 1) = 0
c(5, 2) = 0
c(5, 3) = 1
c(5, 4) = 0
c(5, 5) = 0
c(5, 6) = 0
c(5, 7) = 1
c(5, 8) = 0
c(5, 9) = 0
c(5, 10) = 1

c(6, 1) = 1
c(6, 2) = 0
c(6, 3) = 0
c(6, 4) = 0
c(6, 5) = 1
c(6, 6) = 0
c(6, 7) = 1
c(6, 8) = 0
c(6, 9) = 0
c(6, 10) = 1

c(7, 1) = 0
c(7, 2) = 0
c(7, 3) = 1
c(7, 4) = 1
c(7, 5) = 0
c(7, 6) = 1
c(7, 7) = 1
c(7, 8) = 0
c(7, 9) = 0
c(7, 10) = 0

c(8, 1) = 1
c(8, 2) = 0
c(8, 3) = 1
c(8, 4) = 1
c(8, 5) = 0
c(8, 6) = 0
c(8, 7) = 0
c(8, 8) = 0
c(8, 9) = 1
c(8, 10) = 0

c(9, 1) = 0
c(9, 2) = 1
c(9, 3) = 1
c(9, 4) = 1
c(9, 5) = 1
c(9, 6) = 0
c(9, 7) = 0
c(9, 8) = 0
c(9, 9) = 0
c(9, 10) = 0

c(10, 1) = 0
c(10, 2) = 0
c(10, 3) = 0
c(10, 4) = 0
c(10, 5) = 0
c(10, 6) = 0
c(10, 7) = 0
c(10, 8) = 0
c(10, 9) = 0
c(10, 10) = 0

ReDim WRTP(1 To 10)
For i = 1 To 10
  WRTP(i) = i
Next

Per = SortMatrixDiagonalNonZero(c, Out, B)
' Solution must be 6 3 2 8 7 10 4 1 5 9




End Function

Public Sub TestHeatExchanger()

NTU = 2
Cstar = 1

EffectivenessHeatTransferRate


End Sub


Public Function TestLongRecieve(n As Long) As Single()
   Dim arr() As Single
   
   ReDim arr(0 To n)
   
   For i = 0 To n
      arr(i) = Rnd(1)
   Next
   
   TestLongRecieve = arr
End Function

Public Sub Test_LongRecieve()

Dim n As Long
Dim nn() As Single
n = 10

nn = TestLongRecieve(n)


End Sub

Public Function TestConjugateGradient()

Dim Func As clsTestCojugateGradientMethod

Set Func = New clsTestCojugateGradientMethod

Dim SysEqn As clsSystemEquations


Dim Result() As Double
Dim Max() As Double
Dim Min() As Double
Dim CV As clsControlVolumePair

ReDim Result(1 To 2)
ReDim Min(1 To 2)
ReDim Max(1 To 2)
' First guess
Result(1) = 2
Result(2) = 1

Min(1) = 0
Min(2) = 0

Max(1) = 10
Max(2) = 10

' Answer is 0, 0


' This initializes The Inputs (clsInput, SysEqns.Inputs), the Connectivity (clsConnectivity, SysEqns.Connectivity),
' and an array of control volume pairs (hot/cold) (clsControlVolumePair, Private SysEqns.CVP)

mdlMath.ConjugateGradient Func, Result, CV, SysEqn, 0.000000000001, 1000, Max, Min


End Function

Public Sub TestDetermineFlowPaths()

Dim Con As clsConnectivity

Set Con = New clsConnectivity

Con.DetermineFlowPaths


End Sub

Public Sub TestMaxHeatFlow()
Dim cph As Double
Dim cpc As Double
Dim mMD As Double
Dim T_hot As Double
Dim T_cold As Double
Dim S_hot As Double
Dim S_cold As Double
Dim m_hot As Double
Dim m_cold As Double
Dim Qmax As Double
Dim Lavg As Double
Dim T_hot_out As Double
Dim T_cold_out As Double

T_hot = 320.15
T_cold = 293.15
S_hot = 0.05
S_cold = 0
m_hot = 0.1
m_cold = 0.1

mMD = 0.000953

cph = mdlProperties.SeaWaterSpecificHeat(T_hot, S_hot)
cpc = mdlProperties.SeaWaterSpecificHeat(T_cold, S_cold)
Lavg = mdlProperties.LatentHeatOfPureWater((T_hot + T_cold) / 2)

mdlTransferCoefficient.EstimateOutputTemperatures T_hot, T_cold, m_hot, m_cold, S_hot, S_cold, mMD, cpc, cph, Lavg, T_hot_out, T_cold_out

End Sub

Public Sub TestCubicRootFunction()

Dim i As Long
Dim A As Double
Dim B As Double
Dim c As Double
Dim d As Double
Dim x As Double

Const Tol = 0.000001

For i = 1 To 1000
   A = i * (Rnd() - 0.5)
   B = i * (Rnd() - 0.5)
   c = i * (Rnd() - 0.5)
   d = i * (Rnd() - 0.5)
   
   x = mdlMath.SmallestPositiveRealCubicRoot(A, B, c, d, i * Tol)
   
   Sheet9.Range("A1").Offset(i - 1, 1) = A
   Sheet9.Range("A1").Offset(i - 1, 2) = B
   Sheet9.Range("A1").Offset(i - 1, 3) = c
   Sheet9.Range("A1").Offset(i - 1, 4) = d
   Sheet9.Range("A1").Offset(i - 1, 5) = x
Next


End Sub
