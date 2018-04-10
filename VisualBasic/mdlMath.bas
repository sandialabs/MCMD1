Attribute VB_Name = "mdlMath"
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

Public Function AbsDbl(d As Double) As Double
   If d < 0 Then
      AbsDbl = -d
   Else
      AbsDbl = d
   End If
End Function

Public Function SmallestPositiveRealCubicRoot(A As Double, B As Double, c As Double, d As Double, Tol As Double) As Double
' This function has undergone thorough testing against matlab's root function.
' it can easily be transformed to return other roots including imaginary roots.

' You have to be careful to not set the tol (tolerance) too tight or you can cause the function to reject valid positive roots.
' The function returns glbINVALID_VALUE if no real positive root exists for the polynomial of interest.

' This subroutine returns the smallest positive real root of the cubic polynomial
' If the only real root is negative this function returns glbINVALID_VALUE
' a*x^3 + b*x^2 + c*x + d = 0
' it does not return complex roots
Dim del As Double 'discrimenant of the equation
Dim del1 As Double
Dim del0 As Double
Dim ReC As Double 'real part of C
Dim ImC As Double 'imaginary part of C
Dim r As Double 'radius in the complex plane (polar coordinates)
Dim r1_3 As Double
Dim theta As Double ' angle in the complex plane (polar coordinates)
Dim x() As Double
Dim xtriple As Double
Dim xdouble As Double
Dim xsimple As Double
Dim xPosMin As Double
Dim i As Long
Dim err As Double
Dim yy As Double
Dim xx As Double
Dim tol3 As Double

If Abs(A) < Tol And Abs(B) < Tol And Abs(c) < Tol Then ' This is an absurdity, throw an error
    SmallestPositiveRealCubicRoot = mdlError.ReturnError("mdlMath.SmallestPositiveRealCubicRoot: The first three coefficients of the polynomial are zero and the equation is d = 0 which is only true if d is nonzero.", , True)
ElseIf Abs(A) < Tol And Abs(B) < Tol Then ' solve a linear equation
    If -d / c > 0 Then ' c is already verified to not be close to zero!
        SmallestPositiveRealCubicRoot = -d / c
    Else
        SmallestPositiveRealCubicRoot = mdlConstants.glbINVALID_VALUE
    End If
ElseIf Abs(A) < Tol Then ' we solve a quadratic equation b*x^2 + c*x + d = 0
    ReDim x(0 To 1)
    del = c ^ 2 - 4 * B * d
    If del < 0 Then 'There are no real positive roots, return an invalid value
       SmallestPositiveRealCubicRoot = mdlConstants.glbINVALID_VALUE
    Else
       x(0) = (-c + Sqr(del)) / (2 * B)
       x(1) = (-c - Sqr(del)) / (2 * B)
       xPosMin = 1E+20
       For i = 0 To 1
            If xPosMin > x(i) And x(i) >= 0 Then
                xPosMin = x(i)
            End If
       Next
       If xPosMin = 1E+20 Then
          SmallestPositiveRealCubicRoot = mdlConstants.glbINVALID_VALUE
       Else
          SmallestPositiveRealCubicRoot = xPosMin
       End If
    End If
ElseIf Abs(B) < Tol And Abs(c) < Tol Then ' We can easily solve the cubic A > tol
    If (-d / A) >= 0 Then
       SmallestPositiveRealCubicRoot = (-d / A) ^ (1 / 3)
    Else
       SmallestPositiveRealCubicRoot = mdlConstants.glbINVALID_VALUE
    End If
Else ' We are dealing with the general case where all coefficients are non-zero
    

    ' Test for first condition:
    ' MsgBox SmallestPositiveRealCubicRoot(3,3,1,(1/9),.0000000001)
    ' Test for second condition:
    ' MsgBox SmallestPositiveRealCubicRoot(3,3,-1,-1.18409491661026,.0000000001)
    
    del = 18 * A * B * c * d - 4 * B ^ 3 * d + B ^ 2 * c ^ 2 - 4 * A * c ^ 3 - 27 * A ^ 2 * d ^ 2
    del1 = 2 * B ^ 3 - 9 * A * B * c + 27 * A ^ 2 * d
    del0 = B ^ 2 - 3 * A * c

    tol3 = Tol ^ 3
    If Abs(del) < tol3 And Abs(del0) < tol3 Then ' Case that will cause the general algorithm to divide by zero
    ' a triple root exists, is it positive?
        
        xtriple = -B / (3 * A)
        If xtriple <= 0 Then
           SmallestPositiveRealCubicRoot = mdlConstants.glbINVALID_VALUE
        Else
           SmallestPositiveRealCubicRoot = xtriple
        End If
    ElseIf Abs(del) < tol3 Then 'another case that the general algorithm cannot handle
    ' a double root and a simple root exist.
       If A = 0 Then
           ' We just have a quadratic equation
           ' But we have already handled this situation
       Else
            xdouble = (9 * A * d - B * c) / (2 * del0)
            xsimple = (4 * A * B * c - 9 * A ^ 2 * d - B ^ 3) / (A * del0)
            
            If xdouble < 0 And xsimple < 0 Then
               SmallestPositiveRealCubicRoot = mdlConstants.glbINVALID_VALUE
            ElseIf xdouble < 0 Then
               SmallestPositiveRealCubicRoot = xsimple
            ElseIf xsimple < 0 Then
               SmallestPositiveRealCubicRoot = xdouble
            Else ' both roots are positive or zero
               If xdouble < xsimple Then
                  SmallestPositiveRealCubicRoot = xdouble
               Else
                  SmallestPositiveRealCubicRoot = xsimple
               End If
            End If
       End If
      
    Else ' Use the general algorithm
       
        ' Decompose the Cubic root to polar coordinates
        If del1 ^ 2 > 4 * del0 ^ 3 Then
            r = (del1 + (del1 ^ 2 - 4 * del0 ^ 3) ^ 0.5) / 2
            If r < 0 Then
               r = -r
               theta = glbPi
            Else
               theta = 0
            End If
        Else
            ' I wish there was an atan2 function in VBA that discerns the quadrants.
            ' for this case only the 1st and 2nd quadrants are possible.
            r = Sqr((del1 / 2) ^ 2 + mdlMath.AbsDbl(del1 ^ 2 - 4 * del0 ^ 3) / 4) ' square root and square functions cancel here but result must be positive!
            xx = del1
            yy = (-del1 ^ 2 + 4 * del0 ^ 3) ^ 0.5
            
            theta = Atn(yy / xx) 'minus sign changed to avoid a negative square root
            If xx < 0 Then 'yy here is ALWAYS positive
               theta = theta + glbPi
            End If

        End If
        
        ReDim x(0 To 2)
        
        r1_3 = r ^ (1 / 3)
        For i = 0 To 2
            ReC = r1_3 * Cos((theta + 2 * i * glbPi) / 3)
            ImC = r1_3 * Sin((theta + 2 * i * glbPi) / 3)
            x(i) = -1 / (3 * A) * (B + ReC + del0 * (ReC / (ReC ^ 2 + ImC ^ 2)))
            ' Now check to see if this root actually works
            If Abs(A * x(i) ^ 3 + B * x(i) ^ 2 + c * x(i) + d) > Tol Then
               x(i) = 1E+20 'This eliminates this room in the minimum evaluation
            End If
        Next
        
        xPosMin = 1E+20
        
        For i = 0 To 2
           If xPosMin > x(i) And x(i) >= 0 Then
               xPosMin = x(i)
           End If
        Next
        If xPosMin = 1E+20 Then
           SmallestPositiveRealCubicRoot = mdlConstants.glbINVALID_VALUE
        Else
           SmallestPositiveRealCubicRoot = xPosMin
        End If
    End If
End If
' Test diagnostic
' Comment out after this function has been thoroughly tested or leave it in if you want to keep the function safe from
' careless changes that cause it to fail.
xPosMin = SmallestPositiveRealCubicRoot
err = A * xPosMin ^ 3 + B * xPosMin ^ 2 + c * xPosMin + d
If Abs(err) > Tol And xPosMin <> mdlConstants.glbINVALID_VALUE Then
   SmallestPositiveRealCubicRoot = mdlError.ReturnError("mdlMath.SmallestPositiveRealCubicRoot: The value found is NOT a root of the original polynomial equation.  There is a bug in this function!", , True)
End If

End Function

Public Function LeastCommonMultiple(A As Long, B As Long) As Long
 ' This was tested on 3/14/2017 by
   'LeastCommonMultiple(8, 9) = 72
   'LeastCommonMultiple(2 * 3 * 5 * 13, 3 * 17) = 2 * 3 * 5 * 13 * 17
   ' It came from http://www.vbforums.com/showthread.php?640902-Best-way-to-find-the-LCM-of-2-or-more-numbers
   ' accessed on 3/14/2017
    If (A < 0) Or (B < 0) Then
       mdlError.ReturnError ("mdlMath.LeastCommonMultiple: Negative integer input. Both inputs a and b must be positive integers!")
    End If
   
    LeastCommonMultiple = (A / GCD(A, B)) * B
End Function

Private Function GCD(A As Long, B As Long) As Long
    'Conventionally set a>=b
    ' comes from http://www.vbforums.com/showthread.php?640902-Best-way-to-find-the-LCM-of-2-or-more-numbers
    ' accessed on 3/14/2017
    If A < B Then
        GCD = GCD(B, A)
Exit Function
    End If
    
    If B <> 0 Then
        GCD = GCD(B, A Mod B)
Exit Function
    End If
    
    GCD = A
Exit Function
End Function

Public Function RealCeiling(ByVal x As Double, Optional ByVal Factor As Double = 1) As Double
    ' X is the value you want to round
    ' Factor is the multiple to which you want to round
    RealCeiling = (Int(x / Factor) - (x / Factor - Int(x / Factor) > 0)) * Factor
End Function

Public Function RealFloor(ByVal x As Double, Optional ByVal Factor As Double = 1) As Double
    ' X is the value you want to round
    ' Factor is the multiple to which you want to round
    RealFloor = Int(x / Factor) * Factor
End Function

Public Function Ceiling(ByVal x As Double) As Long
    ' X is the value you want to round
    ' Factor is the multiple to which you want to round
    Ceiling = (Int(x) - (x - Int(x) > 0))
End Function

Public Function Floor(ByVal x As Double) As Long
    ' X is the value you want to round
    ' Factor is the multiple to which you want to round
    Floor = Int(x)
End Function

Public Function SortMatrixDiagonalNonZero(Matrix() As Double, OutMatrix() As Double, Vector() As Double) As Variant
' This function has been thoroughly tested. It employees a back tracking algorithm that
' is a recursive formula that can be a bit difficult to understand (at least for me).

' https://en.wikipedia.org/wiki/Backtracking

Dim WhichRowsToPlace() As Long
Dim Mat0() As Double
Dim MatFinal() As Double
Dim i As Long
Dim j As Long
Dim Row As Variant
Dim Col As Variant
Dim ColSum() As Long
Dim RowSum() As Long
Dim IX As Variant
Dim IX2 As Variant
Dim IX3 As Variant
Dim Bool As Boolean
Dim TempVector() As Double
Dim Dum As Double

'More error checking may be needed but this function is only used in a limited context.

If LBound(Matrix, 1) <> 1 Or LBound(Matrix, 2) <> 1 Then
    Dum = mdlError.ReturnError("mdlMath.SortMatrixDiagonalNonZero: The input matrix must have a 1 to m, 1 to n index structure!", , True)
    GoTo FunctionError
End If

ReDim WhichRowsToPlace(1 To UBound(Matrix, 1))
ReDim TempVector(1 To UBound(Vector))
For i = 1 To UBound(Matrix, 1)
   WhichRowsToPlace(i) = i
   TempVector(i) = Vector(i)
Next i

' Create a matrix that has 1 for all non-zero entries and zero otherwise
ArrayMod_shared.ConvertToZerosAndOnes Matrix, Mat0

' Now find the summation of all rows
ReDim Row(0)
ReDim Col(0)
ReDim RowSum(1 To UBound(Matrix, 1))
ReDim ColSum(1 To UBound(Matrix, 2))
For i = 1 To UBound(Matrix, 1)
   Bool = ArrayMod_shared.GetRow(Mat0, Row, i)
   RowSum(i) = CLng(ArrayMod_shared.SumVariantVector(Row))
   If RowSum(i) = 0 Then
      Dum = mdlError.ReturnError("mdlMath.SortMatrixDiagonalNonZero: The input matrix row " & CStr(i) & " is all zeros. No solution with all diagonals nonzero exists!", , True)
   End If
Next
For i = 1 To UBound(Matrix, 2)
   Bool = ArrayMod_shared.GetColumn(Mat0, Col, i)
   ColSum(i) = CLng(ArrayMod_shared.SumVariantVector(Col))
   If ColSum(i) = 0 Then
      Dum = mdlError.ReturnError("mdlMath.SortMatrixDiagonalNonZero: The input matrix column " & CStr(i) & " is all zeros. No solution with all diagonals nonzero exists!", , True)
   End If
Next


IX = Sort(RowSum)
'Rearrange the entire matrix so that the rows with the fewest entries come first (this increase the likelihood of convergence)
ReDim MatFinal(1 To UBound(Matrix, 1), 1 To UBound(Matrix, 2))

For i = 1 To UBound(Mat0, 1)
   For j = 1 To UBound(Mat0, 1)
       MatFinal(i, j) = Mat0(IX(i), j)
   Next
Next

IX2 = First_SortMatrixDiagonalNonZero(MatFinal, 0, WhichRowsToPlace)

' Now rearrange the rows to the final matrix
ReDim IX3(1 To UBound(Matrix, 1))
ReDim OutMatrix(1 To UBound(Matrix, 1), 1 To UBound(Matrix, 2))
For i = 1 To UBound(Matrix, 1)
   IX3(IX(i)) = IX2(i)
   For j = 1 To UBound(Matrix, 2)
       OutMatrix(IX2(i), j) = Matrix(IX(i), j)
       Vector(IX2(i)) = TempVector(IX(i))
   Next
Next

' Output the transformation of original matrix to output matrix so that
' we can convert back to the original order if necessary.
SortMatrixDiagonalNonZero = IX3

FunctionEnd:

Exit Function
FunctionError:
   SortMatrixDiagonalNonZero = mdlError.ReturnError()
GoTo FunctionEnd
End Function

Public Function First_SortMatrixDiagonalNonZero(Matrix() As Double, Row As Long, WhichRowsToPlace() As Long) As Variant

' This function is copied from a matlab function from
' https://www.mathworks.com/matlabcentral/answers/125528-how-to-sort-rows-of-a-2d-array-such-that-all-elements-along-a-diagonal-are-non-zero
' The matrix here must contain beginning indices equal to 1.
'
' It uses a backtracking algorithm https://en.wikipedia.org/wiki/Backtracking

Dim nr As Long

ReDim First_SortMatrixToHaveAllDiagonalsNonZero(0)
For nr = 1 To UBound(Matrix, 1) - Row
    'Recursive Relationship between First and Next.
    First_SortMatrixDiagonalNonZero = Next_SortMatrixDiagonalNonZero(Matrix, Row + 1, nr, WhichRowsToPlace)
    If Not ArrayMod_shared.ContainsNonNumeric(First_SortMatrixDiagonalNonZero) Then 'We got ourselves a solution, we can go back up a level
Exit Function
    End If
Next

End Function

Public Function Next_SortMatrixDiagonalNonZero(Matrix() As Double, Row As Long, nr As Long, WhichRowsToPlace() As Long) As Variant

Dim firstrow As Variant
Dim Bool As Boolean
Dim nonzero() As Long
Dim i As Long
Dim Ind As Long
Dim NumColumn As Long
Dim includeInLeaves() As Long
Dim leaves() As Long
Dim NumLeaves As Long
Dim NextRowsToPlace() As Long
Dim nextSol As Variant
Dim Dum As Double
Dim TempVariant As Variant
Dim WhichRowsToPlaceBackup() As Long

NumColumn = UBound(Matrix, 2) - LBound(Matrix, 2) + 1

ReDim firstrow(0)
Bool = ArrayMod_shared.GetRow(Matrix, firstrow, Row)

'Find location of all entries not equal to zero
ReDim nonzero(0 To NumColumn - 1)
Ind = 0
For i = LBound(firstrow) To UBound(firstrow)
   If firstrow(i) <> 0 Then
      nonzero(Ind) = i
      Ind = Ind + 1
   End If
Next i
CLngArr (ArrayMod_shared.SubArray1D(nonzero, LBound(nonzero), Ind - 1)), nonzero

'find all rows that the current row can still fill and supply a non-zero diagonal
Ind = 1
ArrayMod_shared.IsMemberLng nonzero, WhichRowsToPlace, includeInLeaves
NumLeaves = CLng(ArrayMod_shared.SumVariantVector(includeInLeaves))
   
If nr > NumLeaves Then
   Next_SortMatrixDiagonalNonZero = Array("NaN")
Else
    ' formulate the leaves vector
    ReDim leaves(1 To NumLeaves)
    For i = LBound(includeInLeaves) To UBound(includeInLeaves)
       If includeInLeaves(i) = 1 Then
           leaves(Ind) = nonzero(i)
           Ind = Ind + 1
       End If
    Next i

' If the row number is greater than the number of leaves,
' then you have exhausted all possibilities and the root must therefore be bad and you have to go back.
'

   Ind = ArrayMod_shared.FindIndexLng(leaves(nr), WhichRowsToPlace)
   If Ind <> -1 Then
      CLngArr WhichRowsToPlace, WhichRowsToPlaceBackup
      Bool = ArrayMod_shared.DeleteArrayElement(WhichRowsToPlace, Ind, True)
      If ArrayMod_shared.IsArrayEmpty(WhichRowsToPlace) Then GoTo AssignValue
   Else
      ' This should never happen because the leaves have been confirmed to be a member of WhichRowsToPlace.
      ' it is here just in case VBA does something strange.
      err.Source = "mdlMath.Next_SortMatrixDiagonalNonZero: The desired leaf is NOT in the rows needing placement! Something is wrong in the coding!"
      Dum = mdlError.ReturnError()
   End If
   
   ' HERE IS A RECURSIVE RELATIONSHIP!!! This makes the code hard to understand but
   ' we work through the rows going between First and Next.  The set WhichRowsToPlace Shrinks by one for every
   '   recursive depth level.  Once you reach no more rows to place, you have found a solution and the recursive levels
   '   exit.
   nextSol = First_SortMatrixDiagonalNonZero(Matrix, Row, WhichRowsToPlace)
   
   If ArrayMod_shared.ContainsNonNumeric(nextSol) Or ArrayMod_shared.IsArrayEmpty(WhichRowsToPlace) Then
      'Revert to back matlab does this seemlessly but the ByRef passing attribute of VBA kills us.
      If ArrayMod_shared.IsArrayEmpty(WhichRowsToPlace) Then
          Ind = Ind
      End If
      CLngArr WhichRowsToPlaceBackup, WhichRowsToPlace
   End If
   
   If IsArray(nextSol) Then
      Bool = ArrayMod_shared.InsertElementIntoArray( _
                      nextSol, 1, leaves(nr))
      Next_SortMatrixDiagonalNonZero = nextSol
   Else
AssignValue:
      ReDim TempVariant(1 To 1)
      TempVariant(1) = leaves(nr)
      Next_SortMatrixDiagonalNonZero = TempVariant
   End If
      
End If

End Function

Public Function Polynomial(coef() As Double, Val As Double) As Double
    
    ' This function orders the coefficient so that the first coefficient is the constant term (i.e. x^0 term)
    ' this is the opposite of many approaches!
    
    Dim i As Long
    Dim LB As Long
    Dim Temp As Double
    
    LB = LBound(coef)
    
    Temp = 0
    
    For i = LB To UBound(coef)
        If coef(i) <> 0 Then
            Temp = Temp + coef(i) * Val ^ (i - LB)
        End If
    Next i

    Polynomial = Temp
End Function

Public Function Polynomial2D(coef() As Double, x As Double, Y As Double) As Double

    Dim i As Long
    Dim j As Long
    Dim LBx As Long
    Dim LBy As Long
    Dim Temp As Double
    
    LBx = LBound(coef, 1)
    LBy = LBound(coef, 2)
    
    Temp = 0
    For i = LBx To UBound(coef, 1)
        For j = LBy To UBound(coef, 2)
            If coef(i, j) <> 0 Then
                Temp = Temp + coef(i, j) * x ^ (i - LBx) * Y ^ (j - LBy)
            End If
        Next j
    Next i

    Polynomial2D = Temp
End Function

Public Sub MatrixMultiply(A() As Double, B() As Double, ByRef r() As Double)
' Thorough testing of this function has not been completed.
Dim i As Long
Dim j As Long
Dim k As Long
Dim Dum As Double
Dim Sum As Double

If ArrayMod_shared.NumberOfArrayDimensions(A) = 1 And ArrayMod_shared.NumberOfArrayDimensions(B) = 1 Then

    If UBound(A) - LBound(A) <> UBound(B) - LBound(B) Then
       'Error Condition - this does not work
        
        Dum = mdlError.ReturnError("MatrixMultiply: A and B must be vectors of the same length to multiply them (i.e. dot product)!")
    End If
    
    ReDim r(0 To 0)

    
    Sum = 0
    For i = LBound(A) To UBound(A)
        Sum = Sum + A(i) * B(i - LBound(A) + LBound(B))
    Next
    r(0) = Sum

ElseIf ArrayMod_shared.NumberOfArrayDimensions(A) = 1 And ArrayMod_shared.NumberOfArrayDimensions(B) = 2 Then

    If UBound(A) - LBound(A) <> UBound(B, 1) - LBound(B, 1) Then
       'Error Condition - this does not work
        
        Dum = mdlError.ReturnError("MatrixMultiply: The inner dimensions of inputs A and B are incompatible for matrix multiplication!")
    End If
    
    ReDim r(LBound(A) To UBound(A))

       
    For j = LBound(B, 2) To UBound(B, 2)
        Sum = 0
        For k = LBound(A) To UBound(A)
            Sum = Sum + A(k) * B(k - LBound(A) + LBound(B, 1), j)
        Next
        r(j - LBound(B, 2) + LBound(A)) = Sum
    Next


ElseIf ArrayMod_shared.NumberOfArrayDimensions(A) = 2 And ArrayMod_shared.NumberOfArrayDimensions(B) = 1 Then

    If UBound(A, 2) - LBound(A, 2) <> UBound(B) - LBound(B) Then
       'Error Condition - this does not work
        
        Dum = mdlError.ReturnError("MatrixMultiply: The inner dimensions of inputs A and B are incompatible for matrix multiplication!")
    End If
    
    ReDim r(LBound(B) To UBound(B))
    
    For i = LBound(A, 1) To UBound(A, 1)

            Sum = 0
            For k = LBound(A, 2) To UBound(A, 2)
                Sum = Sum + A(i, k) * B(k - LBound(A, 2) + LBound(B))
            Next
            r(i - LBound(A, 1) + LBound(B)) = Sum
    Next

ElseIf ArrayMod_shared.NumberOfArrayDimensions(A) = 2 And ArrayMod_shared.NumberOfArrayDimensions(B) = 2 Then

    If UBound(A, 2) - LBound(A, 2) <> UBound(B, 1) - LBound(B, 1) Then
       'Error Condition - this does not work
        
        Dum = mdlError.ReturnError("MatrixMultiply: The inner dimensions of inputs A and B are incompatible for matrix multiplication!")
    End If
    
    ReDim r(LBound(A, 1) To UBound(A, 1), LBound(B, 2) To UBound(B, 2))
    
    For i = LBound(A, 1) To UBound(A, 1)
        For j = LBound(B, 2) To UBound(B, 2)
            Sum = 0
            For k = LBound(A, 2) To UBound(A, 2)
                Sum = Sum + A(i, k) * B(k - LBound(A, 2) + LBound(B, 2), j)
            Next
            r(i, j) = Sum
        Next
    Next

Else



End If

End Sub

Public Sub ConjugateGradient(Func As Object, Xk() As Double, CV As clsControlVolumePair, SysEq As clsSystemEquations, ErrorReductionFactor As Double, MaxIter As Long, XkMax() As Double, XkMin() As Double)

' This algorithm provides a new search direction for a nonlinear set of equations g(X) = 0, Where X is a 1 to n vector. The evaluation of g must occur outside of this function
' and this function has to be applied iteratively to work.  The Output dk is the next search direction

'Xk - the current variable solution of guess g(Xk) = Error_k which needs to approach zero (kth step)
'gk - the current result output (kth step)
'gkm1 - the k-1th result output (k-1 th step).  If this is the first step, then gkm1 = 0 just produces the steepest gradient method which is an acceptable first step.
'dkm1 - the k-1th direction. If this is the first step then dkm1 = 0 just produces the steepest gradient method which is an acceptable first step.
'Xkp1 - the next variable solution (k+1th step) g(Xkp1) = Error_k+1 - a test is performed
'ConvergedError - The amount of error allowed for a summation of the entire Xkp1 vector.

Dim beta As Double
Dim alpha As Double
Dim gk() As Double
Dim gkm1() As Double
Dim dk() As Double 'Direction
Dim Bool As Boolean
Dim Numerator As Double
Dim denominator
Dim i As Long '
Dim j As Long
Dim n As Long
Dim Error As Double
Dim Error1 As Double
Dim ErrorP() As Double
Dim A As Double 'Parabola coefficients
Dim B As Double
Dim c As Double
Dim Dum As Double
Dim alpha_guess() As Double
Dim Xkp1() As Double
Dim Iter As Long
Dim Magdk As Double
Dim TrySteepestDescent As Boolean

' These parameters may need to be adjusted to make the solution procedure robust. If we
' have a quadratic surface, then they will be optimal.
'ReDim alpha_guess(1 To 3)
'ReDim ErrorP(1 To 3)
'alpha_guess(1) = 0.001
'alpha_guess(2) = 0.002
'alpha_guess(3) = 0.003

' This can be any class as long as it has an EvaluateFunction Subroutine and has the four required arguments below
' Xk is the first guess at the solution, CV is a control volume pair (Doesn't necessarily have to be used)
' SysEq is the system of equations object (doesn't have to be used but gives access to all material properties)
' and gk is the result from the evaluation of the equations.

n = UBound(Xk)

'Initialize Values
ReDim gkm1(1 To n)
ReDim dkm1(1 To n)
ReDim Xkp1(1 To n)
ReDim dk(1 To n)
For i = 1 To n
   gkm1(i) = 0 ' Setting all of these to zero makes the first step the steepest descent method.
   dkm1(i) = 0
Next
Iter = 0
TrySteepestDescent = False

' First evaluation
Func.EvaluateFunction Xk, CV, SysEq, gk
mdlMath.SqrtSumSquareOfVector gk, Error1


Error = Error1

Do While Error / Error1 > ErrorReductionFactor And Iter < MaxIter
    '
    'beta = gkT(gk - gkm1) / gkm1T * gkm1 - Polak Ribiere direction
    Numerator = 0
    denominator = 0
    For i = 1 To n
       Numerator = Numerator + gk(i) * (gk(i) - gkm1(i))
       denominator = denominator + gkm1(i) ^ 2
    Next i
Restart:
    If denominator = 0 Or TrySteepestDescent Then ' Revert to steepest ascent, this should only happen the first time through.
       beta = 0
       TrySteepestDescent = False
    Else
       ' reset to steepest decsent if  directions are no longer conjugate
       beta = mdlMath.DblMax(0, Numerator / denominator)
    End If
    
    'Calculate the new search direction - steepest gradient with an adjustment to stay "A" orthogonal to the error.
    ' dk = -gk + beta * dkm1
    Magdk = 1
    For i = 1 To n
        dk(i) = -gk(i) + beta * dkm1(i) '
        ' See if we need to adjust the magnitude to keep from flying out of the range of interest
        If dk(i) > XkMax(i) - Xk(i) And XkMax(i) - Xk(i) <> 0 Then
           Magdk = mdlMath.DblMax(Magdk, mdlMath.AbsDbl(dk(i) / (XkMax(i) - Xk(i))))
        ElseIf dk(i) < XkMin(i) - Xk(i) And XkMin(i) - Xk(i) <> 0 Then
           Magdk = mdlMath.DblMax(Magdk, mdlMath.AbsDbl(dk(i) / (Xk(i) - XkMin(i))))
        End If
    Next i
    For i = 1 To n
        dk(i) = dk(i) / Magdk
    Next i
    
    'Perform a line search
    
      
    ' Make a quadratic fit to 3 points in the error direction to the new search direction and then go to the resulting quadratic minimum. alpha = 0.1, 0.25, 0.5
' THIS APPROACH WAS NOT ROBUST
'    mdlMath.FitParabolaTo3Points alpha_guess(1), ErrorP(1), alpha_guess(2), ErrorP(2), alpha_guess(3), ErrorP(3), a, b, c
'    ' check to see if the error direction is upward facing (conjugate gradient direction calculation should avoid this!)
'    If a <= 0 Then
'       If TrySteepestDescent Then
'          Dum = mdlError.ReturnError("mdlMath.ConjugateGradient: The calculated search direction has produced a negative parabola or line that has a maximum rather than a minimum", , True)
'          GoTo ErrorHappened
'       Else
'          TrySteepestDescent = True
'          GoTo Restart
'       End If
'    End If
    
    ' calculate alpha with the projected minimum and take a step
    alpha = GoldenLineSearch(Func, dk, Xk, 0, 1, ErrorReductionFactor, MaxIter, CV, SysEq)
    'Update Xk
    For i = 1 To n
       Xk(i) = Xk(i) + alpha * dk(i)
       ' Save old direction and function evaluation
       gkm1(i) = gk(i)
       dkm1(i) = dk(i)
    Next
    
    'Reevaluate function and error
    Func.EvaluateFunction Xk, CV, SysEq, gk
    mdlMath.SqrtSumSquareOfVector gk, Error
    
    Iter = Iter + 1

Loop 'Check for convergence again.

If Iter > MaxIter Then
   Dum = mdlError.ReturnError("mdlMath.ConjugateGradient: Maximum number of iterations has been exceeded without meeting the error tolerance needed!", , True)
   GoTo ErrorHappened
End If

EndOfSub:



Exit Sub
ErrorHappened:

GoTo EndOfSub
End Sub

Public Function GoldenLineSearch(Func As Object, dk() As Double, Xk() As Double, A As Double, B As Double, Tol As Double, MaxIter As Long, CV As clsControlVolumePair, SysEq As clsSystemEquations) As Double

' Interval search is from a to b.
Dim c As Double
Dim d As Double
Dim gr As Double
Dim Iter As Long


gr = (5 ^ 0.5 + 1) / 2

c = B - (B - A) / gr
d = A + (B - A) / gr

Do While mdlMath.AbsDbl(c - d) > Tol And Iter < MaxIter

   If NewErrorDueToIncrement(Func, Xk, dk, c, CV, SysEq) < NewErrorDueToIncrement(Func, Xk, dk, d, CV, SysEq) Then
       B = d
   Else
       A = c
   End If
   
   ' We recalculate c and d here to avoid loss of precision which may lead to
   ' incorrect results or infinite loop
   c = B - (B - A) / gr
   d = A + (B - A) / gr

   Iter = Iter + 1
Loop

GoldenLineSearch = (A + B) / 2

End Function

Private Function NewErrorDueToIncrement(Func As Object, Xk() As Double, dk() As Double, alpha As Double, CV As clsControlVolumePair, SysEq As clsSystemEquations) As Double

Dim n As Long
Dim Xknew() As Double
Dim i As Long
Dim gk() As Double
Dim ErrorA As Double
   
   
   n = UBound(Xk)
   ReDim Xknew(1 To n)
   ReDim gk(1 To n)
   
   For i = 1 To n
      Xknew(i) = Xk(i) + alpha * dk(i)
   Next
   Func.EvaluateFunction Xknew, CV, SysEq, gk
   mdlMath.SqrtSumSquareOfVector gk, ErrorA
   
   NewErrorDueToIncrement = ErrorA
  
End Function


Public Sub SqrtSumSquareOfVector(x() As Double, SumSquare As Double)

Dim i As Long

SumSquare = 0

For i = LBound(x) To UBound(x)
    SumSquare = SumSquare + x(i) ^ 2
Next

SumSquare = SumSquare ^ 0.5

End Sub

Public Sub SqrtSumSquareOfNormalizedVector(x() As Double, Xnorm() As Double, SumSquare As Double)

Dim i As Long

SumSquare = 0

For i = LBound(x) To UBound(x)
    If Xnorm(i) <> 0 Then
        SumSquare = SumSquare + (x(i) / Xnorm(i)) ^ 2
    End If
Next

SumSquare = SumSquare ^ 0.5

End Sub

Function DblMax(x1 As Double, x2 As Double) As Double

   If x1 > x2 Then
       DblMax = x1
   Else
       DblMax = x2
   End If

End Function

Function DblMin(x1 As Double, x2 As Double) As Double

   If x1 < x2 Then
       DblMin = x1
   Else
       DblMin = x2
   End If

End Function

Public Sub GaussianElimination(A() As Double, B() As Double, Xi() As Double)
' Assure that none of the A(k,k) values are zero
' This procedure changes Xi
Dim n As Long
Dim k As Long
Dim i As Long
Dim j As Long
Dim S As Double
Dim Bool As Boolean
Dim Dum As Double

If LBound(A, 1) <> 1 Or LBound(A, 2) <> 1 Then
    Dum = mdlError.ReturnError("mdlMath.GaussianElimination: matrix and vector input must have first element = 1!", , True)
Else

    n = UBound(A, 1)
    
    For i = 1 To n
       If A(i, i) = 0 Then
           Dum = mdlError.ReturnError("mdlMath.GaussianElimination: a diagonal entry has a value of 0, Gaussian elimination will not work!", , True)
       End If
    Next
    
    ReDim Xi(1 To n)
    ReDim Xi_Map(1 To n)
    
    ' Gaussian Elimination
    For k = 1 To n - 1
        For i = k + 1 To n
            If A(i, k) <> 0 Then
                If A(k, k) = 0 Then
                   Dum = mdlError.ReturnError("mdlMath.GaussianElimination: During solution it is clear that all of the rows of the" & _
                                              " matrix provided are not linearly independent. No solution exists!", , True)
                End If
                A(i, k) = A(i, k) / A(k, k)
                For j = k + 1 To n
                    A(i, j) = A(i, j) - A(i, k) * A(k, j)
                Next
                'Forward elimination
                B(i) = B(i) - A(i, k) * B(k)
            End If
        Next i
    Next k
    
    'Backward Solve
    For i = n To 1 Step -1
        S = B(i)
        For j = i + 1 To n
            S = S - A(i, j) * Xi(j)
        Next
        Xi(i) = S / A(i, i)
        
    Next

End If

End Sub

Public Sub NewtonsMethod(Func As Object, Xk() As Double, CV As clsControlVolumePair, SysEq As clsSystemEquations, _
                         ErrorReductionFactor As Double, MaxIter As Long, DerivativeIncrement As Double, _
                         XkMax() As Double, XkMin() As Double, AbsConvergenceCriteria() As Double, Optional EvaluateConstraints As Boolean = False)

Dim gk() As Double ' This is the function evaluation
Dim gknorm() As Double ' These are normalization factors for the error (since equations of different units are being used)
Dim gkinc() As Double ' incremented solution for evaluation of a numerical derivative
Dim Jacobian() As Double ' This is the jacobian matrix
'Dim JacobianSorted() As Double
Dim InverseJacobian() As Double
Dim Xkp1() As Double ' X_k+1 - next solution step
Dim Xinc() As Double
Dim Bool As Boolean
Dim i As Long
Dim j As Long
Dim n As Long
Dim minus_gk() As Double
Dim Error1 As Double
Dim Error As Double
Dim ErrorHist() As Double
Dim Iter As Long
Dim SortRowsNeeded As Boolean
Dim RareCase As Boolean
Dim RareCaseMinus1 As Boolean
Dim TempVariant As Variant
Dim Dum As Double
Dim max_inc As Double
Dim Out As clsOutput
Dim AbsError As Double
Dim AbsConverged As Boolean
Dim red_inc As Double
Dim Xkp1inc() As Double
Dim Xkold() As Double 'see the last step!
Dim Xkp1old() As Double

Set Out = New clsOutput

max_inc = mdlConstants.glbNewtonMaxIncrementToMaxOrMin

' Func - This can be any class as long as it has an EvaluateFunction Subroutine and has the four required arguments below
' Xk is the first guess at the solution, CV is a control volume pair (Doesn't necessarily have to be used)
' SysEq is the system of equations object (doesn't have to be used but gives access to all material properties)
' and gk is the result from the evaluation of the equations.

n = UBound(Xk)

'Initialize Values
ReDim Xkp1(1 To n)
ReDim Xkp1inc(1 To n)
ReDim Xinc(1 To n)
ReDim gkinc(1 To n)
ReDim gknorm(1 To n)
ReDim minus_gk(1 To n)
ReDim Jacobian(1 To n, 1 To n)
ReDim gk(1 To n)
ReDim Xkold(1 To n) ' this variable serves no purpose besides debugging (memory hog!)
ReDim Xkp1old(1 To n) ' this variable serves no purpose besides debugging

Set Out = New clsOutput


' First evaluation
Func.EvaluateFunction Xk, CV, SysEq, gknorm
mdlMath.SqrtSumSquareOfVector gknorm, AbsError
AbsConverged = True ' Will be changed unless all convergence criteria are met!
For i = 1 To UBound(gknorm)
   If Abs(gknorm(i)) < AbsConvergenceCriteria(i) Then
      gknorm(i) = 1  ' This guarantees that this measure is already considered convergent.
   Else
      AbsConverged = False
   End If
Next

' All errors for now are gknorm/gknorm so the simplified error is:
Error1 = (UBound(gknorm)) ^ 0.5
Error = Error1
ReDim ErrorHist(0 To MaxIter)
ErrorHist(0) = Error
    
Do While Error / Error1 > ErrorReductionFactor And Iter < MaxIter And Not AbsConverged
    
    ' You cannot just set a break point on this function because it is used recursively. You can capture
    ' the outer loop by setting an if-then statement that has at least 12 variables!
    
    If mdlConstants.glbTroubleshoot And UBound(Xk) >= 12 Then
       Out.HeatExchangerPerformanceCurves Func
    End If
    
    For i = 1 To n
       Xinc(i) = Xk(i)
    Next
    
    ' Form the jacobian matrix based on numerical derivatives. - continuous functions are expected!
    SortRowsNeeded = False
    
    For i = 1 To n
       ' Rare case tracks the remote possibility that Xinc(i) = DerivativeIncrement / (1 + DerivativeIncrement)
       ' which would satisfy the criterion to set Xinc(i) = 0 after applying the increment.
       RareCase = False
       If Xinc(i) = 0 Then
          Xinc(i) = DerivativeIncrement
       Else
          'This is only a remote possibility but we need to eliminate it so that
          'we do not inadvertently reset a value to zero
          If Xinc(i) = DerivativeIncrement / (1 + DerivativeIncrement) Then
             RareCase = True
          End If
          Xinc(i) = Xinc(i) * (1 + DerivativeIncrement)
       End If
       If i > 1 Then
          If Xinc(i - 1) = DerivativeIncrement And Not RareCaseMinus1 Then
             Xinc(i - 1) = 0
          Else
             Xinc(i - 1) = Xinc(i - 1) / (1 + DerivativeIncrement)
          End If
       End If
       
       ' I wish I didn't have to add this function evaluation but for
       ' the control volume pair functions something is going wrong and
       ' I have to re-calculate. This is very inefficient
       Func.EvaluateFunction Xk, CV, SysEq, gk
       
       If EvaluateConstraints Then
        ' THIS MAY REDUCE Xkp1 by a factor to avoid violating constraints that are expressed in Func.EvaluateConstraints!!!
           For j = 1 To n
               Xkp1inc(j) = Xinc(j) - Xk(j)
           Next
           ' Make sure that the increment does not violate constraints - set back the increment if it does! - this will drive
           ' derivatives to zero if extremely small steps have to be taken!
           Func.EvaluateConstraints Xk, Xkp1inc, glbMaxReductionInStepSizeDueToConstraint
           For j = 1 To n
               Xinc(j) = Xk(j) + Xkp1inc(j)
           Next
       Else
           
       End If
       
       
       Func.EvaluateFunction Xinc, CV, SysEq, gkinc
       
       For j = 1 To n
           If Xinc(i) = DerivativeIncrement And Not RareCase Then
               Jacobian(j, i) = (gkinc(j) - gk(j)) / DerivativeIncrement
           Else
               Jacobian(j, i) = (gkinc(j) - gk(j)) / (DerivativeIncrement * Xinc(i))
           End If
       Next
       If Jacobian(i, i) = 0 Then
          SortRowsNeeded = True
       End If
       
       ' We have to keep the i-1 rare case for the next round to remember if we need to restore
       ' a non-zero value
       RareCaseMinus1 = RareCase
    Next i
    
    
    For i = 1 To n
       minus_gk(i) = -gk(i)
    Next i
    
    If SortRowsNeeded Then
    
       ReDim InverseJacobian(1 To n, 1 To n)
       'Rearrange the rows to assure that all diagonal entries are non-zero
       TempVariant = Jacobian
       'Sheet6.Range("A1:CR96") = TempVariant
       
       'mdlMath.SortMatrixDiagonalNonZero Jacobian, JacobianSorted, minus_gk
       'GaussianElimination JacobianSorted, minus_gk, Xkp1
        
        ' Not that efficient but my sort routine would not work and gets stuck computing for a
        ' very long time.  This is faster.
        TempVariant = Application.WorksheetFunction.MInverse(TempVariant)
        
        CDblArr TempVariant, InverseJacobian
        
        mdlMath.MatrixMultiply InverseJacobian, minus_gk, Xkp1
       
    Else
       GaussianElimination Jacobian, minus_gk, Xkp1
    End If



    'Update Xk but relax the solution if it overshoots the maximum or minimum limits only move max_inc of the way to the maximum or minimum
    '
    For i = 1 To n
       If Xkp1(i) + Xk(i) > XkMax(i) Then
           Xkp1(i) = max_inc * (XkMax(i) - Xk(i))
       ElseIf Xkp1(i) + Xk(i) < XkMin(i) Then
           Xkp1(i) = max_inc * (XkMin(i) - Xk(i))
       End If
    Next i
    
    'Update the solution based on any non-square constraints - this was first needed for solving the air gap algorithm
    '7/19/2017 Though this is a very helpful addition, air-gap no longer uses it
    If EvaluateConstraints Then
        ' THIS MAY REDUCE Xkp1 by a factor to avoid violating constraints that are expressed in Func.EvaluateConstraints!!!
        For i = 1 To n
            Xkold(i) = Xk(i)
            Xkp1old(i) = Xkp1(i)
        Next
        Func.EvaluateConstraints Xk, Xkp1, glbMaxReductionInStepSizeDueToConstraint
    End If
    
    For i = 1 To n
       Xk(i) = glbNewtonReductionFactor * Xkp1(i) + Xk(i)
    Next i
    
    ' Reevaluate the function and get ready to check the error again.
    If mdlError.NoError Then
        Func.EvaluateFunction Xk, CV, SysEq, gk
        mdlMath.SqrtSumSquareOfNormalizedVector gk, gknorm, Error
        ErrorHist(Iter + 1) = Error
        AbsConverged = True
        For i = 1 To n
           If gk(i) > AbsConvergenceCriteria(i) Then
              AbsConverged = False
              Exit For
           End If
        Next
    Else
        GoTo ErrorHappened
    End If

    Iter = Iter + 1
    
Loop
mdlMath.SqrtSumSquareOfVector gk, AbsError
If (Error / Error1 > ErrorReductionFactor) And Not AbsConverged Then
    Dum = mdlError.ReturnError("mdlMath.NewtonsMethod: Maximum number of iterations for convergence happened with no convergence!", , True)
    GoTo ErrorHappened
End If
    

If glbEvaluateEquationsAtEnd Then
   If IsFinalRun Then
       Range("IsFinalNewtonIterationRange").Value = "TRUE"
       Func.EvaluateFunction Xk, CV, SysEq, gk
   End If
End If


EndOfSub:

If mdlConstants.glbTroubleshoot And UBound(Xk) >= 12 Then 'Troubleshoot
   Iter = Iter
End If

Exit Sub
ErrorHappened:
       MsgBox err.Source
GoTo EndOfSub
End Sub


Public Sub CalculateLeastSquareErrorBetweenVectors(x1() As Double, x2() As Double, LeastSquareErrorOut As Double)

Dim X1U As Long
Dim X1L As Long
Dim X2U As Long
Dim X2L As Long
Dim i As Long
Dim Dum As Double

X1U = UBound(x1)
X1L = LBound(x1)
X2U = UBound(x2)
X2L = LBound(x2)

If X1U - X1L <> X2U - X2L Then
    Dum = mdlError.ReturnError("mdlMath.CalculateLeastSquareError: X1 and X2 inputs must be of the same length!")
End If

LeastSquareErrorOut = 0

For i = X1L To X1U
    LeastSquareErrorOut = LeastSquareErrorOut + (x1(i) - x2(i - (X1L - X2L))) ^ 2
Next

End Sub

Public Sub FitParabolaTo3Points(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, A As Double, B As Double, c As Double)
' This function just gives the somewhat long analytical solution to the following problem
      ' [A]   [X1^2 X1 1]^-1 [Y1]
      ' [B] = [X2^2 X2 1]    [Y2]
      ' [C]   [X3^2 X3 1]    [Y3]
' This function was tested and shown to provide the correct answers required.
      Dim detA As Double
      
      Dim X1sq As Double
      Dim X2sq As Double
      Dim X3sq As Double
      
      X1sq = x1 ^ 2
      X2sq = x2 ^ 2
      X3sq = x3 ^ 2
      
      
      detA = X1sq * (x2 - x3) - x1 * (X2sq - X3sq) + (X2sq * x3 - x2 * X3sq)
      'Cramers rule
      A = ((x2 - x3) * y1 + (x3 - x1) * y2 + (x1 - x2) * y3) / detA
      B = ((X3sq - X2sq) * y1 + (X1sq - X3sq) * y2 + (X2sq - X1sq) * y3) / detA
      ' this third value did not work and I used Cramer's rule instead to derive the formula.
      c = (X1sq * (y3 * x2 - y2 * x3) - x1 * (X2sq * y3 - X3sq * y2) + y1 * (X2sq * x3 - X3sq * x2)) / detA
End Sub







