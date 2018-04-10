Attribute VB_Name = "ArrayMod_shared"
' The following functions come from
' http://www.cpearson.com/excel/vbaarrays.htm.
Option Explicit
Option Compare Text
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modArraySupport
' By Chip Pearson, chip@cpearson.com, www.cpearson.com
'
' This module contains procedures that provide information about and manipulate
' VB/VBA arrays. NOTE: These functions call one another. It is strongly suggested
' that you Import this entire module to a VBProject rather then copy/pasting
' individual procedures.
'
' For details on these functions, see www.cpearson.com/excel/VBAArrays.htm
'
' This module contains the following functions:
'       AreDataTypesCompatible
'       ChangeBoundsOfArray
'       CombineTwoDArrays
'       CompareArrays
'       ConcatenateArrays
'       CopyArray
'       CopyArraySubSetToArray
'       CopyNonNothingObjectsToArray
'       DataTypeOfArray
'       DeleteArrayElement
'       ExpandArray
'       FirstNonEmptyStringIndexInArray
'       GetColumn
'       GetRow
'       InsertElementIntoArray
'       IsArrayAllDefault
'       IsArrayAllNumeric
'       IsArrayAllocated
'       IsArrayDynamic
'       IsArrayEmpty
'       IsArrayObjects
'       IsArraySorted
'       IsNumericDataType
'       IsVariantArrayConsistent
'       IsVariantArrayNumeric
'       ModuleBase
'       MoveEmptyStringsToEndOfArray
'       NumberOfArrayDimensions
'       NumElements
'       ResetVariantArrayToDefaults
'       ReverseArrayInPlace
'       ReverseArrayOfObjectsInPlace
'       SetObjectArrayToNothing
'       SetVariableToDefault
'       TransposeArray
'       VectorsToArray
'
' Function documentation is in each function.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''
' Error Number Constants
'''''''''''''''''''''''''''
Public Const C_ERR_NO_ERROR = 0&
Public Const C_ERR_SUBSCRIPT_OUT_OF_RANGE = 9&
Public Const C_ERR_ARRAY_IS_FIXED_OR_LOCKED = 10&

        
Public Function CompareArrays(Array1 As Variant, Array2 As Variant, _
    ResultArray As Variant, Optional CompareMode As VbCompareMethod = vbTextCompare) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CompareArrays
' This function compares two arrays, Array1 and Array2, element by element, and puts the results of
' the comparisons in ResultArray. Each element of ResultArray will be -1, 0, or +1. A -1 indicates that
' the element in Array1 was less than the corresponding element in Array2. A 0 indicates that the
' elements are equal, and +1 indicates that the element in Array1 is greater than Array2. Both
' Array1 and Array2 must be allocated single-dimensional arrays, and ResultArray must be dynamic array
' of a numeric data type (typically Longs). Array1 and Array2 must contain the same number of elements,
' and have the same lower bound. The LBound of ResultArray will be the same as the data arrays.
'
' An error will occur if Array1 or Array2 contains an Object or User Defined Type.
'
' When comparing elements, the procedure does the following:
' If both elements are numeric data types, they are compared arithmetically.

' If one element is a numeric data type and the other is a string and that string is numeric,
' then both elements are converted to Doubles and compared arithmetically. If the string is not
' numeric, both elements are converted to strings and compared using StrComp, with the
' compare mode set by CompareMode.
'
' If both elements are numeric strings, they are converted to Doubles and compared arithmetically.
'
' If either element is not a numeric string, the elements are converted and compared with StrComp.
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx1 As Long
Dim Ndx2 As Long
Dim ResNdx As Long
Dim S1 As String
Dim S2 As String
Dim d1 As Double
Dim d2 As Double
Dim Done As Boolean
Dim Compare As VbCompareMethod
Dim LB As Long

''''''''''''''''''''''''''''''''''''
' Set the default return value.
''''''''''''''''''''''''''''''''''''
CompareArrays = False

''''''''''''''''''''''''''''''''''''
' Ensure we have a Compare mode
' value.
''''''''''''''''''''''''''''''''''''
If CompareMode = vbBinaryCompare Then
    Compare = vbBinaryCompare
Else
    Compare = vbTextCompare
End If


''''''''''''''''''''''''''''''''''''
' Ensure we have arrays.
''''''''''''''''''''''''''''''''''''
If IsArray(Array1) = False Then
Exit Function
End If
If IsArray(Array2) = False Then
Exit Function
End If
If IsArray(ResultArray) = False Then
Exit Function
End If

'''''''''''''''''''''''''''''''''''
' Ensure ResultArray is dynamic
'''''''''''''''''''''''''''''''''''
If IsArrayDynamic(arr:=ResultArray) = False Then
Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''
' Ensure the arrays are single-dimensional.
''''''''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(arr:=Array1) <> 1 Then
Exit Function
End If
If NumberOfArrayDimensions(arr:=Array2) <> 1 Then
Exit Function
End If
If NumberOfArrayDimensions(arr:=Array1) > 1 Then 'allow 0 indicating non-allocated array
Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''
' Ensure the LBounds are the same
''''''''''''''''''''''''''''''''''''''''''
If LBound(Array1) <> LBound(Array2) Then
Exit Function
End If
    

''''''''''''''''''''''''''''''''''''''''''
' Ensure the arrays are the same size.
''''''''''''''''''''''''''''''''''''''''''
If (UBound(Array1) - LBound(Array1)) <> (UBound(Array2) - LBound(Array2)) Then
Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''
' Redim ResultArray to the numbr of elements
' in Array1.
''''''''''''''''''''''''''''''''''''''''''''''
ReDim ResultArray(LBound(Array1) To UBound(Array1))

Ndx1 = LBound(Array1)
Ndx2 = LBound(Array2)

''''''''''''''''''''''''''''''''''''''''''''''
' Scan each array to see if it contains objects
' or User-Defined Types. If found, exit with
' False.
''''''''''''''''''''''''''''''''''''''''''''''
For Ndx1 = LBound(Array1) To UBound(Array1)
    If IsObject(Array1(Ndx1)) = True Then
Exit Function
    End If
    If VarType(Array1(Ndx1)) >= vbArray Then
Exit Function
    End If
    If VarType(Array1(Ndx1)) = vbUserDefinedType Then
Exit Function
    End If
Next Ndx1

For Ndx1 = LBound(Array2) To UBound(Array2)
    If IsObject(Array2(Ndx1)) = True Then
Exit Function
    End If
    If VarType(Array2(Ndx1)) >= vbArray Then
Exit Function
    End If
    If VarType(Array2(Ndx1)) = vbUserDefinedType Then
Exit Function
    End If
Next Ndx1

Ndx1 = LBound(Array1)
Ndx2 = Ndx1
ResNdx = LBound(ResultArray)
Done = False
Do Until Done = True
''''''''''''''''''''''''''''''''''''
' Loop until we reach the end of
' the array.
''''''''''''''''''''''''''''''''''''
    If IsNumeric(Array1(Ndx1)) = True And IsNumeric(Array2(Ndx2)) Then
        d1 = CDbl(Array1(Ndx1))
        d2 = CDbl(Array2(Ndx2))
        If d1 = d2 Then
            ResultArray(ResNdx) = 0
        ElseIf d1 < d2 Then
            ResultArray(ResNdx) = -1
        Else
            ResultArray(ResNdx) = 1
        End If
    Else
        S1 = CStr(Array1(Ndx1))
        S2 = CStr(Array2(Ndx1))
        ResultArray(ResNdx) = StrComp(S1, S2, Compare)
    End If
        
    ResNdx = ResNdx + 1
    Ndx1 = Ndx1 + 1
    Ndx2 = Ndx2 + 1
    ''''''''''''''''''''''''''''''''''''''''
    ' If Ndx1 is greater than UBound(Array1)
    ' we've hit the end of the arrays.
    ''''''''''''''''''''''''''''''''''''''''
    If Ndx1 > UBound(Array1) Then
        Done = True
    End If
Loop

CompareArrays = True
End Function



Public Function ConcatenateArrays(ResultArray As Variant, ArrayToAppend As Variant, _
        Optional NoCompatabilityCheck As Boolean = False) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ConcatenateArrays
' This function appends ArrayToAppend to the end of ResultArray, increasing the size of ResultArray
' as needed. ResultArray must be a dynamic array, but it need not be allocated. ArrayToAppend
' may be either static or dynamic, and if dynamic it may be unallocted. If ArrayToAppend is
' unallocated, ResultArray is left unchanged.
'
' The data types of ResultArray and ArrayToAppend must be either the same data type or
' compatible numeric types. A compatible numeric type is a type that will not cause a loss of
' precision or cause an overflow. For example, ReturnArray may be Longs, and ArrayToAppend amy
' by Longs or Integers, but not Single or Doubles because information might be lost when
' converting from Double to Long (the decimal portion would be lost). To skip the compatability
' check and allow any variable type in ResultArray and ArrayToAppend, set the NoCompatabilityCheck
' parameter to True. If you do this, be aware that you may loose precision and you may will
' get an overflow error which will cause a result of 0 in that element of ResultArra.
'
' Both ReaultArray and ArrayToAppend must be one-dimensional arrays.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim VTypeResult As VbVarType
Dim Ndx As Long
Dim Res As Long
Dim NumElementsToAdd As Long
Dim AppendNdx As Long
Dim VTypeAppend As VbVarType
Dim ResultLB As Long
Dim ResultUB As Long
Dim ResultWasAllocated As Boolean

'''''''''''''''''''''''''''''''''
' Set the default result.
''''''''''''''''''''''''''''''''
ConcatenateArrays = False

'''''''''''''''''''''''''''''''''
' Ensure ResultArray is an array.
'''''''''''''''''''''''''''''''''
If IsArray(ResultArray) = False Then
Exit Function
End If
''''''''''''''''''''''''''''''''''
' Ensure ArrayToAppend is an array.
''''''''''''''''''''''''''''''''''
If IsArray(ArrayToAppend) = False Then
Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure both arrays are single
' dimensional.
''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(ResultArray) > 1 Then
Exit Function
End If
If NumberOfArrayDimensions(ArrayToAppend) > 1 Then
Exit Function
End If
'''''''''''''''''''''''''''''''''''
' Ensure ResultArray is dynamic.
'''''''''''''''''''''''''''''''''''
If IsArrayDynamic(arr:=ResultArray) = False Then
Exit Function
End If

''''''''''''''''''''''''''''''''''''
' Ensure ArrayToAppend is allocated.
' If ArrayToAppend is not allocated,
' we have nothing to append, so
' exit with a True result.
''''''''''''''''''''''''''''''''''''
If IsArrayAllocated(arr:=ArrayToAppend) = False Then
    ConcatenateArrays = True
Exit Function
End If


If NoCompatabilityCheck = False Then
    ''''''''''''''''''''''''''''''''''''''
    ' Ensure the array are compatible
    ' data types.
    ''''''''''''''''''''''''''''''''''''''
    If AreDataTypesCompatible(DestVar:=ResultArray, SourceVar:=ArrayToAppend) = False Then
        '''''''''''''''''''''''''''''''''''''''''''
        ' The arrays are not compatible data types.
        '''''''''''''''''''''''''''''''''''''''''''
Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''
    ' If one array is an array of
    ' objects, ensure the other contains
    ' all objects (or Nothing)
    ''''''''''''''''''''''''''''''''''''
    If VarType(ResultArray) - vbArray = vbObject Then
        If IsArrayAllocated(ArrayToAppend) = True Then
            For Ndx = LBound(ArrayToAppend) To UBound(ArrayToAppend)
                If IsObject(ArrayToAppend(Ndx)) = False Then
Exit Function
                End If
            Next Ndx
        End If
    End If
End If
    
    
'''''''''''''''''''''''''''''''''''''''
' Get the number of elements in
' ArrrayToAppend
'''''''''''''''''''''''''''''''''''''''
NumElementsToAdd = UBound(ArrayToAppend) - LBound(ArrayToAppend) + 1
''''''''''''''''''''''''''''''''''''''''
' Get the bounds for resizing the
' ResultArray. If ResultArray is allocated
' use the LBound and UBound+1. If
' ResultArray is not allocated, use
' the LBound of ArrayToAppend for both
' the LBound and UBound of ResultArray.
''''''''''''''''''''''''''''''''''''''''

If IsArrayAllocated(arr:=ResultArray) = True Then
    ResultLB = LBound(ResultArray)
    ResultUB = UBound(ResultArray)
    ResultWasAllocated = True
    ReDim Preserve ResultArray(ResultLB To ResultUB + NumElementsToAdd)
Else
    ResultUB = UBound(ArrayToAppend)
    ResultWasAllocated = False
    ReDim ResultArray(LBound(ArrayToAppend) To UBound(ArrayToAppend))
End If

''''''''''''''''''''''''''''''''''''''''
' Copy the data from ArrayToAppend to
' ResultArray.
''''''''''''''''''''''''''''''''''''''''
If ResultWasAllocated = True Then
    ''''''''''''''''''''''''''''''''''''''''''
    ' If ResultArray was allocated, we
    ' have to put the data from ArrayToAppend
    ' at the end of the ResultArray.
    ''''''''''''''''''''''''''''''''''''''''''
    AppendNdx = LBound(ArrayToAppend)
    For Ndx = ResultUB + 1 To UBound(ResultArray)
        If IsObject(ArrayToAppend(AppendNdx)) = True Then
            Set ResultArray(Ndx) = ArrayToAppend(AppendNdx)
        Else
            ResultArray(Ndx) = ArrayToAppend(AppendNdx)
        End If
        AppendNdx = AppendNdx + 1
        If AppendNdx > UBound(ArrayToAppend) Then
            Exit For
        End If
    Next Ndx
Else
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' If ResultArray was not allocated, we simply
    ' copy element by element from ArrayToAppend
    ' to ResultArray.
    ''''''''''''''''''''''''''''''''''''''''''''''
    For Ndx = LBound(ResultArray) To UBound(ResultArray)
        If IsObject(ArrayToAppend(Ndx)) = True Then
            Set ResultArray(Ndx) = ArrayToAppend(Ndx)
        Else
            ResultArray(Ndx) = ArrayToAppend(Ndx)
        End If
    Next Ndx

End If
'''''''''''''''''''''''
' Success. Return True.
'''''''''''''''''''''''
ConcatenateArrays = True

End Function

Public Function CopyArray(DestinationArray As Variant, SourceArray As Variant, _
        Optional NoCompatabilityCheck As Boolean = False) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CopyArray
' This function copies the contents of SourceArray to the DestinationaArray. Both SourceArray
' and DestinationArray may be either static or dynamic and either or both may be unallocated.
'
' If DestinationArray is dynamic, it is resized to match SourceArray. The LBound and UBound
' of DestinationArray will be the same as SourceArray, and all elements of SourceArray will
' be copied to DestinationArray.
'
' If DestinationArray is static and has more elements than SourceArray, all of SourceArray
' is copied to DestinationArray and the right-most elements of DestinationArray are left
' intact.
'
' If DestinationArray is static and has fewer elements that SourceArray, only the left-most
' elements of SourceArray are copied to fill out DestinationArray.
'
' If SourceArray is an unallocated array, DestinationArray remains unchanged and the procedure
' terminates.
'
' If both SourceArray and DestinationArray are unallocated, no changes are made to either array
' and the procedure terminates.
'
' SourceArray may contain any type of data, including Objects and Objects that are Nothing
' (the procedure does not support arrays of User Defined Types since these cannot be coerced
' to Variants -- use classes instead of types).
'
' The function tests to ensure that the data types of the arrays are the same or are compatible.
' See the function AreDataTypesCompatible for information about compatible data types. To skip
' this compability checking, set the NoCompatabilityCheck parameter to True. Note that you may
' lose information during data conversion (e.g., losing decimal places when converting a Double
' to a Long) or you may get an overflow (storing a Long in an Integer) which will result in that
' element in DestinationArray having a value of 0.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim VTypeSource As VbVarType
Dim VTypeDest As VbVarType
Dim SNdx As Long
Dim DNdx As Long


'''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''
CopyArray = False

''''''''''''''''''''''''''''''''''
' Ensure both DestinationArray and
' SourceArray are arrays.
''''''''''''''''''''''''''''''''''
If IsArray(DestinationArray) = False Then
Exit Function
End If
If IsArray(SourceArray) = False Then
Exit Function
End If

'''''''''''''''''''''''''''''''''''''
' Ensure DestinationArray and
' SourceArray are single-dimensional.
' 0 indicates an unallocated array,
' which is allowed.
'''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(arr:=SourceArray) > 1 Then
Exit Function
End If
If NumberOfArrayDimensions(arr:=DestinationArray) > 1 Then
Exit Function
End If

''''''''''''''''''''''''''''''''''''
' If SourceArray is not allocated,
' leave DestinationArray intact and
' return a result of True.
''''''''''''''''''''''''''''''''''''
If IsArrayAllocated(arr:=SourceArray) = False Then
    CopyArray = True
Exit Function
End If

If NoCompatabilityCheck = False Then
    ''''''''''''''''''''''''''''''''''''''
    ' Ensure both arrays are the same
    ' type or compatible data types. See
    ' the function AreDataTypesCompatible
    ' for information about compatible
    ' types.
    ''''''''''''''''''''''''''''''''''''''
    If AreDataTypesCompatible(DestVar:=DestinationArray, SourceVar:=SourceArray) = False Then
        CopyArray = False
Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''
    ' If one array is an array of
    ' objects, ensure the other contains
    ' all objects (or Nothing)
    ''''''''''''''''''''''''''''''''''''
    If VarType(DestinationArray) - vbArray = vbObject Then
        If IsArrayAllocated(SourceArray) = True Then
            For SNdx = LBound(SourceArray) To UBound(SourceArray)
                If IsObject(SourceArray(SNdx)) = False Then
Exit Function
                End If
            Next SNdx
        End If
    End If
End If

If IsArrayAllocated(arr:=DestinationArray) = True Then
    If IsArrayAllocated(arr:=SourceArray) = True Then
        '''''''''''''''''''''''''''''''''''''''''''''''''
        ' If both arrays are allocated, copy from
        ' SourceArray to DestinationArray. If
        ' SourceArray is smaller that DesetinationArray,
        ' the right-most elements of DestinationArray
        ' are left unchanged. If SourceArray is larger
        ' than DestinationArray, the right most elements
        ' of SourceArray are not copied.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        DNdx = LBound(DestinationArray)
        On Error Resume Next
        For SNdx = LBound(SourceArray) To UBound(SourceArray)
            If IsObject(SourceArray(SNdx)) = True Then
                Set DestinationArray(DNdx) = SourceArray(DNdx)
            Else
                DestinationArray(DNdx) = SourceArray(DNdx)
            End If
            DNdx = DNdx + 1
            If DNdx > UBound(DestinationArray) Then
                Exit For
            End If
        Next SNdx
        On Error GoTo 0
    Else
        '''''''''''''''''''''''''''''''''''''''''''''''
        ' If SourceArray is not allocated, so we have
        ' nothing to copy. Exit with a result
        ' of True. Leave DestinationArray intact.
        '''''''''''''''''''''''''''''''''''''''''''''''
        CopyArray = True
Exit Function
    End If
        
Else
    If IsArrayAllocated(arr:=SourceArray) = True Then
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' If Destination array is not allocated and
        ' SourceArray is allocated, Redim DestinationArray
        ' to the same size as SourceArray and copy
        ' the elements from SourceArray to DestinationArray.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        On Error Resume Next
        ReDim DestinationArray(LBound(SourceArray) To UBound(SourceArray))
        For SNdx = LBound(SourceArray) To UBound(SourceArray)
            If IsObject(SourceArray(SNdx)) = True Then
                Set DestinationArray(SNdx) = SourceArray(SNdx)
            Else
                DestinationArray(SNdx) = SourceArray(SNdx)
            End If
        Next SNdx
        On Error GoTo 0
    Else
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' If both SourceArray and DestinationArray are
        ' unallocated, we have nothing to copy (this condition
        ' is actually detected above, but included here
        ' for consistancy), so get out with a result of True.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        CopyArray = True
Exit Function
    End If
End If

'''''''''''''''''''''''
' Success. Return True.
'''''''''''''''''''''''
CopyArray = True

End Function



Public Function CopyArraySubSetToArray(InputArray As Variant, ResultArray As Variant, _
    FirstElementToCopy As Long, LastElementToCopy As Long, DestinationElement As Long) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CopyArraySubSetToArray
' This function copies elements of InputArray to ResultArray. It takes the elements
' from FirstElementToCopy to LastElementToCopy (inclusive) from InputArray and
' copies them to ResultArray, starting at DestinationElement. Existing data in
' ResultArray will be overwrittten. If ResultArray is a dynamic array, it will
' be resized if needed. If ResultArray is a static array and it is not large
' enough to copy all the elements, no elements are copied and the function
' returns False.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
Dim SrcNdx As Long
Dim DestNdx As Long
Dim NumElementsToCopy As Long

'''''''''''''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''''''''''''
CopyArraySubSetToArray = False

''''''''''''''''''''''''''''''''''''''''''
' Ensure InputArray and ResultArray are
' arrays.
''''''''''''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
Exit Function
End If
If IsArray(ResultArray) = False Then
Exit Function
End If
'''''''''''''''''''''''''''''''''''''''''''
' Ensure InputArray is single dimensional.
'''''''''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(arr:=InputArray) <> 1 Then
Exit Function
End If
'''''''''''''''''''''''''''''''''''''''''''
' Ensure ResultArray is unallocated or
' single dimensional.
'''''''''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(arr:=ResultArray) > 1 Then
Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''
' Ensure the bounds and indexes are valid.
''''''''''''''''''''''''''''''''''''''''''''
If FirstElementToCopy < LBound(InputArray) Then
Exit Function
End If
If LastElementToCopy > UBound(InputArray) Then
Exit Function
End If
If FirstElementToCopy > LastElementToCopy Then
Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''
' Calc the number of elements we'll copy
' from InputArray to ResultArray.
'''''''''''''''''''''''''''''''''''''''''
NumElementsToCopy = LastElementToCopy - FirstElementToCopy + 1

If IsArrayDynamic(arr:=ResultArray) = False Then
    If (DestinationElement + NumElementsToCopy - 1) > UBound(ResultArray) Then
        '''''''''''''''''''''''''''''''''''''''''''''
        ' ResultArray is static and can't be resized.
        ' There is not enough room in the array to
        ' copy all the data.
        '''''''''''''''''''''''''''''''''''''''''''''
Exit Function
    End If
Else
    ''''''''''''''''''''''''''''''''''''''''''''
    ' ResultArray is dynamic and can be resized.
    ' Test whether we need to resize the array,
    ' and resize it if required.
    '''''''''''''''''''''''''''''''''''''''''''''
    If IsArrayEmpty(arr:=ResultArray) = True Then
        '''''''''''''''''''''''''''''''''''''''
        ' ResultArray is unallocated. Resize it
        ' to DestinationElement + NumElementsToCopy - 1.
        ' This provides empty elements to the left
        ' of the DestinationElement and room to
        ' copy NumElementsToCopy.
        '''''''''''''''''''''''''''''''''''''''''
        ReDim ResultArray(1 To DestinationElement + NumElementsToCopy - 1)
    Else
        '''''''''''''''''''''''''''''''''''''''''''''''''
        ' ResultArray is allocated. If there isn't room
        ' enough in ResultArray to hold NumElementsToCopy
        ' starting at DestinationElement, we need to
        ' resize the array.
        '''''''''''''''''''''''''''''''''''''''''''''''''
        If (DestinationElement + NumElementsToCopy - 1) > UBound(ResultArray) Then
            If DestinationElement + NumElementsToCopy > UBound(ResultArray) Then
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Resize the ResultArray.
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If NumElementsToCopy + DestinationElement > UBound(ResultArray) Then
                    ReDim Preserve ResultArray(LBound(ResultArray) To UBound(ResultArray) + DestinationElement - 1)
                Else
                    ReDim Preserve ResultArray(LBound(ResultArray) To UBound(ResultArray) + NumElementsToCopy)
                End If
            Else
                ''''''''''''''''''''''''''''''''''''''''''''
                ' Resize the array to hold NumElementsToCopy
                ' starting at DestinationElement.
                ''''''''''''''''''''''''''''''''''''''''''''
                ReDim Preserve ResultArray(LBound(ResultArray) To UBound(ResultArray) + NumElementsToCopy - DestinationElement + 2)
            End If
        Else
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' The ResultArray is large enough to hold
            ' NumberOfElementToCopy starting at DestinationElement.
            ' No need to resize the array.
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
    End If
End If


'''''''''''''''''''''''''''''''''''''''''''''''''''
' Copy the elements from InputArray to ResultArray
' Note that there is no type compatibility checking
' when copying the elements.
'''''''''''''''''''''''''''''''''''''''''''''''''''
DestNdx = DestinationElement
For SrcNdx = FirstElementToCopy To LastElementToCopy
    If IsObject(InputArray(SrcNdx)) = True Then
        Set ResultArray(DestNdx) = InputArray(DestNdx)
    Else
        On Error Resume Next
        ResultArray(DestNdx) = InputArray(SrcNdx)
        On Error GoTo 0
    End If
    DestNdx = DestNdx + 1
Next SrcNdx
    
CopyArraySubSetToArray = True
    
End Function



Public Function CopyNonNothingObjectsToArray(ByRef SourceArray As Variant, _
    ByRef ResultArray As Variant, Optional NoAlerts As Boolean = False) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CopyNonNothingObjectsToArray
' This function copies all objects that are not Nothing from SourceArray
' to ResultArray. ResultArray MUST be a dynamic array of type Object or Variant.
' E.g.,
'       Dim ResultArray() As Object ' Or
'       Dim ResultArray() as Variant
'
' ResultArray will be Erased and then resized to hold the non-Nothing elements
' from SourceArray. The LBound of ResultArray will be the same as the LBound
' of SourceArray, regardless of what its LBound was prior to calling this
' procedure.
'
' This function returns True if the operation was successful or False if an
' an error occurs. If an error occurs, a message box is displayed indicating
' the error. To suppress the message boxes, set the NoAlerts parameter to
' True.
'
' This function uses the following procedures. They are declared as Private
' procedures at the end of this module.
'       IsArrayDynamic
'       IsArrayEmpty
'       NumberOfArrayDimensions
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim ResNdx As Long
Dim InNdx  As Long

'''''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''''
CopyNonNothingObjectsToArray = False

'''''''''''''''''''''''''''''''''''
' Ensure SourceArray is an array.
'''''''''''''''''''''''''''''''''''
If IsArray(SourceArray) = False Then
    If NoAlerts = False Then
        MsgBox "SourceArray is not an array."
    End If
Exit Function
End If
'''''''''''''''''''''''''''''''''''
' Ensure SourceArray is a single
' dimensional array.
'''''''''''''''''''''''''''''''''''
Select Case NumberOfArrayDimensions(arr:=SourceArray)
    Case 0
        '''''''''''''''''''''''''''''
        ' Unallocated dynamic array.
        ' Not Allowed.
        '''''''''''''''''''''''''''''
        If NoAlerts = False Then
            MsgBox "SourceArray is an unallocated array."
        End If
Exit Function
        
    Case 1
        '''''''''''''''''''''''''''''
        ' Single-dimensional array.
        ' This is OK.
        '''''''''''''''''''''''''''''
    Case Else
        '''''''''''''''''''''''''''''
        ' Multi-dimensional array.
        ' This is not allowed.
        '''''''''''''''''''''''''''''
        If NoAlerts = False Then
            MsgBox "SourceArray is a multi-dimensional array. This is not allowed."
        End If
Exit Function
End Select
'''''''''''''''''''''''''''''''''''
' Ensure ResultArray is an array.
'''''''''''''''''''''''''''''''''''
If IsArray(ResultArray) = False Then
    If NoAlerts = False Then
        MsgBox "ResultArray is not an array."
    End If
Exit Function
End If
'''''''''''''''''''''''''''''''''''
' Ensure ResultArray is an dynamic.
'''''''''''''''''''''''''''''''''''
If IsArrayDynamic(arr:=ResultArray) = False Then
    If NoAlerts = False Then
        MsgBox "ResultArray is not a dynamic array."
    End If
Exit Function
End If
'''''''''''''''''''''''''''''''''''
' Ensure ResultArray is a single
' dimensional array.
'''''''''''''''''''''''''''''''''''
Select Case NumberOfArrayDimensions(arr:=ResultArray)
    Case 0
        '''''''''''''''''''''''''''''
        ' Unallocated dynamic array.
        ' This is OK.
        '''''''''''''''''''''''''''''
    Case 1
        '''''''''''''''''''''''''''''
        ' Single-dimensional array.
        ' This is OK.
        '''''''''''''''''''''''''''''
    Case Else
        '''''''''''''''''''''''''''''
        ' Multi-dimensional array.
        ' This is not allowed.
        '''''''''''''''''''''''''''''
        If NoAlerts = False Then
            MsgBox "SourceArray is a multi-dimensional array. This is not allowed."
        End If
Exit Function
End Select

'''''''''''''''''''''''''''''''''
' Ensure that all the elements of
' SourceArray are in fact objects.
'''''''''''''''''''''''''''''''''
For InNdx = LBound(SourceArray) To UBound(SourceArray)
    If IsObject(SourceArray(InNdx)) = False Then
        If NoAlerts = False Then
            MsgBox "Element " & CStr(InNdx) & " of SourceArray is not an object."
        End If
Exit Function
    End If
Next InNdx

''''''''''''''''''''''''''''''
' Erase the ResultArray. Since
' ResultArray is dynamic, this
' will relase the memory used
' by ResultArray and return
' the array to an unallocated
' state.
''''''''''''''''''''''''''''''
Erase ResultArray
''''''''''''''''''''''''''''''
' Now, size ResultArray to the
' size of SourceArray. After
' moving all the non-Nothing
' elements, we'll do another
' resize to get ResultArray
' to the used size. This method
' allows us to avoid Redim
' Preserve for every element.
'''''''''''''''''''''''''''''
ReDim ResultArray(LBound(SourceArray) To UBound(SourceArray))

ResNdx = LBound(SourceArray)
For InNdx = LBound(SourceArray) To UBound(SourceArray)
    If Not SourceArray(InNdx) Is Nothing Then
        Set ResultArray(ResNdx) = SourceArray(InNdx)
        ResNdx = ResNdx + 1
    End If
Next InNdx
''''''''''''''''''''''''''''''''''''''''''
' Now that we've copied all the
' non-Nothing elements from SourceArray
' to ResultArray, we call Redim Preserve
' to resize the ResultArray to the size
' actually used. Test ResNdx to see
' if we actually copied any elements.
''''''''''''''''''''''''''''''''''''''''''
If ResNdx > LBound(SourceArray) Then
    '''''''''''''''''''''''''''''''''''''''
    ' If ResNdx > LBound(SourceArray) then
    ' we copied at least one element out of
    ' SourceArray.
    '''''''''''''''''''''''''''''''''''''''
    ReDim Preserve ResultArray(LBound(ResultArray) To ResNdx - 1)
Else
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Otherwise, we didn't copy any elements
    ' from SourceArray (all elements in SourceArray
    ' were Nothing). In this case, Erase ResultArray.
    '''''''''''''''''''''''''''''''''''''''''''''''''
    Erase ResultArray
End If
'''''''''''''''''''''''''''''
' No errors were encountered.
' Return True.
'''''''''''''''''''''''''''''
CopyNonNothingObjectsToArray = True


End Function



Public Function DataTypeOfArray(arr As Variant) As VbVarType
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DataTypeOfArray
'
' Returns a VbVarType value indicating data type of the elements of
' Arr.
'
' The VarType of an array is the value vbArray plus the VbVarType value of the
' data type of the array. For example the VarType of an array of Longs is 8195,
' which equal to vbArray + vbLong. This code subtracts the value of vbArray to
' return the native data type.
'
' If Arr is a simple array, either single- or mulit-
' dimensional, the function returns the data type of the array. Arr
' may be an unallocated array. We can still get the data type of an unallocated
' array.
'
' If Arr is an array of arrays, the function returns vbArray. To retrieve
' the data type of a subarray, pass into the function one of the sub-arrays. E.g.,
' Dim R As VbVarType
' R = DataTypeOfArray(A(LBound(A)))
'
' This function support single and multidimensional arrays. It does not
' support user-defined types. If Arr is an array of empty variants (vbEmpty)
' it returns vbVariant
'
' Returns -1 if Arr is not an array.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Element As Variant
Dim NumDimensions As Long

' If Arr is not an array, return
' vbEmpty and get out.
If IsArray(arr) = False Then
    DataTypeOfArray = -1
Exit Function
End If

If IsArrayEmpty(arr) = True Then
    ' If the array is unallocated, we can still get its data type.
    ' The result of VarType of an array is vbArray + the VarType
    ' of elements of the array (e.g., the VarType of an array of Longs
    ' is 8195, which is vbArray + vbLong). Thus, to get the basic data
    ' type of the array, we subtract the value vbArray.
    DataTypeOfArray = VarType(arr) - vbArray
Else
    ' get the number of dimensions in the array.
    NumDimensions = NumberOfArrayDimensions(arr)
    ' set variable Element to first element of the first dimension
    ' of the array
    If NumDimensions = 1 Then
        If IsObject(arr(LBound(arr))) = True Then
            DataTypeOfArray = vbObject
Exit Function
        End If
        Element = arr(LBound(arr))
    Else
        If IsObject(arr(LBound(arr), 1)) = True Then
            DataTypeOfArray = vbObject
Exit Function
        End If
        Element = arr(LBound(arr), 1)
    End If
    ' if we were passed an array of arrays, IsArray(Element) will
    ' be true. Therefore, return vbArray. If IsArray(Element) is false,
    ' we weren't passed an array of arrays, so simply return the data type of
    ' Element.
    If IsArray(Element) = True Then
        DataTypeOfArray = vbArray
    Else
        If VarType(Element) = vbEmpty Then
            DataTypeOfArray = vbVariant
        Else
            DataTypeOfArray = VarType(Element)
        End If
    End If
End If


End Function

Public Sub SubtractAllElementsWithMatchFromVariantArray(InputArray As Variant, MatchString)
   ' Assumes 1-D Variant array
   Dim TempVariant As Variant
   Dim i As Long
   Dim Bool As Boolean
   
   i = LBound(InputArray)
   
   Do
      If InputArray(i) = MatchString Then
           Bool = DeleteArrayElement(InputArray, i, True)
      Else
        i = i + 1
      End If
   Loop Until i >= UBound(InputArray)
   
   'Now check the UBound element
   If InputArray(UBound(InputArray)) = MatchString Then
        Bool = DeleteArrayElement(InputArray, UBound(InputArray), True)
   End If
   
   
   
End Sub

Public Function DeleteArrayElement(InputArray As Variant, ElementNumber As Long, _
    Optional ResizeDynamic As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DeleteArrayElement
' This function deletes an element from InputArray, and shifts elements that are to the
' right of the deleted element to the left. If InputArray is a dynamic array, and the
' ResizeDynamic parameter is True, the array will be resized one element smaller. Otherwise,
' the right-most entry in the array is set to the default value appropriate to the data
' type of the array (0, vbNullString, Empty, or Nothing). If the array is an array of Variant
' types, the default data type is the data type of the last element in the array.
' The function returns True if the elememt was successfully deleted, or False if an error
' occurrred. This procedure works only on single-dimensional
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long
Dim VType As VbVarType

''''''''''''''''''''''''''''''''
' Set the default result
''''''''''''''''''''''''''''''''
DeleteArrayElement = False

''''''''''''''''''''''''''''''''
' Ensure InputArray is an array.
''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''
' Ensure we have a single dimensional array
'''''''''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(arr:=InputArray) <> 1 Then
Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''
' Ensure we have a valid ElementNumber
''''''''''''''''''''''''''''''''''''''''''''''
If (ElementNumber < LBound(InputArray)) Or (ElementNumber > UBound(InputArray)) Then
Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''
' If we have a single element array, Erase it.
''''''''''''''''''''''''''''''''''''''''''''''
If LBound(InputArray) = UBound(InputArray) Then
    Erase InputArray
Exit Function
End If


''''''''''''''''''''''''''''''''''''''''''''''
' Get the variable data type of the element
' we're deleting.
''''''''''''''''''''''''''''''''''''''''''''''
VType = VarType(InputArray(UBound(InputArray)))
If VType >= vbArray Then
    VType = VType - vbArray
End If
''''''''''''''''''''''''''''''''''''''''''''''
' Shift everything to the left
''''''''''''''''''''''''''''''''''''''''''''''
For Ndx = ElementNumber To UBound(InputArray) - 1
    InputArray(Ndx) = InputArray(Ndx + 1)
Next Ndx
''''''''''''''''''''''''''''''''''''''''''''''
' If ResizeDynamic is True, resize the array
' if it is dynamic.
''''''''''''''''''''''''''''''''''''''''''''''
If IsArrayDynamic(arr:=InputArray) = True Then
    If ResizeDynamic = True Then
        ''''''''''''''''''''''''''''''''
        ' Resize the array and get out.
        ''''''''''''''''''''''''''''''''
        ReDim Preserve InputArray(LBound(InputArray) To UBound(InputArray) - 1)
        DeleteArrayElement = True
Exit Function
    End If
End If
'''''''''''''''''''''''''''''
' Set the last element of the
' InputArray to the proper
' default value.
'''''''''''''''''''''''''''''
Select Case VType
    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbDate, vbCurrency, vbDecimal
        InputArray(UBound(InputArray)) = 0
    Case vbString
        InputArray(UBound(InputArray)) = vbNullString
    Case vbArray, vbVariant, vbEmpty, vbError, vbNull, vbUserDefinedType
        InputArray(UBound(InputArray)) = Empty
    Case vbBoolean
        InputArray(UBound(InputArray)) = False
    Case vbObject
        Set InputArray(UBound(InputArray)) = Nothing
    Case Else
        InputArray(UBound(InputArray)) = 0
End Select

DeleteArrayElement = True

End Function

Public Function FirstNonEmptyStringIndexInArray(InputArray As Variant) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FirstNonEmptyStringIndexInArray
' This returns the index into InputArray of the first non-empty string.
' This is generally used when InputArray is the result of a sort operation,
' which puts empty strings at the beginning of the array.
' Returns -1 is an error occurred or if the entire array is empty strings.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long

If IsArray(InputArray) = False Then
    FirstNonEmptyStringIndexInArray = -1
Exit Function
End If
   
Select Case NumberOfArrayDimensions(arr:=InputArray)
    Case 0
        '''''''''''''''''''''''''''''''''''''''''
        ' indicates an unallocated dynamic array.
        '''''''''''''''''''''''''''''''''''''''''
        FirstNonEmptyStringIndexInArray = -1
Exit Function
    Case 1
        '''''''''''''''''''''''''''''''''''''''''
        ' single dimensional array. OK.
        '''''''''''''''''''''''''''''''''''''''''
    Case Else
        '''''''''''''''''''''''''''''''''''''''''
        ' multidimensional array. Invalid.
        '''''''''''''''''''''''''''''''''''''''''
        FirstNonEmptyStringIndexInArray = -1
Exit Function
End Select

For Ndx = LBound(InputArray) To UBound(InputArray)
    If InputArray(Ndx) <> vbNullString Then
        FirstNonEmptyStringIndexInArray = Ndx
Exit Function
    End If
Next Ndx

FirstNonEmptyStringIndexInArray = -1
End Function


Public Function InsertElementIntoArray(InputArray As Variant, Index As Long, _
    Value As Variant) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' InsertElementIntoArray
' This function inserts an element with a value of Value into InputArray at locatation Index.
' InputArray must be a dynamic array. The Value is stored in location Index, and everything
' to the right of Index is shifted to the right. The array is resized to make room for
' the new element. The value of Index must be greater than or equal to the LBound of
' InputArray and less than or equal to UBound+1. If Index is UBound+1, the Value is
' placed at the end of the array.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long

'''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''
InsertElementIntoArray = False

''''''''''''''''''''''''''''''''
' Ensure InputArray is an array.
''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
Exit Function
End If

''''''''''''''''''''''''''''''''
' Ensure InputArray is dynamic.
''''''''''''''''''''''''''''''''
If IsArrayDynamic(arr:=InputArray) = False Then
Exit Function
End If

'''''''''''''''''''''''''''''''''
' Ensure InputArray is allocated.
'''''''''''''''''''''''''''''''''
If IsArrayAllocated(arr:=InputArray) = False Then
Exit Function
End If

'''''''''''''''''''''''''''''''''
' Ensure InputArray is a single
' dimensional array.
'''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(arr:=InputArray) <> 1 Then
Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''
' Ensure Index is a valid element index.
' We allow Index to be equal to
' UBound + 1 to facilitate inserting
' a value at the end of the array. E.g.,
' InsertElementIntoArray(Arr,UBound(Arr)+1,123)
' will insert 123 at the end of the array.
'''''''''''''''''''''''''''''''''''''''''
If (Index < LBound(InputArray)) Or (Index > UBound(InputArray) + 1) Then
Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''
' Resize the array
'''''''''''''''''''''''''''''''''''''''''''''
ReDim Preserve InputArray(LBound(InputArray) To UBound(InputArray) + 1)
'''''''''''''''''''''''''''''''''''''''''''''
' First, we set the newly created last element
' of InputArray to Value. This is done to trap
' an error 13, type mismatch. This last entry
' will be overwritten when we shift elements
' to the right, and the Value will be inserted
' at Index.
'''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
err.Clear
InputArray(UBound(InputArray)) = Value
If err.Number <> 0 Then
    ''''''''''''''''''''''''''''''''''''''
    ' An error occurred, most likely
    ' an error 13, type mismatch.
    ' Redim the array back to its original
    ' size and exit the function.
    '''''''''''''''''''''''''''''''''''''''
    ReDim Preserve InputArray(LBound(InputArray) To UBound(InputArray) - 1)
Exit Function
End If
'''''''''''''''''''''''''''''''''''''''''''''
' Shift everything to the right.
'''''''''''''''''''''''''''''''''''''''''''''
For Ndx = UBound(InputArray) To Index + 1 Step -1
    InputArray(Ndx) = InputArray(Ndx - 1)
Next Ndx

'''''''''''''''''''''''''''''''''''''''''''''
' Insert Value at Index
'''''''''''''''''''''''''''''''''''''''''''''
InputArray(Index) = Value

    
InsertElementIntoArray = True


End Function



Public Function IsArrayAllDefault(InputArray As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayAllEmpty
' Returns True if the array contains all default values for its
' data type:
'   Variable Type           Value
'   -------------           -------------------
'   Variant                 Empty
'   String                  vbNullString
'   Numeric                 0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long
Dim DefaultValue As Variant

'''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''
IsArrayAllDefault = False

'''''''''''''''''''''''''''''''''''
' Ensure InputArray is an array.
'''''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    IsArrayAllDefault = False
Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure array is allocated. An
' unallocated is considered to be
' all the same type. Return True.
''''''''''''''''''''''''''''''''''
If IsArrayAllocated(arr:=InputArray) = False Then
    IsArrayAllDefault = True
Exit Function
End If
    
''''''''''''''''''''''''''''''''''
' Test the type of variable
''''''''''''''''''''''''''''''''''
Select Case VarType(InputArray)
    Case vbArray + vbVariant
        DefaultValue = Empty
    Case vbArray + vbString
        DefaultValue = vbNullString
    Case Is > vbArray
        DefaultValue = 0
End Select
For Ndx = LBound(InputArray) To UBound(InputArray)
    If IsObject(InputArray(Ndx)) Then
        If Not InputArray(Ndx) Is Nothing Then
Exit Function
        Else
            
        End If
    Else
        If VarType(InputArray(Ndx)) <> vbEmpty Then
            If InputArray(Ndx) <> DefaultValue Then
Exit Function
            End If
        End If
    End If
Next Ndx

'''''''''''''''''''''''''''''''
' If we make it out of the loop,
' the array is all defaults.
' Return True.
'''''''''''''''''''''''''''''''
IsArrayAllDefault = True


End Function



Public Function IsArrayAllNumeric(arr As Variant, _
    Optional AllowNumericStrings As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayAllNumeric
' This function returns True is Arr is entirely numeric. False otherwise. The AllowNumericStrings
' parameter indicates whether strings containing numeric data are considered numeric. If this
' parameter is True, a numeric string is considered a numeric variable. If this parameter is
' omitted or False, a numeric string is not considered a numeric variable.
' Variants that are numeric or Empty are allowed. Variants that are arrays, objects, or
' non-numeric data are not allowed.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long

''''''''''''''''''''''''''''
' Ensure Arr is an array.
''''''''''''''''''''''''''''
If IsArray(arr) = False Then
    IsArrayAllNumeric = False
Exit Function
End If

''''''''''''''''''''''''''''''''''''''
' Ensure Arr is allocated (non-empty).
''''''''''''''''''''''''''''''''''''''
If IsArrayEmpty(arr:=arr) = True Then
    IsArrayAllNumeric = False
Exit Function
End If

''''''''''''''''''''''''''''''''''''''
' Loop through the array.
'''''''''''''''''''''''''''''''''''''
For Ndx = LBound(arr) To UBound(arr)
    Select Case VarType(arr(Ndx))
        Case vbInteger, vbLong, vbDouble, vbSingle, vbCurrency, vbDecimal, vbEmpty
            ' all valid numeric types
        
        Case vbString
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' For strings, check the AllowNumericStrings parameter.
            ' If True and the element is a numeric string, allow it.
            ' If it is a non-numeric string, exit with False.
            ' If AllowNumericStrings is False, all strings, even
            ' numeric strings, will cause a result of False.
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If AllowNumericStrings = True Then
                '''''''''''''''''''''''''''''''''
                ' Allow numeric strings.
                '''''''''''''''''''''''''''''''''
                If IsNumeric(arr(Ndx)) = False Then
                    IsArrayAllNumeric = False
Exit Function
                End If
            Else
                IsArrayAllNumeric = False
Exit Function
            End If
        Case vbVariant
            '''''''''''''''''''''''''''''''''''''''''''''
            ' For Variants, disallow Arrays and Objects.
            ' If the element is not an array or an object,
            ' test whether it is numeric. Allow numeric
            ' Varaints.
            '''''''''''''''''''''''''''''''''''''''''''''
            If IsArray(arr(Ndx)) = True Then
                IsArrayAllNumeric = False
Exit Function
            End If
            If IsObject(arr(Ndx)) = True Then
                IsArrayAllNumeric = False
Exit Function
            End If
            
            If IsNumeric(arr(Ndx)) = False Then
                IsArrayAllNumeric = False
Exit Function
            End If
                
        Case Else
            ' any other data type returns False
            IsArrayAllNumeric = False
Exit Function
    End Select
Next Ndx

IsArrayAllNumeric = True

End Function



Public Function IsArrayAllocated(arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayAllocated
' Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
' sized with Redim) or FALSE if the array is not allocated (a dynamic that has not yet
' been sized with Redim, or a dynamic array that has been Erased). Static arrays are always
' allocated.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
' This function is just the reverse of IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim n As Long
On Error Resume Next

' if Arr is not an array, return FALSE and get out.
If IsArray(arr) = False Then
    IsArrayAllocated = False
Exit Function
End If

' Attempt to get the UBound of the array. If the array has not been allocated,
' an error will occur. Test Err.Number to see if an error occurred.
n = UBound(arr, 1)
If (err.Number = 0) Then
    ''''''''''''''''''''''''''''''''''''''
    ' Under some circumstances, if an array
    ' is not allocated, Err.Number will be
    ' 0. To acccomodate this case, we test
    ' whether LBound <= Ubound. If this
    ' is True, the array is allocated. Otherwise,
    ' the array is not allocated.
    '''''''''''''''''''''''''''''''''''''''
    If LBound(arr) <= UBound(arr) Then
        ' no error. array has been allocated.
        IsArrayAllocated = True
    Else
        IsArrayAllocated = False
    End If
Else
    ' error. unallocated array
    IsArrayAllocated = False
End If

End Function



Public Function IsArrayDynamic(ByRef arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayDynamic
' This function returns TRUE or FALSE indicating whether Arr is a dynamic array.
' Note that if you attempt to ReDim a static array in the same procedure in which it is
' declared, you'll get a compiler error and your code won't run at all.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim LUBound As Long

' If we weren't passed an array, get out now with a FALSE result
If IsArray(arr) = False Then
    IsArrayDynamic = False
Exit Function
End If

' If the array is empty, it hasn't been allocated yet, so we know
' it must be a dynamic array.
If IsArrayEmpty(arr:=arr) = True Then
    IsArrayDynamic = True
Exit Function
End If

' Save the UBound of Arr.
' This value will be used to restore the original UBound if Arr
' is a single-dimensional dynamic array. Unused if Arr is multi-dimensional,
' or if Arr is a static array.
LUBound = UBound(arr)

On Error Resume Next
err.Clear

' Attempt to increase the UBound of Arr and test the value of Err.Number.
' If Arr is a static array, either single- or multi-dimensional, we'll get a
' C_ERR_ARRAY_IS_FIXED_OR_LOCKED error. In this case, return FALSE.
'
' If Arr is a single-dimensional dynamic array, we'll get C_ERR_NO_ERROR error.
'
' If Arr is a multi-dimensional dynamic array, we'll get a
' C_ERR_SUBSCRIPT_OUT_OF_RANGE error.
'
' For either C_NO_ERROR or C_ERR_SUBSCRIPT_OUT_OF_RANGE, return TRUE.
' For C_ERR_ARRAY_IS_FIXED_OR_LOCKED, return FALSE.

ReDim Preserve arr(LBound(arr) To LUBound + 1)

Select Case err.Number
    Case C_ERR_NO_ERROR
        ' We successfully increased the UBound of Arr.
        ' Do a ReDim Preserve to restore the original UBound.
        ReDim Preserve arr(LBound(arr) To LUBound)
        IsArrayDynamic = True
    Case C_ERR_SUBSCRIPT_OUT_OF_RANGE
        ' Arr is a multi-dimensional dynamic array.
        ' Return True.
        IsArrayDynamic = True
    Case C_ERR_ARRAY_IS_FIXED_OR_LOCKED
        ' Arr is a static single- or multi-dimensional array.
        ' Return False
        IsArrayDynamic = False
    Case Else
        ' We should never get here.
        ' Some unexpected error occurred. Be safe and return False.
        IsArrayDynamic = False
End Select

End Function


Public Function IsArrayEmpty(arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayEmpty
' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
' This function is really the reverse of IsArrayAllocated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim LB As Long
Dim UB As Long

err.Clear
On Error Resume Next
If IsArray(arr) = False Then
    ' we weren't passed an array, return True
    IsArrayEmpty = True
End If

' Attempt to get the UBound of the array. If the array is
' unallocated, an error will occur.
UB = UBound(arr, 1)
If (err.Number <> 0) Then
    IsArrayEmpty = True
Else
    ''''''''''''''''''''''''''''''''''''''''''
    ' On rare occassion, under circumstances I
    ' cannot reliably replictate, Err.Number
    ' will be 0 for an unallocated, empty array.
    ' On these occassions, LBound is 0 and
    ' UBoung is -1.
    ' To accomodate the weird behavior, test to
    ' see if LB > UB. If so, the array is not
    ' allocated.
    ''''''''''''''''''''''''''''''''''''''''''
    err.Clear
    LB = LBound(arr)
    If LB > UB Then
        IsArrayEmpty = True
    Else
        IsArrayEmpty = False
    End If
End If

End Function






Public Function IsArrayObjects(InputArray As Variant, _
    Optional AllowNothing As Boolean = True) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayObjects
' Returns True if InputArray is entirely objects (Nothing objects are
' optionally allowed -- default it true, allow Nothing objects). Set the
' AllowNothing to true or false to indicate whether Nothing objects
' are allowed.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long

''''''''''''''''''''''''''''''''
' Set the default return value.
''''''''''''''''''''''''''''''''
IsArrayObjects = False

''''''''''''''''''''''''''''''''
' Ensure InputArray is an array.
''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
Exit Function
End If

'''''''''''''''''''''''''''''''''''''
' Ensure we have a single dimensional
' array.
'''''''''''''''''''''''''''''''''''''
Select Case NumberOfArrayDimensions(arr:=InputArray)
    Case 0
        ''''''''''''''''''''''''''''''''''
        ' Unallocated dynamic array.
        ' Not allowed.
        ''''''''''''''''''''''''''''''''''
Exit Function
    Case 1
         '''''''''''''''''''''''''''''''''
         ' OK
         '''''''''''''''''''''''''''''''''
    Case Else
        '''''''''''''''''''''''''''''''''
        ' Multi-dimensional array.
        ' Not allowed.
        ''''''''''''''''''''''''''''''''
Exit Function
End Select

For Ndx = LBound(InputArray) To UBound(InputArray)
    If IsObject(InputArray(Ndx)) = False Then
Exit Function
    End If
    If InputArray(Ndx) Is Nothing Then
        If AllowNothing = False Then
Exit Function
        End If
    End If
Next Ndx

IsArrayObjects = True

End Function




Public Function IsNumericDataType(TestVar As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsNumericDataType
'
' This function returns TRUE or FALSE indicating whether the data
' type of a variable is a numeric data type. It will return TRUE
' for all of the following data types:
'       vbCurrency
'       vbDecimal
'       vbDouble
'       vbInteger
'       vbLong
'       vbSingle
'
' It will return FALSE for any other data type, including empty Variants and objects.
' If TestVar is an allocated array, it will test data type of the array
' and return TRUE or FALSE for that data type. If TestVar is an allocated
' array, it tests the data type of the first element of the array. If
' TestVar is an array of Variants, the function will indicate only whether
' the first element of the array is numeric. Other elements of the array
' may not be numeric data types. To test an entire array of variants
' to ensure they are all numeric data types, use the IsVariantArrayNumeric
' function.
'
' It will return FALSE for any other data type. Use this procedure
' instead of VBA's IsNumeric function because IsNumeric will return
' TRUE if the variable is a string containing numeric data. This
' will cause problems with code like
'        Dim V1 As Variant
'        Dim V2 As Variant
'        V1 = "1"
'        V2 = "2"
'        If IsNumeric(V1) = True Then
'            If IsNumeric(V2) = True Then
'                Debug.Print  V1 + V2
'            End If
'        End If
'
' The output of the Debug.Print statement will be "12", not 3,
' because V1 and V2 are strings and the '+' operator acts like
' the '&' operator when used with strings. This can lead to
' unexpected results.
'
' IsNumeric should only be used to test strings for numeric content
' when converting a string value to a numeric variable.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Element As Variant
    Dim NumDims As Long
    If IsArray(TestVar) = True Then
        NumDims = NumberOfArrayDimensions(arr:=TestVar)
        If NumDims > 1 Then
            '''''''''''''''''''''''''''''''''
            ' this procedure does not support
            ' multi-dimensional arrays.
            '''''''''''''''''''''''''''''''''
            IsNumericDataType = False
Exit Function
        End If
        If IsArrayAllocated(arr:=TestVar) = True Then
            Element = TestVar(LBound(TestVar))
            Select Case VarType(Element)
                Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
                    IsNumericDataType = True
Exit Function
                Case Else
                    IsNumericDataType = False
Exit Function
            End Select
        Else
            Select Case VarType(TestVar) - vbArray
                Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
                    IsNumericDataType = True
Exit Function
                Case Else
                    IsNumericDataType = False
Exit Function
            End Select
        End If
    End If
    Select Case VarType(TestVar)
        Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
            IsNumericDataType = True
        Case Else
            IsNumericDataType = False
    End Select
End Function



Public Function IsVariantArrayConsistent(arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsVariantArrayConsistent
'
' This returns TRUE or FALSE indicating whether an array of variants
' contains all the same data types. Returns FALSE under the following
' circumstances:
'       Arr is not an array
'       Arr is an array but is unallocated
'       Arr is a multidimensional array
'       Arr is allocated but does not contain consistant data types.
'
' If Arr is an array of objects, objects that are Nothing are ignored.
' As long as all non-Nothing objects are the same object type, the
' function returns True.
'
' It returns TRUE if all the elements of the array have the same
' data type. If Arr is an array of a specific data types, not variants,
' (E.g., Dim V(1 To 3) As Long), the function will return True. If
' an array of variants contains an uninitialized element (VarType =
' vbEmpty) that element is skipped and not used in the comparison. The
' reasoning behind this is that an empty variable will return the
' data type of the variable to which it is assigned (e.g., it will
' return vbNullString to a String and 0 to a Double).
'
' The function does not support arrays of User Defined Types.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim FirstDataType As VbVarType
Dim Ndx As Long
'''''''''''''''''''''''''''''''''''''''''
' Exit with False if Arr is not an array.
'''''''''''''''''''''''''''''''''''''''''
If IsArray(arr) = False Then
    IsVariantArrayConsistent = False
Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''
' Exit with False if Arr is not allocated.
''''''''''''''''''''''''''''''''''''''''''
If IsArrayAllocated(arr) = False Then
    IsVariantArrayConsistent = False
Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''
' Exit with false on multi-dimensional
' arrays.
''''''''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(arr) <> 1 Then
    IsVariantArrayConsistent = False
Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''
' Test if we have an array of a specific
' type rather than Variants. If so,
' return TRUE and get out.
''''''''''''''''''''''''''''''''''''''''''
If (VarType(arr) <= vbArray) And _
    (VarType(arr) <> vbVariant) Then
    IsVariantArrayConsistent = True
Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''
' Get the data type of the first element.
''''''''''''''''''''''''''''''''''''''''''
FirstDataType = VarType(arr(LBound(arr)))
''''''''''''''''''''''''''''''''''''''''''
' Loop through the array and exit if
' a differing data type if found.
''''''''''''''''''''''''''''''''''''''''''
For Ndx = LBound(arr) + 1 To UBound(arr)
    If VarType(arr(Ndx)) <> vbEmpty Then
        If IsObject(arr(Ndx)) = True Then
            If Not arr(Ndx) Is Nothing Then
                If VarType(arr(Ndx)) <> FirstDataType Then
                    IsVariantArrayConsistent = False
Exit Function
                End If
            End If
        Else
            If VarType(arr(Ndx)) <> FirstDataType Then
                IsVariantArrayConsistent = False
Exit Function
            End If
        End If
    End If
Next Ndx

''''''''''''''''''''''''''''''''''''''''''
' If we make it out of the loop,
' then the array is consistent.
''''''''''''''''''''''''''''''''''''''''''
IsVariantArrayConsistent = True

End Function



Public Function IsVariantArrayNumeric(TestArray As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsVariantArrayNumeric
'
' This function return TRUE if all the elements of an array of
' variants are numeric data types. They need not all be the same data
' type. You can have a mix of Integer, Longs, Doubles, and Singles.
' As long as they are all numeric data types, the function will
' return TRUE. If a non-numeric data type is encountered, the
' function will return FALSE. Also, it will return FALSE if
' TestArray is not an array, or if TestArray has not been
' allocated. TestArray may be a multi-dimensional array. This
' procedure uses the IsNumericDataType function to determine whether
' a variable is a numeric data type. If there is an uninitialized
' variant (VarType = vbEmpty) in the array, it is skipped and not
' used in the comparison (i.e., Empty is considered a valid numeric
' data type since you can assign a number to it).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long
Dim DimNdx As Long
Dim NumDims As Long

''''''''''''''''''''''''''''''''
' Ensure TestArray is an array.
''''''''''''''''''''''''''''''''
If IsArray(TestArray) = False Then
    IsVariantArrayNumeric = False
Exit Function
End If
''''''''''''''''''''''''''''''''
' Ensure that TestArray has been
' allocated.
''''''''''''''''''''''''''''''''
If IsArrayAllocated(arr:=TestArray) = False Then
    IsVariantArrayNumeric = False
Exit Function
End If
''''''''''''''''''''''''''''''''''''
' Ensure the array is a one
' dimensional array. This procedure
' will not work on multi-dimensional
' arrays.
''''''''''''''''''''''''''''''''''''
'If NumberOfArrayDimensions(Arr:=TestArray) > 1 Then
'    IsVariantArrayNumeric = False
'    Exit Function
'End If
    
NumDims = NumberOfArrayDimensions(arr:=TestArray)
If NumDims = 1 Then
    '''''''''''''''''''''''''''''''''''
    ' single dimensional array
    '''''''''''''''''''''''''''''''''''
    For Ndx = LBound(TestArray) To UBound(TestArray)
        If IsObject(TestArray(Ndx)) = True Then
            IsVariantArrayNumeric = False
Exit Function
        End If
        
        If VarType(TestArray(Ndx)) <> vbEmpty Then
            If IsNumericDataType(TestVar:=TestArray(Ndx)) = False Then
                IsVariantArrayNumeric = False
Exit Function
            End If
        End If
    Next Ndx
Else
    ''''''''''''''''''''''''''''''''''''
    ' multi-dimensional array
    ''''''''''''''''''''''''''''''''''''
    For DimNdx = 1 To NumDims
        For Ndx = LBound(TestArray, DimNdx) To UBound(TestArray, DimNdx)
            If VarType(TestArray(Ndx, DimNdx)) <> vbEmpty Then
                If IsNumericDataType(TestVar:=TestArray(Ndx, DimNdx)) = False Then
                    IsVariantArrayNumeric = False
Exit Function
                End If
            End If
        Next Ndx
    Next DimNdx
End If

'''''''''''''''''''''''''''''''''''''''
' If we made it out of the loop, then
' the array is entirely numeric.
'''''''''''''''''''''''''''''''''''''''
IsVariantArrayNumeric = True

End Function



Public Function MoveEmptyStringsToEndOfArray(InputArray As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This procedure takes the SORTED array InputArray, which, if sorted in
' ascending order, will have all empty strings at the front of the array.
' This procedure moves those strings to the end of the array, shifting
' the non-empty strings forward in the array.
' Note that InputArray MUST be sorted in ascending order.
' Returns True if the array was correctly shifted (if necessary) and False
' if an error occurred.
' This function uses the following functions, which are included as Private
' procedures at the end of this module.
'       FirstNonEmptyStringIndexInArray
'       NumberOfArrayDimensions
'       IsArrayAllocated

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Temp As String
Dim Ndx As Long
Dim Ndx2 As Long
Dim NonEmptyNdx As Long
Dim FirstNonEmptyNdx As Long


''''''''''''''''''''''''''''''''
' Ensure InpuyArray is an array.
''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    MoveEmptyStringsToEndOfArray = False
Exit Function
End If


''''''''''''''''''''''''''''''''''''
' Ensure that the array is allocated
' (not an empty array).
''''''''''''''''''''''''''''''''''''
If IsArrayAllocated(arr:=InputArray) = False Then
    MoveEmptyStringsToEndOfArray = False
Exit Function
End If


FirstNonEmptyNdx = FirstNonEmptyStringIndexInArray(InputArray:=InputArray)
If FirstNonEmptyNdx <= LBound(InputArray) Then
    ''''''''''''''''''''''''''''''''''''''''''
    ' No empty strings at the beginning of the
    ' array. Get out now.
    ''''''''''''''''''''''''''''''''''''''''''
    MoveEmptyStringsToEndOfArray = True
Exit Function
End If


''''''''''''''''''''''''''''''''''''''''''''''''
' Loop through the array, swapping vbNullStrings
' at the beginning with values at the end.
''''''''''''''''''''''''''''''''''''''''''''''''
NonEmptyNdx = FirstNonEmptyNdx
For Ndx = LBound(InputArray) To UBound(InputArray)
    If InputArray(Ndx) = vbNullString Then
        InputArray(Ndx) = InputArray(NonEmptyNdx)
        InputArray(NonEmptyNdx) = vbNullString
        NonEmptyNdx = NonEmptyNdx + 1
        If NonEmptyNdx > UBound(InputArray) Then
            Exit For
        End If
    End If
Next Ndx
''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set entires (Ndx+1) to UBound(InputArray) to
' vbNullStrings.
''''''''''''''''''''''''''''''''''''''''''''''''''''
For Ndx2 = Ndx + 1 To UBound(InputArray)
    InputArray(Ndx2) = vbNullString
Next Ndx2
MoveEmptyStringsToEndOfArray = True

End Function



Public Function NumberOfArrayDimensions(arr As Variant) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NumberOfArrayDimensions
' This function returns the number of dimensions of an array. An unallocated dynamic array
' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long
Dim Res As Long
On Error Resume Next
' Loop, increasing the dimension index Ndx, until an error occurs.
' An error will occur when Ndx exceeds the number of dimension
' in the array. Return Ndx - 1.
Do
    Ndx = Ndx + 1
    Res = UBound(arr, Ndx)
Loop Until err.Number <> 0

NumberOfArrayDimensions = Ndx - 1
err.Number = 0
err.Source = ""

End Function
 



Public Function NumElements(arr As Variant, Optional Dimension = 1) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NumElements
' Returns the number of elements in the specified dimension (Dimension) of the array in
' Arr. If you omit Dimension, the first dimension is used. The function will return
' 0 under the following circumstances:
'     Arr is not an array, or
'     Arr is an unallocated array, or
'     Dimension is greater than the number of dimension of Arr, or
'     Dimension is less than 1.
'
' This function does not support arrays of user-defined Type variables.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim NumDimensions As Long

' if Arr is not an array, return 0 and get out.
If IsArray(arr) = False Then
    NumElements = 0
Exit Function
End If

' if the array is unallocated, return 0 and get out.
If IsArrayEmpty(arr) = True Then
    NumElements = 0
Exit Function
End If

' ensure that Dimension is at least 1.
If Dimension < 1 Then
    NumElements = 0
Exit Function
End If

' get the number of dimensions
NumDimensions = NumberOfArrayDimensions(arr)
If NumDimensions < Dimension Then
    NumElements = 0
Exit Function
End If

' returns the number of elements in the array
NumElements = UBound(arr, Dimension) - LBound(arr, Dimension) + 1

End Function

Public Function ResetVariantArrayToDefaults(InputArray As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ResetVariantArrayToDefaults
' This resets all the elements of an array of Variants back to their appropriate
' default values. The elements of the array may be of mixed types (e.g., some Longs,
' some Objects, some Strings, etc). Each data type will be set to the appropriate
' default value (0, vbNullString, Empty, or Nothing). It returns True if the
' array was set to defautls, or False if an error occurred. InputArray must be
' an allocated single-dimensional array. This function differs from the Erase
' function in that it preserves the original data types, while Erase sets every
' element to Empty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long
'''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''
ResetVariantArrayToDefaults = False

'''''''''''''''''''''''''''''''
' Ensure InputArray is an array
'''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
Exit Function
End If

'''''''''''''''''''''''''''''''
' Ensure InputArray is a single
' dimensional allocated array.
'''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(arr:=InputArray) <> 1 Then
Exit Function
End If

For Ndx = LBound(InputArray) To UBound(InputArray)
    SetVariableToDefault InputArray(Ndx)
Next Ndx

ResetVariantArrayToDefaults = True

End Function

Public Function FillVariantArray(FillValue As Variant, NumRow As Long, Optional NumColumn As Long, Optional Num3D As Long, Optional Num4D As Long)
Dim i As Long
Dim j As Long
Dim k As Long
Dim L As Long
Dim TempVariant As Variant



If NumRow <= 0 Or NumColumn <= 0 Or Num3D <= 0 Then
   FillVariantArray = Empty
End If


If NumColumn = 0 Then

    ReDim TempVariant(NumRow - 1)
    
    For i = 0 To NumRow - 1
       TempVariant(i, j) = FillValue
    Next

ElseIf Num3D = 0 Then

    ReDim TempVariant(NumRow - 1, NumColumn - 1)
    
    For i = 0 To NumRow - 1
      For j = 0 To NumColumn - 1
          TempVariant(i, j) = FillValue
      Next
    Next

ElseIf Num4D = 0 Then

    ReDim TempVariant(NumRow - 1, NumColumn - 1, Num3D - 1)
    
    For i = 0 To NumRow - 1
      For j = 0 To NumColumn - 1
         For k = 0 To Num3D - 1
              TempVariant(i, j, k) = FillValue
         Next
      Next
    Next

Else

    ReDim TempVariant(NumRow - 1, NumColumn - 1, Num3D - 1, Num4D - 1)
    
    For i = 0 To NumRow - 1
      For j = 0 To NumColumn - 1
         For k = 0 To Num3D - 1
            For L = 0 To Num4D - 1
                 TempVariant(i, j, k, L) = FillValue
            Next
         Next
      Next
    Next

End If 'Out of luck for 5-D and above (extend this function)

FillVariantArray = TempVariant


End Function
 

Public Function ReverseArrayInPlace(InputArray As Variant, _
    Optional NoAlerts As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ReverseArrayInPlace
' This procedure reverses the order of an array in place -- this is, the array variable
' in the calling procedure is reversed. This works only on single-dimensional arrays
' of simple data types (String, Single, Double, Integer, Long). It will not work
' on arrays of objects. Use ReverseArrayOfObjectsInPlace to reverse an array of objects.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Temp As Variant
Dim Ndx As Long
Dim Ndx2 As Long


'''''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''''
ReverseArrayInPlace = False

'''''''''''''''''''''''''''''''''
' ensure we have an array
'''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    If NoAlerts = False Then
        MsgBox "The InputArray parameter is not an array."
    End If
Exit Function
End If

''''''''''''''''''''''''''''''''''''''
' Test the number of dimensions of the
' InputArray. If 0, we have an empty,
' unallocated array. Get out with
' an error message. If greater than
' one, we have a multi-dimensional
' array, which is not allowed. Only
' an allocated 1-dimensional array is
' allowed.
''''''''''''''''''''''''''''''''''''''
Select Case NumberOfArrayDimensions(InputArray)
    Case 0
        If NoAlerts = False Then
            MsgBox "The input array is an empty, unallocated array."
        End If
Exit Function
    Case 1
        ' ok
    Case Else
        If NoAlerts = False Then
            MsgBox "The input array is multi-dimensional. ReverseArrayInPlace works only " & _
                   "on single-dimensional arrays."
        End If
Exit Function
End Select

Ndx2 = UBound(InputArray)
''''''''''''''''''''''''''''''''''''''
' loop from the LBound of InputArray to
' the midpoint of InputArray
''''''''''''''''''''''''''''''''''''''
For Ndx = LBound(InputArray) To ((UBound(InputArray) - LBound(InputArray) + 1) \ 2)
    'swap the elements
    Temp = InputArray(Ndx)
    InputArray(Ndx) = InputArray(Ndx2)
    InputArray(Ndx2) = Temp
    ' decrement the upper index
    Ndx2 = Ndx2 - 1
Next Ndx

''''''''''''''''''''''''''''''''''''''
' OK - Return True
''''''''''''''''''''''''''''''''''''''
ReverseArrayInPlace = True

End Function



Public Function ReverseArrayOfObjectsInPlace(InputArray As Variant, _
    Optional NoAlerts As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ReverseArrayOfObjectsInPlace
' This procedure reverses the order of an array in place -- this is, the array variable
' in the calling procedure is reversed. This works only with arrays of objects. It does
' not work on simple variables. Use ReverseArrayInPlace for simple variables. An error
' will occur if an element of the array is not an object.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Temp As Variant
Dim Ndx As Long
Dim Ndx2 As Long


'''''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''''
ReverseArrayOfObjectsInPlace = False

'''''''''''''''''''''''''''''''''
' ensure we have an array
'''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    If NoAlerts = False Then
        MsgBox "The InputArray parameter is not an array."
    End If
Exit Function
End If

''''''''''''''''''''''''''''''''''''''
' Test the number of dimensions of the
' InputArray. If 0, we have an empty,
' unallocated array. Get out with
' an error message. If greater than
' one, we have a multi-dimensional
' array, which is not allowed. Only
' an allocated 1-dimensional array is
' allowed.
''''''''''''''''''''''''''''''''''''''
Select Case NumberOfArrayDimensions(InputArray)
    Case 0
        If NoAlerts = False Then
            MsgBox "The input array is an empty, unallocated array."
        End If
Exit Function
    Case 1
        ' ok
    Case Else
        If NoAlerts = False Then
            MsgBox "The input array is multi-dimensional. ReverseArrayInPlace works only " & _
                   "on single-dimensional arrays."
        End If
Exit Function
End Select

Ndx2 = UBound(InputArray)

'''''''''''''''''''''''''''''''''''''
' ensure the entire array consists
' of objects (Nothing objects are
' allowed).
'''''''''''''''''''''''''''''''''''''
For Ndx = LBound(InputArray) To UBound(InputArray)
    If IsObject(InputArray(Ndx)) = False Then
        If NoAlerts = False Then
            MsgBox "Array item " & CStr(Ndx) & " is not an object."
        End If
Exit Function
    End If
Next Ndx

''''''''''''''''''''''''''''''''''''''
' loop from the LBound of InputArray to
' the midpoint of InputArray
''''''''''''''''''''''''''''''''''''''
For Ndx = LBound(InputArray) To ((UBound(InputArray) - LBound(InputArray) + 1) \ 2)
    Set Temp = InputArray(Ndx)
    Set InputArray(Ndx) = InputArray(Ndx2)
    Set InputArray(Ndx2) = Temp
    ' decrement the upper index
    Ndx2 = Ndx2 - 1
Next Ndx

''''''''''''''''''''''''''''''''''''''
' OK - Return True
''''''''''''''''''''''''''''''''''''''
ReverseArrayOfObjectsInPlace = True

End Function


Public Function SetObjectArrayToNothing(InputArray As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SetObjectArrrayToNothing
' This sets all the elements of InputArray to Nothing. Use this function
' rather than Erase because if InputArray is an array of Variants, Erase
' will set each element to Empty, not Nothing, and the element will cease
' to be an object.
'
' The function returns True if successful, False otherwise.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim n As Long

''''''''''''''''''''''''''''''''''''''
' Ensure InputArray is an array.
''''''''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    SetObjectArrayToNothing = False
Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''
' Ensure we have a single-dimensional array.
''''''''''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(arr:=InputArray) <> 1 Then
    SetObjectArrayToNothing = False
Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure the array is allocated and that each
' element is an object (or Nothing). If the
' array is not allocated, return True.
' We do this test before setting any element
' to Nothing so we don't end up with an array
' that is a mix of Empty and Nothing values.
' This means looping through the array twice,
' but it ensures all or none of the elements
' get set to Nothing.
''''''''''''''''''''''''''''''''''''''''''''''''
If IsArrayAllocated(arr:=InputArray) = True Then
    For n = LBound(InputArray) To UBound(InputArray)
        If IsObject(InputArray(n)) = False Then
            SetObjectArrayToNothing = False
Exit Function
        End If
    Next n
Else
    SetObjectArrayToNothing = True
Exit Function
End If


'''''''''''''''''''''''''''''''''''''''''''''
' Set each element of InputArray to Nothing.
'''''''''''''''''''''''''''''''''''''''''''''
For n = LBound(InputArray) To UBound(InputArray)
    Set InputArray(n) = Nothing
Next n

SetObjectArrayToNothing = True

End Function

Public Function AreDataTypesCompatible(DestVar As Variant, SourceVar As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' AreDataTypesCompatible
' This function determines if SourceVar is compatiable with DestVar. If the two
' data types are the same, they are compatible. If the value of SourceVar can
' be stored in DestVar with no loss of precision or an overflow, they are compatible.
' For example, if DestVar is a Long and SourceVar is an Integer, they are compatible
' because an integer can be stored in a Long with no loss of information. If DestVar
' is a Long and SourceVar is a Double, they are not compatible because information
' will be lost converting from a Double to a Long (the decimal portion will be lost).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim SVType As VbVarType
Dim DVType As VbVarType

'''''''''''''''''''''''''''''''''''
' Set the default return type.
'''''''''''''''''''''''''''''''''''
AreDataTypesCompatible = False

'''''''''''''''''''''''''''''''''''
' If DestVar is an array, get the
' type of array. If it is an array
' its VarType is vbArray + VarType(element)
' so we subtract vbArray to get then
' data type of the aray. E.g.,
' the VarType of an array of Longs
' is 8195 = vbArray + vbLong,
' 8195 - vbArray = vbLong (=3).
'''''''''''''''''''''''''''''''''''
If IsArray(DestVar) = True Then
    DVType = VarType(DestVar) - vbArray
Else
    DVType = VarType(DestVar)
End If
'''''''''''''''''''''''''''''''''''
' If SourceVar is an array, get the
' type of array.
'''''''''''''''''''''''''''''''''''
If IsArray(SourceVar) = True Then
    SVType = VarType(SourceVar) - vbArray
Else
    SVType = VarType(SourceVar)
End If

''''''''''''''''''''''''''''''''''''
' If one variable is an array and
' the other is not an array, they
' are incompatible.
''''''''''''''''''''''''''''''''''''
If ((IsArray(DestVar) = True) And (IsArray(SourceVar) = False) Or _
    (IsArray(DestVar) = False) And (IsArray(SourceVar) = True)) Then
Exit Function
End If


''''''''''''''''''''''''''''''''''''
' Test the data type of DestVar
' and return a result if SourceVar
' is compatible with that type.
''''''''''''''''''''''''''''''''''''
If SVType = DVType Then
    '''''''''''''''''''''''''''''''''
    ' The the variable types are the
    ' same, they are compatible.
    ''''''''''''''''''''''''''''''''
    AreDataTypesCompatible = True
Exit Function
Else
    '''''''''''''''''''''''''''''''''''''''''
    ' If the data types are not the same,
    ' determine whether they are compatible.
    '''''''''''''''''''''''''''''''''''''''''
    Select Case DVType
        Case vbInteger
            Select Case SVType
                Case vbInteger
                    AreDataTypesCompatible = True
Exit Function
                Case Else
                    AreDataTypesCompatible = False
Exit Function
            End Select
        
        Case vbLong
            Select Case SVType
                Case vbInteger, vbLong
                    AreDataTypesCompatible = True
Exit Function
                Case Else
                    AreDataTypesCompatible = False
Exit Function
            End Select
        Case vbSingle
            Select Case SVType
                Case vbInteger, vbLong, vbSingle
                    AreDataTypesCompatible = True
Exit Function
                Case Else
                    AreDataTypesCompatible = False
Exit Function
            End Select
        Case vbDouble
            Select Case SVType
                Case vbInteger, vbLong, vbSingle, vbDouble
                    AreDataTypesCompatible = True
Exit Function
                Case Else
                    AreDataTypesCompatible = False
Exit Function
            End Select
        Case vbString
            Select Case SVType
                Case vbString
                    AreDataTypesCompatible = True
Exit Function
                Case Else
                    AreDataTypesCompatible = False
Exit Function
            End Select
        Case vbObject
            Select Case SVType
                Case vbObject
                    AreDataTypesCompatible = True
Exit Function
                Case Else
                    AreDataTypesCompatible = False
Exit Function
            End Select
        Case vbBoolean
            Select Case SVType
                Case vbBoolean, vbInteger
                    AreDataTypesCompatible = True
Exit Function
                Case Else
                    AreDataTypesCompatible = False
Exit Function
            End Select
         Case vbByte
            Select Case SVType
                Case vbByte
                    AreDataTypesCompatible = True
Exit Function
                Case Else
                    AreDataTypesCompatible = False
Exit Function
            End Select
        Case vbCurrency
            Select Case SVType
                Case vbInteger, vbLong, vbSingle, vbDouble
                    AreDataTypesCompatible = True
Exit Function
                Case Else
                    AreDataTypesCompatible = False
Exit Function
            End Select
        Case vbDecimal
            Select Case SVType
                Case vbInteger, vbLong, vbSingle, vbDouble
                    AreDataTypesCompatible = True
Exit Function
                Case Else
                    AreDataTypesCompatible = False
Exit Function
            End Select
        Case vbDate
            Select Case SVType
                Case vbLong, vbSingle, vbDouble
                    AreDataTypesCompatible = True
Exit Function
                Case Else
                    AreDataTypesCompatible = False
Exit Function
            End Select
        
         Case vbEmpty
            Select Case SVType
                Case vbVariant
                    AreDataTypesCompatible = True
Exit Function
                Case Else
                    AreDataTypesCompatible = False
Exit Function
            End Select
         Case vbError
            AreDataTypesCompatible = False
Exit Function
         Case vbNull
            AreDataTypesCompatible = False
Exit Function
         Case vbObject
            Select Case SVType
                Case vbObject
                    AreDataTypesCompatible = True
Exit Function
                Case Else
                    AreDataTypesCompatible = False
Exit Function
            End Select
         Case vbVariant
            AreDataTypesCompatible = True
Exit Function
        
    End Select
End If


End Function

Public Sub SetVariableToDefault(ByRef Variable As Variant)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SetVariableToDefault
' This procedure sets Variable to the appropriate default
' value for its data type. Note that it cannot change User-Defined
' Types.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IsObject(Variable) Then
    '''''''''''''''''''''''''''''''''''''''
    ' We test with IsObject here so that
    ' the object itself, not the default
    ' property of the object, is evaluated.
    ''''''''''''''''''''''''''''''''''''''''
    Set Variable = Nothing
Else
    Select Case VarType(Variable)
        Case Is >= vbArray
            ''''''''''''''''''''''''''''''''''''''''''''
            ' The VarType of an array is
            ' equal to vbArray + VarType(ArrayElement).
            ' Here we check for anything >= vbArray
            ''''''''''''''''''''''''''''''''''''''''''''
            Erase Variable
        Case vbBoolean
            Variable = False
        Case vbByte
            Variable = CByte(0)
        Case vbCurrency
            Variable = CCur(0)
        Case vbDataObject
            Set Variable = Nothing
        Case vbDate
            Variable = CDate(0)
        Case vbDecimal
            Variable = CDec(0)
        Case vbDouble
            Variable = CDbl(0)
        Case vbEmpty
            Variable = Empty
        Case vbError
            Variable = Empty
        Case vbInteger
            Variable = CInt(0)
        Case vbLong
            Variable = CLng(0)
        Case vbNull
            Variable = Empty
        Case vbObject
            Set Variable = Nothing
        Case vbSingle
            Variable = CSng(0)
        Case vbString
            Variable = vbNullString
        Case vbUserDefinedType
            '''''''''''''''''''''''''''''''''
            ' User-Defined-Types cannot be
            ' set to a general default value.
            ' Each element must be explicitly
            ' set to its default value. No
            ' assignment takes place in this
            ' procedure.
            ''''''''''''''''''''''''''''''''''
        Case vbVariant
            ''''''''''''''''''''''''''''''''''''''''''''''''
            ' This case is included for constistancy,
            ' but we will never get here. If the Variant
            ' contains data, VarType returns the type of
            ' that data. An Empty Variant is type vbEmpty.
            ''''''''''''''''''''''''''''''''''''''''''''''''
            Variable = Empty
    End Select
End If

End Sub

Public Function TransposeArray(InputArr As Variant, OutputArr As Variant) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' TransposeArray
' This transposes a two-dimensional array. It returns True if successful or
' False if an error occurs. InputArr must be two-dimensions. OutputArr must be
' a dynamic array. It will be Erased and resized, so any existing content will
' be destroyed.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim RowNdx As Long
Dim ColNdx As Long
Dim LB1 As Long
Dim LB2 As Long
Dim UB1 As Long
Dim UB2 As Long

'''''''''''''''''''''''''''''''''''
' Ensure InputArr and OutputArr
' are arrays.
'''''''''''''''''''''''''''''''''''
If (IsArray(InputArr) = False) Or (IsArray(OutputArr) = False) Then
    TransposeArray = False
Exit Function
End If

'''''''''''''''''''''''''''''''''''
' Ensure OutputArr is a dynamic
' array.
'''''''''''''''''''''''''''''''''''
If IsArrayDynamic(arr:=OutputArr) = False Then
    TransposeArray = False
Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure InputArr is two-dimensions,
' no more, no lesss.
''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(arr:=InputArr) <> 2 Then
    TransposeArray = False
Exit Function
End If

'''''''''''''''''''''''''''''''''''''''
' Get the Lower and Upper bounds of
' InputArr.
'''''''''''''''''''''''''''''''''''''''
LB1 = LBound(InputArr, 1)
LB2 = LBound(InputArr, 2)
UB1 = UBound(InputArr, 1)
UB2 = UBound(InputArr, 2)

'''''''''''''''''''''''''''''''''''''''''
' Erase and ReDim OutputArr
'''''''''''''''''''''''''''''''''''''''''
Erase OutputArr
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Redim the Output array. Not the that the LBound and UBound
' values are preserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReDim OutputArr(LB2 To LB2 + UB2 - LB2, LB1 To LB1 + UB1 - LB1)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Loop through the elemetns of InputArr and put each value
' in the proper element of the tranposed array.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For RowNdx = LBound(InputArr, 2) To UBound(InputArr, 2)
    For ColNdx = LBound(InputArr, 1) To UBound(InputArr, 1)
        OutputArr(RowNdx, ColNdx) = InputArr(ColNdx, RowNdx)
    Next ColNdx
Next RowNdx

'''''''''''''''''''''''''
' Success -- return True.
'''''''''''''''''''''''''
TransposeArray = True

End Function


Public Function VectorsToArray(arr As Variant, ParamArray Vectors()) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' VectorsToArray
' This function takes 1 or more single-dimensional arrays and converts
' them into a single multi-dimensional array. Each array in Vectors
' comprises one row of the new array. The number of columns in the
' new array is the maximum of the number of elements in each vector.
' Arr MUST be a dynamic array of a data type compatible with ALL the
' elements in each Vector. The code does NOT trap for an error
' 13 - Type Mismatch.
'
' If the Vectors are of differing sizes, Arr is sized to hold the
' maximum number of elements in a Vector. The procedure Erases the
' Arr array, so when it is reallocated with Redim, all elements will
' be the reset to their default value (0 or vbNullString or Empty).
' Unused elements in the new array will remain the default value for
' that data type.
'
' Each Vector in Vectors must be a single dimensional array, but
' the Vectors may be of different sizes and LBounds.
'
' Each element in each Vector must be a simple data type. The elements
' may NOT be Object, Arrays, or User-Defined Types.
'
' The rows and columns of the result array are 0-based, regardless of
' the LBound of each vector and regardless of the Option Base statement.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Vector As Variant
Dim VectorNdx As Long
Dim NumElements As Long
Dim NumRows As Long
Dim NumCols As Long
Dim RowNdx As Long
Dim ColNdx As Long
Dim VType As VbVarType

'''''''''''''''''''''''''''''''''''
' Ensure we have an Array
''''''''''''''''''''''''''''''''''
If IsArray(arr) = False Then
    VectorsToArray = False
Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure we have a dynamic array
''''''''''''''''''''''''''''''''''
If IsArrayDynamic(arr) = False Then
    VectorsToArray = False
Exit Function
End If
    
'''''''''''''''''''''''''''''''''
' Ensure that at least one vector
' was passed in Vectors
'''''''''''''''''''''''''''''''''
If IsMissing(Vectors) = True Then
    VectorsToArray = False
Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''''
' Loop through Vectors to determine the
' size of the result array. We do this
' loop first to prevent having to do
' a Redim Preserve. This requires looping
' through Vectors a second time, but this
' is still faster than doing Redim Preserves.
'''''''''''''''''''''''''''''''''''''''''''''''
For Each Vector In Vectors
    ''''''''''''''''''''''''''''
    ' Ensure Vector is single
    ' dimensional array. This
    ' will take care of the case
    ' if Vector is an unallocated
    ' array (NumberOfArrayDimensions = 0
    ' for an unallocated array).
    ''''''''''''''''''''''''''''
    If NumberOfArrayDimensions(Vector) <> 1 Then
        VectorsToArray = False
Exit Function
    End If
    '''''''''''''''''''''''''''''''''''''
    ' Ensure that Vector is not an array.
    '''''''''''''''''''''''''''''''''''''
    If IsArray(Vector) = False Then
        VectorsToArray = False
Exit Function
    End If
    '''''''''''''''''''''''''''''''''
    ' Increment the number of rows.
    ' Each Vector is one row or the
    ' result array. Test the size
    ' of Vector. If it is larger
    ' than the existing value of
    ' NumCols, set NumCols to the
    ' new, larger, value.
    '''''''''''''''''''''''''''''''''
    NumRows = NumRows + 1
    If NumCols < UBound(Vector) - LBound(Vector) + 1 Then
        NumCols = UBound(Vector) - LBound(Vector) + 1
    End If
Next Vector
''''''''''''''''''''''''''''''''''''''''''''
' Redim Arr to the appropriate size. Arr
' is 0-based in both directions, regardless
' of the LBound of the original Arr and
' regardless of the LBounds of the Vectors.
''''''''''''''''''''''''''''''''''''''''''''
ReDim arr(0 To NumRows - 1, 0 To NumCols - 1)

'''''''''''''''''''''''''''''''
' Loop row-by-row.
For RowNdx = 0 To NumRows - 1
    ''''''''''''''''''''''''''''''''
    ' Loop through the columns.
    ''''''''''''''''''''''''''''''''
    For ColNdx = 0 To NumCols - 1
        ''''''''''''''''''''''''''''
        ' Set Vector (a Variant) to
        ' the Vectors(RowNdx) array.
        ' We declare Vector as a
        ' variant so it can take an
        ' array of any simple data
        ' type.
        ''''''''''''''''''''''''''''
        Vector = Vectors(RowNdx)
        '''''''''''''''''''''''''''''
        ' The vectors need not ber
        If ColNdx < UBound(Vector) - LBound(Vector) + 1 Then
            VType = VarType(Vector(LBound(Vector) + ColNdx))
            If VType >= vbArray Then
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Test for VType >= vbArray. The VarType of an array
                ' is vbArray + VarType(element of array). E.g., the
                ' VarType of an array of Longs equal vbArray + vbLong.
                ' Anything greater than or equal to vbArray is an
                ' array of some time.
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                VectorsToArray = False
Exit Function
            End If
            If VType = vbObject Then
                VectorsToArray = False
Exit Function
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Vector(LBound(Vector) + ColNdx) is
            ' a simple data type. If Vector(LBound(Vector) + ColNdx)
            ' is not a compatible data type with Arr, then a Type
            ' Mismatch error will occur. We do NOT trap this error.
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            arr(RowNdx, ColNdx) = Vector(LBound(Vector) + ColNdx)
        End If
    Next ColNdx
Next RowNdx

VectorsToArray = True

End Function

Public Function ChangeBoundsOfArray(InputArr As Variant, _
    NewLowerBound As Long, NewUpperBound, Optional Dimension As Long = 1) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ChangeBoundsOfArray
' This function changes the upper and lower bounds of the specified
' array. InputArr MUST be a single-dimensional dynamic array.
' If the new size of the array (NewUpperBound - NewLowerBound + 1)
' is greater than the original array, the unused elements on
' right side of the array are the default values for the data type
' of the array. If the new size is less than the original size,
' only the first (left-most) N elements are included in the new array.
' The elements of the array may be simple variables (Strings, Longs, etc)
' Object, or Arrays. User-Defined Types are not supported.
'
' The function returns True if successful, False otherwise.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim TempArr() As Variant
Dim InNdx As Long
Dim OutNdx As Long
Dim TempNdx As Long
Dim FirstIsObject As Boolean

''''''''''''''''''''''''''''''''''''
' Ensure we have an array.
''''''''''''''''''''''''''''''''''''
If IsArray(InputArr) = False Then
    ChangeBoundsOfArray = False
Exit Function
End If
''''''''''''''''''''''''''''''''''''
' Ensure the array is dynamic.
''''''''''''''''''''''''''''''''''''
If IsArrayDynamic(InputArr) = False Then
    ChangeBoundsOfArray = False
Exit Function
End If
''''''''''''''''''''''''''''''''''''
' Ensure the array is allocated.
''''''''''''''''''''''''''''''''''''
If IsArrayAllocated(InputArr) = False Then
    ChangeBoundsOfArray = False
Exit Function
End If
'''''''''''''''''''''''''''''''''''''''''''
' Ensure the NewLowerBound > NewUpperBound.
'''''''''''''''''''''''''''''''''''''''''''
If NewLowerBound > NewUpperBound Then
    ChangeBoundsOfArray = False
Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''''
' Ensure Arr is a single dimensional array.
'''''''''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(InputArr) > 2 Then
    ChangeBoundsOfArray = False
Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''
' We need to save the IsObject status of the first
' element of the InputArr to properly handle
' the Empty variables is we are making the array
' larger than it was before.
'''''''''''''''''''''''''''''''''''''''''''''''''''
FirstIsObject = IsObject(InputArr(LBound(InputArr)))


''''''''''''''''''''''''''''''''''''''''''''
' Resize TempArr and save the values in
' InputArr in TempArr. TempArr will have
' an LBound of 1 and a UBound of the size
' of (NewUpperBound - NewLowerBound +1)
'''''''''''''''''''''''''''''''''''''''''''
ReDim TempArr(1 To (NewUpperBound - NewLowerBound + 1))
'''''''''''''''''''''''''''''''''''''''''''
' Load up TempArr
'''''''''''''''''''''''''''''''''''''''''''
TempNdx = 0
For InNdx = LBound(InputArr) To UBound(InputArr)
    TempNdx = TempNdx + 1
    If TempNdx > UBound(TempArr) Then
        Exit For
    End If
    
    If (IsObject(InputArr(InNdx)) = True) Then
        If InputArr(InNdx) Is Nothing Then
            Set TempArr(TempNdx) = Nothing
        Else
            Set TempArr(TempNdx) = InputArr(InNdx)
        End If
    Else
        TempArr(TempNdx) = InputArr(InNdx)
    End If
Next InNdx

''''''''''''''''''''''''''''''''''''
' Now, Erase InputArr, resize it to the
' new bounds, and load up the values from
' TempArr to the new InputArr.
''''''''''''''''''''''''''''''''''''
Erase InputArr
ReDim InputArr(NewLowerBound To NewUpperBound)
OutNdx = LBound(InputArr)
For TempNdx = LBound(TempArr) To UBound(TempArr)
    If OutNdx <= UBound(InputArr) Then
        If IsObject(TempArr(TempNdx)) = True Then
            Set InputArr(OutNdx) = TempArr(TempNdx)
        Else
            If FirstIsObject = True Then
                If IsEmpty(TempArr(TempNdx)) = True Then
                    Set InputArr(OutNdx) = Nothing
                Else
                    Set InputArr(OutNdx) = TempArr(TempNdx)
                End If
            Else
                InputArr(OutNdx) = TempArr(TempNdx)
            End If
        End If
    Else
        Exit For
    End If
    OutNdx = OutNdx + 1
Next TempNdx

'''''''''''''''''''''''''''''
' Success -- Return True
'''''''''''''''''''''''''''''
ChangeBoundsOfArray = True


End Function


Public Function IsArraySorted(TestArray As Variant, _
    Optional Descending As Boolean = False) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArraySorted
' This function determines whether a single-dimensional array is sorted. Because
' sorting is an expensive operation, especially so on large array of Variants,
' you may want to determine if an array is already in sorted order prior to
' doing an actual sort.
' This function returns True if an array is in sorted order (either ascending or
' descending order, depending on the value of the Descending parameter -- default
' is false = Ascending). The decision to do a string comparison (with StrComp) or
' a numeric comparison (with < or >) is based on the data type of the first
' element of the array.
' If TestArray is not an array, is an unallocated dynamic array, or has more than
' one dimension, or the VarType of TestArray is not compatible, the function
' returns NULL.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim StrCompResultFail As Long
Dim NumericResultFail As Boolean
Dim Ndx As Long
Dim NumCompareResult As Boolean
Dim StrCompResult As Long

Dim IsString As Boolean
Dim VType As VbVarType

''''''''''''''''''''''''''''''''''
' Ensure TestArray is an array.
''''''''''''''''''''''''''''''''''
If IsArray(TestArray) = False Then
    IsArraySorted = Null
Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''
' Ensure we have a single dimensional array.
''''''''''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(arr:=TestArray) <> 1 Then
    IsArraySorted = Null
Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''
' The following code sets the values of
' comparison that will indicate that the
' array is unsorted. It the result of
' StrComp (for strings) or ">=" (for
' numerics) equals the value specified
' below, we know that the array is
' unsorted.
'''''''''''''''''''''''''''''''''''''''''''''
If Descending = True Then
    StrCompResultFail = -1
    NumericResultFail = False
Else
    StrCompResultFail = 1
    NumericResultFail = True
End If

''''''''''''''''''''''''''''''''''''''''''''''
' Determine whether we are going to do a string
' comparison or a numeric comparison.
''''''''''''''''''''''''''''''''''''''''''''''
VType = VarType(TestArray(LBound(TestArray)))
Select Case VType
    Case vbArray, vbDataObject, vbEmpty, vbError, vbNull, vbObject, vbUserDefinedType
    '''''''''''''''''''''''''''''''''
    ' Unsupported types. Reutrn Null.
    '''''''''''''''''''''''''''''''''
        IsArraySorted = Null
Exit Function
    Case vbString, vbVariant
    '''''''''''''''''''''''''''''''''
    ' Compare as string
    '''''''''''''''''''''''''''''''''
        IsString = True
    Case Else
    '''''''''''''''''''''''''''''''''
    ' Compare as numeric
    '''''''''''''''''''''''''''''''''
        IsString = False
End Select

For Ndx = LBound(TestArray) To UBound(TestArray) - 1
    If IsString = True Then
        StrCompResult = StrComp(TestArray(Ndx), TestArray(Ndx + 1))
        If StrCompResult = StrCompResultFail Then
            IsArraySorted = False
Exit Function
        End If
    Else
        NumCompareResult = (TestArray(Ndx) >= TestArray(Ndx + 1))
        If NumCompareResult = NumericResultFail Then
            IsArraySorted = False
Exit Function
        End If
    End If
Next Ndx


''''''''''''''''''''''''''''
' If we made it out of  the
' loop, then the array is
' in sorted order. Return
' True.
''''''''''''''''''''''''''''
IsArraySorted = True
End Function


Public Function CombineTwoDArrays(Arr1 As Variant, _
    Arr2 As Variant) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' TwoArraysToOneArray
' This takes two 2-dimensional arrays, Arr1 and Arr2, and
' returns an array combining the two. The number of Rows
' in the result is NumRows(Arr1) + NumRows(Arr2). Arr1 and
' Arr2 must have the same number of columns, and the result
' array will have that many columns. All the LBounds must
' be the same. E.g.,
' The following arrays are legal:
'        Dim Arr1(0 To 4, 0 To 10)
'        Dim Arr2(0 To 3, 0 To 10)
'
' The following arrays are illegal
'        Dim Arr1(0 To 4, 1 To 10)
'        Dim Arr2(0 To 3, 0 To 10)
'
' The returned result array is Arr1 with additional rows
' appended from Arr2. For example, the arrays
'    a    b        and     e    f
'    c    d                g    h
' become
'    a    b
'    c    d
'    e    f
'    g    h
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''
' Upper and lower bounds of Arr1.
'''''''''''''''''''''''''''''''''
Dim LBoundRow1 As Long
Dim UBoundRow1 As Long
Dim LBoundCol1 As Long
Dim UBoundCol1 As Long

'''''''''''''''''''''''''''''''''
' Upper and lower bounds of Arr2.
'''''''''''''''''''''''''''''''''
Dim LBoundRow2 As Long
Dim UBoundRow2 As Long
Dim LBoundCol2 As Long
Dim UBoundCol2 As Long

'''''''''''''''''''''''''''''''''''
' Upper and lower bounds of Result.
'''''''''''''''''''''''''''''''''''
Dim LBoundRowResult As Long
Dim UBoundRowResult As Long
Dim LBoundColResult As Long
Dim UBoundColResult As Long

'''''''''''''''''
' Index Variables
'''''''''''''''''
Dim RowNdx1 As Long
Dim ColNdx1 As Long
Dim RowNdx2 As Long
Dim ColNdx2 As Long
Dim RowNdxResult As Long
Dim ColNdxResult As Long


'''''''''''''
' Array Sizes
'''''''''''''
Dim NumRows1 As Long
Dim NumCols1 As Long

Dim NumRows2 As Long
Dim NumCols2 As Long

Dim NumRowsResult As Long
Dim NumColsResult As Long

Dim Done As Boolean
Dim Result() As Variant
Dim ResultTrans() As Variant

Dim V As Variant


'''''''''''''''''''''''''''''''
' Ensure that Arr1 and Arr2 are
' arrays.
''''''''''''''''''''''''''''''
If (IsArray(Arr1) = False) Or (IsArray(Arr2) = False) Then
    CombineTwoDArrays = Null
Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure both arrays are allocated
' two dimensional arrays.
''''''''''''''''''''''''''''''''''
If (NumberOfArrayDimensions(Arr1) <> 2) Or (NumberOfArrayDimensions(Arr2) <> 2) Then
    CombineTwoDArrays = Null
Exit Function
End If
    
'''''''''''''''''''''''''''''''''''''''
' Ensure that the LBound and UBounds
' of the second dimension are the
' same for both Arr1 and Arr2.
'''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''
' Get the existing bounds.
''''''''''''''''''''''''''
LBoundRow1 = LBound(Arr1, 1)
UBoundRow1 = UBound(Arr1, 1)

LBoundCol1 = LBound(Arr1, 2)
UBoundCol1 = UBound(Arr1, 2)

LBoundRow2 = LBound(Arr2, 1)
UBoundRow2 = UBound(Arr2, 1)

LBoundCol2 = LBound(Arr2, 2)
UBoundCol2 = UBound(Arr2, 2)

''''''''''''''''''''''''''''''''''''''''''''''''''
' Get the total number of rows for the result
' array.
''''''''''''''''''''''''''''''''''''''''''''''''''
NumRows1 = UBoundRow1 - LBoundRow1 + 1
NumCols1 = UBoundCol1 - LBoundCol1 + 1
NumRows2 = UBoundRow2 - LBoundRow2 + 1
NumCols2 = UBoundCol2 - LBoundCol2 + 1

'''''''''''''''''''''''''''''''''''''''''
' Ensure the number of columns are equal.
'''''''''''''''''''''''''''''''''''''''''
If NumCols1 <> NumCols2 Then
    CombineTwoDArrays = Null
Exit Function
End If

NumRowsResult = NumRows1 + NumRows2

'''''''''''''''''''''''''''''''''''''''
' Ensure that ALL the LBounds are equal.
''''''''''''''''''''''''''''''''''''''''
If (LBoundRow1 <> LBoundRow2) Or _
    (LBoundRow1 <> LBoundCol1) Or _
    (LBoundRow1 <> LBoundCol2) Then
    CombineTwoDArrays = Null
Exit Function
End If
'''''''''''''''''''''''''''''''
' Get the LBound of the columns
' of the result array.
'''''''''''''''''''''''''''''''
LBoundColResult = LBoundRow1
'''''''''''''''''''''''''''''''
' Get the UBound of the columns
' of the result array.
'''''''''''''''''''''''''''''''
UBoundColResult = UBoundCol1

UBoundRowResult = LBound(Arr1, 1) + NumRows1 + NumRows2 - 1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Redim the Result array to have number of rows equal to
' number-of-rows(Arr1) + number-of-rows(Arr2)
' and number-of-columns equal to number-of-columns(Arr1)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReDim Result(LBoundRow1 To UBoundRowResult, LBoundColResult To UBoundColResult)

RowNdxResult = LBound(Result, 1) - 1

Done = False
Do Until Done
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Copy elements of Arr1 to Result
    ''''''''''''''''''''''''''''''''''''''''''''
    For RowNdx1 = LBound(Arr1, 1) To UBound(Arr1, 1)
        RowNdxResult = RowNdxResult + 1
        For ColNdx1 = LBound(Arr1, 2) To UBound(Arr1, 2)
            V = Arr1(RowNdx1, ColNdx1)
            Result(RowNdxResult, ColNdx1) = V
        Next ColNdx1
    Next RowNdx1

    '''''''''''''''''''''''''''''''''''''''''''''
    ' Copy elements of Arr2 to Result
    '''''''''''''''''''''''''''''''''''''''''''''
    For RowNdx2 = LBound(Arr2, 1) To UBound(Arr2, 1)
        RowNdxResult = RowNdxResult + 1
        For ColNdx2 = LBound(Arr2, 2) To UBound(Arr2, 2)
            V = Arr2(RowNdx2, ColNdx2)
            Result(RowNdxResult, ColNdx2) = V
        Next ColNdx2
    Next RowNdx2
    
    If RowNdxResult >= UBound(Result, 1) + (LBoundColResult = 1) Then
        Done = True
    End If
'''''''''''''
' End Of Loop
'''''''''''''
Loop
'''''''''''''''''''''''''
' Return the Result
'''''''''''''''''''''''''
CombineTwoDArrays = Result

End Function


'Daniel Villa created this function based on "ExpandArray"
Public Function AddRowOrColumnToEnd(ByVal arr As Variant, WhichDim As Long, _
        VectorToAdd As Variant) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ExpandArray
' This expands a two-dimensional array in either dimension. It returns the result
' array if successful, or NULL if an error occurred. The original array is never
' changed.
' Paramters:
' --------------------
' Arr                   is the array to be expanded.
'
' WhichDim              is either 1 for additional rows or 2 for
'                       additional columns.
'
' AdditionalElements    is the number of additional rows or columns
'                       to create.
'
' VectorToAdd             is the value to which the new array elements should be
'                       initialized.
'
' You can nest calls to Expand array to expand both the number of rows and
' columns. E.g.,
'
' C = ExpandArray(ExpandArray(Arr:=A, WhichDim:=1, AdditionalElements:=3, VectorToAdd:="R"), _
'    WhichDim:=2, AdditionalElements:=4, VectorToAdd:="C")
' This first adds three rows at the bottom of the array, and then adds four
' columns on the right of the array.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Result As Variant
Dim RowNdx As Long
Dim ColNdx As Long
Dim ResultRowNdx As Long
Dim ResultColNdx As Long
Dim NumRows As Long
Dim NumCols As Long
Dim NewUBound As Long
Dim AdditionalElements As Long
Dim FillIndex As Long
Dim TempVariant As Variant
Dim i As Long

Const ROWS_ As Long = 1
Const COLS_ As Long = 2
AdditionalElements = 1

'''''''''''''''''''''''''''''''''
' Ensure the dimension is 1 or 2.
'''''''''''''''''''''''''''''''''
Select Case WhichDim
    Case 1, 2
    Case Else
        AddRowOrColumnToEnd = Null
Exit Function
End Select

''''''''''''''''''''''''''''
' Ensure Arr is an array.
''''''''''''''''''''''''''''
If IsArray(arr) = False Then
    AddRowOrColumnToEnd = Null
Exit Function
End If
'''''''''''''''''''''''''''''''''
' Ensure Arr has two dimenesions. Allow 1 dimension but switch to 2-D based on WhichDim
'''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(arr:=arr) <> 2 Then
    If NumberOfArrayDimensions(arr) = 1 Then
     ' convert to a 2-D array
      TempVariant = arr
      If WhichDim = 1 Then
            ReDim arr(0, LBound(TempVariant) To UBound(TempVariant))
            For i = LBound(TempVariant) To UBound(TempVariant)
                arr(0, i) = TempVariant(i)
            Next
      Else
            ReDim arr(LBound(TempVariant) To UBound(TempVariant), 0)
            For i = LBound(TempVariant) To UBound(TempVariant)
                arr(i, 0) = TempVariant(i)
            Next
      End If
    
    Else
       AddRowOrColumnToEnd = Null
Exit Function
    End If
End If

' Ensure that the sizes are compatible
If ArrayMod_shared.NumberOfArrayDimensions(VectorToAdd) <> 1 Then
        AddRowOrColumnToEnd = Null
Exit Function
End If

If WhichDim = ROWS_ Then
    If UBound(VectorToAdd) - LBound(VectorToAdd) <> UBound(arr, 2) - LBound(arr, 2) Then
        AddRowOrColumnToEnd = Null
Exit Function
    End If
Else
    If UBound(VectorToAdd) - LBound(VectorToAdd) <> UBound(arr, 2) - LBound(arr, 2) Then
        AddRowOrColumnToEnd = Null
Exit Function
    End If
End If

''''''''''''''''''''''''''''''''''''
' Ensure AdditionalElements is > 0.
' If AdditionalElements  < 0, return NULL.
' If AdditionalElements  = 0, return Arr.
''''''''''''''''''''''''''''''''''''
If AdditionalElements < 0 Then
    AddRowOrColumnToEnd = Null
Exit Function
End If
If AdditionalElements = 0 Then
    AddRowOrColumnToEnd = arr
Exit Function
End If
    
NumRows = UBound(arr, 1) - LBound(arr, 1) + 1
NumCols = UBound(arr, 2) - LBound(arr, 2) + 1
   
If WhichDim = ROWS_ Then
    '''''''''''''''
    ' Redim Result.
    '''''''''''''''
    ReDim Result(LBound(arr, 1) To UBound(arr, 1) + AdditionalElements, LBound(arr, 2) To UBound(arr, 2))
    ''''''''''''''''''''''''''''''
    ' Transfer Arr array to Result
    ''''''''''''''''''''''''''''''
    For RowNdx = LBound(arr, 1) To UBound(arr, 1)
        For ColNdx = LBound(arr, 2) To UBound(arr, 2)
            Result(RowNdx, ColNdx) = arr(RowNdx, ColNdx)
        Next ColNdx
    Next RowNdx
    '''''''''''''''''''''''''''''''
    ' Fill the rest of the result
    ' array with VectorToAdd.
    '''''''''''''''''''''''''''''''
    FillIndex = LBound(VectorToAdd)
    For RowNdx = UBound(arr, 1) + 1 To UBound(Result, 1)
        For ColNdx = LBound(arr, 2) To UBound(arr, 2)
            
            Result(RowNdx, ColNdx) = VectorToAdd(FillIndex)
            FillIndex = FillIndex + 1
        Next ColNdx
    Next RowNdx
Else
    '''''''''''''''
    ' Redim Result.
    '''''''''''''''
    ReDim Result(LBound(arr, 1) To UBound(arr, 1), UBound(arr, 2) + AdditionalElements)
    ''''''''''''''''''''''''''''''
    ' Transfer Arr array to Result
    ''''''''''''''''''''''''''''''
    For RowNdx = LBound(arr, 1) To UBound(arr, 1)
        For ColNdx = LBound(arr, 2) To UBound(arr, 2)
            Result(RowNdx, ColNdx) = arr(RowNdx, ColNdx)
        Next ColNdx
    Next RowNdx
    '''''''''''''''''''''''''''''''
    ' Fill the rest of the result
    ' array with VectorToAdd.
    '''''''''''''''''''''''''''''''
    FillIndex = LBound(VectorToAdd)
    For RowNdx = LBound(arr, 1) To UBound(arr, 1)
        For ColNdx = UBound(arr, 2) + 1 To UBound(Result, 2)
            Result(RowNdx, ColNdx) = VectorToAdd(FillIndex)
            FillIndex = FillIndex + 1
        Next ColNdx
    Next RowNdx
    
End If
''''''''''''''''''''
' Return the result.
''''''''''''''''''''
AddRowOrColumnToEnd = Result

End Function

Function SwapArrayRows(arr As Variant, Row1 As Long, Row2 As Long) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SwapArrayRows
' This function returns an array based on Arr with Row1 and Row2 swapped.
' It returns the result array or NULL if an error occurred.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim V As Variant
Dim Result As Variant
Dim RowNdx As Long
Dim ColNdx As Long

'''''''''''''''''''''''''
' Ensure Arr is an array.
'''''''''''''''''''''''''
If IsArray(arr) = False Then
    SwapArrayRows = Null
Exit Function
End If

''''''''''''''''''''''''''''''''
' Set Result to Arr
''''''''''''''''''''''''''''''''
Result = arr

''''''''''''''''''''''''''''''''
' Ensure Arr is two-dimensional.
''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(arr:=arr) <> 2 Then
    SwapArrayRows = Null
Exit Function
End If

''''''''''''''''''''''''''''''''
' Ensure Row1 and Row2 are less
' than or equal to the number of
' rows.
''''''''''''''''''''''''''''''''
If (Row1 > UBound(arr, 1)) Or (Row2 > UBound(arr, 1)) Then
    SwapArrayRows = Null
Exit Function
End If
    
'''''''''''''''''''''''''''''''''
' If Row1 = Row2, just return the
' array and exit. Nothing to do.
'''''''''''''''''''''''''''''''''
If Row1 = Row2 Then
    SwapArrayRows = arr
Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''
' Redim V to the number of columns.
'''''''''''''''''''''''''''''''''''''''''
ReDim V(LBound(arr, 2) To UBound(arr, 2))
'''''''''''''''''''''''''''''''''''''''''
' Put Row1 in V
'''''''''''''''''''''''''''''''''''''''''
For ColNdx = LBound(arr, 2) To UBound(arr, 2)
    V(ColNdx) = arr(Row1, ColNdx)
    Result(Row1, ColNdx) = arr(Row2, ColNdx)
    Result(Row2, ColNdx) = V(ColNdx)
Next ColNdx

SwapArrayRows = Result

End Function


Function SwapArrayColumns(arr As Variant, Col1 As Long, Col2 As Long) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SwapArrayColumns
' This function returns an array based on Arr with Col1 and Col2 swapped.
' It returns the result array or NULL if an error occurred.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim V As Variant
Dim Result As Variant
Dim RowNdx As Long
Dim ColNdx As Long

'''''''''''''''''''''''''
' Ensure Arr is an array.
'''''''''''''''''''''''''
If IsArray(arr) = False Then
    SwapArrayColumns = Null
Exit Function
End If

''''''''''''''''''''''''''''''''
' Set Result to Arr
''''''''''''''''''''''''''''''''
Result = arr

''''''''''''''''''''''''''''''''
' Ensure Arr is two-dimensional.
''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(arr:=arr) <> 2 Then
    SwapArrayColumns = Null
Exit Function
End If

''''''''''''''''''''''''''''''''
' Ensure Row1 and Row2 are less
' than or equal to the number of
' rows.
''''''''''''''''''''''''''''''''
If (Col1 > UBound(arr, 2)) Or (Col2 > UBound(arr, 2)) Then
    SwapArrayColumns = Null
Exit Function
End If
    
'''''''''''''''''''''''''''''''''
' If Col1 = Col2, just return the
' array and exit. Nothing to do.
'''''''''''''''''''''''''''''''''
If Col1 = Col2 Then
    SwapArrayColumns = arr
Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''
' Redim V to the number of columns.
'''''''''''''''''''''''''''''''''''''''''
ReDim V(LBound(arr, 1) To UBound(arr, 1))
'''''''''''''''''''''''''''''''''''''''''
' Put Col2 in V
'''''''''''''''''''''''''''''''''''''''''
For RowNdx = LBound(arr, 1) To UBound(arr, 1)
    V(RowNdx) = arr(RowNdx, Col1)
    Result(RowNdx, Col1) = arr(RowNdx, Col2)
    Result(RowNdx, Col2) = V(RowNdx)
Next RowNdx

SwapArrayColumns = Result

End Function

Function GetColumn(arr As Variant, ResultArr As Variant, ColumnNumber As Long) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetColumn
' This populates ResultArr with a one-dimensional array that is the
' specified column of Arr. The existing contents of ResultArr are
' destroyed. ResultArr must be a dynamic array.
' Returns True or False indicating success.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RowNdx As Long
''''''''''''''''''''''''''''''
' Ensure Arr is an array.
''''''''''''''''''''''''''''''
If IsArray(arr) = False Then
    GetColumn = False
Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure Arr is a two-dimensional
' array.
''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(arr) <> 2 Then
    GetColumn = False
Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure ResultArr is a dynamic
' array.
''''''''''''''''''''''''''''''''''
If IsArrayDynamic(ResultArr) = False Then
    GetColumn = False
Exit Function
End If

''''''''''''''''''''''''''''''''''''
' Ensure ColumnNumber is less than
' or equal to the number of columns.
''''''''''''''''''''''''''''''''''''
If UBound(arr, 2) < ColumnNumber Then
    GetColumn = False
Exit Function
End If
If LBound(arr, 2) > ColumnNumber Then
    GetColumn = False
Exit Function
End If

Erase ResultArr
ReDim ResultArr(LBound(arr, 1) To UBound(arr, 1))
For RowNdx = LBound(ResultArr) To UBound(ResultArr)
    ResultArr(RowNdx) = arr(RowNdx, ColumnNumber)
Next RowNdx

GetColumn = True


End Function


Function GetRow(arr As Variant, ResultArr As Variant, RowNumber As Long) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetRow
' This populates ResultArr with a one-dimensional array that is the
' specified row of Arr. The existing contents of ResultArr are
' destroyed. ResultArr must be a dynamic array.
' Returns True or False indicating success.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim ColNdx As Long
''''''''''''''''''''''''''''''
' Ensure Arr is an array.
''''''''''''''''''''''''''''''
If IsArray(arr) = False Then
    GetRow = False
Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure Arr is a two-dimensional
' array.
''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(arr) <> 2 Then
    GetRow = False
Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure ResultArr is a dynamic
' array.
''''''''''''''''''''''''''''''''''
If IsArrayDynamic(ResultArr) = False Then
    GetRow = False
Exit Function
End If

''''''''''''''''''''''''''''''''''''
' Ensure ColumnNumber is less than
' or equal to the number of columns.
''''''''''''''''''''''''''''''''''''
If UBound(arr, 1) < RowNumber Then
    GetRow = False
Exit Function
End If
If LBound(arr, 1) > RowNumber Then
    GetRow = False
Exit Function
End If

Erase ResultArr
ReDim ResultArr(LBound(arr, 2) To UBound(arr, 2))
For ColNdx = LBound(ResultArr) To UBound(ResultArr)
    ResultArr(ColNdx) = arr(RowNumber, ColNdx)
Next ColNdx

GetRow = True


End Function

Private Function ModuleBase() As Long
'''''''''''''''''''''''''''''''''''''''''''''
' ModuleBase
' This returns either 0 or 1, indicating the
' array base for this module. You should copy
' this procedure into the module in which it
' will be used.
'''''''''''''''''''''''''''''''''''''''''''''
Dim arr(1) As Long
If UBound(arr) - LBound(arr) + 1 = 2 Then
    ModuleBase = 0
Else
    ModuleBase = 1
End If


End Function


Public Function TestIfItWorks(n As Variant) As Variant

  TestIfItWorks = n ^ 2 + 22 * n + 7

End Function

Sub CreateSequentialArray(BeginNum As Long, EndNum As Long, OutputArray As Variant)
Dim i As Long
Erase OutputArray
ReDim OutputArray(0 To EndNum - BeginNum)
For i = BeginNum To EndNum
   OutputArray(i - BeginNum) = i
Next
End Sub

Public Sub SortIntegerArray(arr() As Long)
'Best for < 100 items to be sorted
  Dim intTemp As String
  Dim i As Long
  Dim j As Long
  Dim lngMin As Long
  Dim lngMax As Long
  Dim TempInd() As Long
  Dim UB As Long
  Dim LB As Long
  Dim IndTemp As Long
  lngMin = LBound(arr)
  lngMax = UBound(arr) '
   
  ReDim TempInd(lngMin To lngMax)
  For i = lngMin To lngMax
      TempInd(i) = i
  Next
  
  For i = lngMin To lngMax - 1
    For j = i + 1 To lngMax
      If arr(i) > arr(j) Then
        intTemp = arr(i)
        arr(i) = arr(j)
        arr(j) = intTemp
        'Hold onto the indices too added by dlvilla
        IndTemp = TempInd(i)
        TempInd(i) = TempInd(j)
        TempInd(j) = IndTemp
      End If
    Next j
    
  Next i



End Sub

Public Function BinarySearchForIntegers(IntArr() As Long, Point As Long, Optional AcceptClosestMatch As Boolean = False) As Long

  ' Search, it is assumed that IntArray grows monotonically
  ' It assumes an array with non-negative indices
  Dim BPos As Long
  Dim EPos As Long
  Dim BInd As Long
  Dim EInd As Long
  Dim MInd As Long
  Dim MaxIter As Long
  Dim Iter As Long
  
  MaxIter = 100
  BInd = LBound(IntArr)
  EInd = UBound(IntArr)
  
  BPos = IntArr(BInd)
  EPos = IntArr(EInd)
  
  If BPos = EPos Then
     MsgBox "ArrayMod_shared.FindInterpolatePoint: zero length array input which does not work!!!  "
     BinarySearchForIntegers = -1
Exit Function
  ElseIf Point > EPos Then
     MsgBox "ArrayMod_shared.FindInterpolatePoint: Point sought " & Point & " is greater than the input array range! " & BPos & " to " & EPos & "!"
     BinarySearchForIntegers = -1
Exit Function
  ElseIf Point < BPos Then
     MsgBox "ArrayMod_shared.FindInterpolatePoint: Point sought " & Point & " is less than the input array range " & BPos & " to " & EPos & "!"
     BinarySearchForIntegers = -1
Exit Function
  End If
   
'  Begin binary search
  Iter = 0
  Do
    MInd = Int((BInd + EInd) / 2)
    If EInd - BInd = 1 Then
       If IntArr(EInd) = Point Then
          MInd = EInd
          Exit Do
       ElseIf IntArr(BInd) = Point Then
          MInd = BInd
          Exit Do
       End If
       MInd = BInd
    End If
    
    If IntArr(MInd) > Point Then
        EInd = MInd
    ElseIf IntArr(MInd) < Point Then
        BInd = MInd
    ElseIf IntArr(MInd) = Point Then
        Exit Do
    End If
    Iter = Iter + 1
  Loop Until Iter > MaxIter
  
  If Iter > MaxIter Then
     If AcceptClosestMatch Then
          BinarySearchForIntegers = MInd
      Else
          BinarySearchForIntegers = -1
      End If
   Else
      BinarySearchForIntegers = MInd
   End If
  
End Function


' BEGIN OF CODE FROM http://www.cpearson.com/excel/SortingArrays.aspx
Public Function QSortInPlace( _
    ByRef InputArray As Variant, _
    Optional ByVal LB As Long = -1&, _
    Optional ByVal UB As Long = -1&, _
    Optional ByVal Descending As Boolean = False, _
    Optional ByVal CompareMode As VbCompareMethod = vbTextCompare, _
    Optional ByVal NoAlerts As Boolean = False) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' QSortInPlace
'
' This function sorts the array InputArray in place -- this is, the original array in the
' calling procedure is sorted. It will work with either string data or numeric data.
' It need not sort the entire array. You can sort only part of the array by setting the LB and
' UB parameters to the first (LB) and last (UB) element indexes that you want to sort.
' LB and UB are optional parameters. If omitted LB is set to the LBound of InputArray, and if
' omitted UB is set to the UBound of the InputArray. If you want to sort the entire array,
' omit the LB and UB parameters, or set both to -1, or set LB = LBound(InputArray) and set
' UB to UBound(InputArray).
'
' By default, the sort method is case INSENSTIVE (case doens't matter: "A", "b", "C", "d").
' To make it case SENSITIVE (case matters: "A" "C" "b" "d"), set the CompareMode argument
' to vbBinaryCompare (=0). If Compare mode is omitted or is any value other than vbBinaryCompare,
' it is assumed to be vbTextCompare and the sorting is done case INSENSITIVE.
'
' The function returns TRUE if the array was successfully sorted or FALSE if an error
' occurred. If an error occurs (e.g., LB > UB), a message box indicating the error is
' displayed. To suppress message boxes, set the NoAlerts parameter to TRUE.
'
''''''''''''''''''''''''''''''''''''''
' MODIFYING THIS CODE:
''''''''''''''''''''''''''''''''''''''
' If you modify this code and you call "Exit Procedure", you MUST decrment the RecursionLevel
' variable. E.g.,
'       If SomethingThatCausesAnExit Then
'           RecursionLevel = RecursionLevel - 1
'           Exit Function
'       End If
'''''''''''''''''''''''''''''''''''''''
'
' Note: If you coerce InputArray to a ByVal argument, QSortInPlace will not be
' able to reference the InputArray in the calling procedure and the array will
' not be sorted.
'
' This function uses the following procedures. These are declared as Private procedures
' at the end of this module:
'       IsArrayAllocated
'       IsSimpleDataType
'       IsSimpleNumericType
'       QSortCompare
'       NumberOfArrayDimensions
'       ReverseArrayInPlace
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Temp As Variant
Dim Buffer As Variant
Dim CurLow As Long
Dim CurHigh As Long
Dim CurMidpoint As Long
Dim Ndx As Long
Dim pCompareMode As VbCompareMethod

'''''''''''''''''''''''''
' Set the default result.
'''''''''''''''''''''''''
QSortInPlace = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This variable is used to determine the level
' of recursion  (the function calling itself).
' RecursionLevel is incremented when this procedure
' is called, either initially by a calling procedure
' or recursively by itself. The variable is decremented
' when the procedure exits. We do the input parameter
' validation only when RecursionLevel is 1 (when
' the function is called by another function, not
' when it is called recursively).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Static RecursionLevel As Long


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Keep track of the recursion level -- that is, how many
' times the procedure has called itself.
' Carry out the validation routines only when this
' procedure is first called. Don't run the
' validations on a recursive call to the
' procedure.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
RecursionLevel = RecursionLevel + 1

If RecursionLevel = 1 Then
    ''''''''''''''''''''''''''''''''''
    ' Ensure InputArray is an array.
    ''''''''''''''''''''''''''''''''''
    If IsArray(InputArray) = False Then
        If NoAlerts = False Then
            MsgBox "The InputArray parameter is not an array."
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' InputArray is not an array. Exit with a False result.
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        RecursionLevel = RecursionLevel - 1
Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Test LB and UB. If < 0 then set to LBound and UBound
    ' of the InputArray.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If LB < 0 Then
        LB = LBound(InputArray)
    End If
    If UB < 0 Then
        UB = UBound(InputArray)
    End If
    
    Select Case NumberOfArrayDimensions(InputArray)
        Case 0
            ''''''''''''''''''''''''''''''''''''''''''
            ' Zero dimensions indicates an unallocated
            ' dynamic array.
            ''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The InputArray is an empty, unallocated array."
            End If
            RecursionLevel = RecursionLevel - 1
Exit Function
        Case 1
            ''''''''''''''''''''''''''''''''''''''''''
            ' We sort ONLY single dimensional arrays.
            ''''''''''''''''''''''''''''''''''''''''''
        Case Else
            ''''''''''''''''''''''''''''''''''''''''''
            ' We sort ONLY single dimensional arrays.
            ''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The InputArray is multi-dimensional." & _
                      "QSortInPlace works only on single-dimensional arrays."
            End If
            RecursionLevel = RecursionLevel - 1
Exit Function
    End Select
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Ensure that InputArray is an array of simple data
    ' types, not other arrays or objects. This tests
    ' the data type of only the first element of
    ' InputArray. If InputArray is an array of Variants,
    ' subsequent data types may not be simple data types
    ' (e.g., they may be objects or other arrays), and
    ' this may cause QSortInPlace to fail on the StrComp
    ' operation.
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    If IsSimpleDataType(InputArray(LBound(InputArray))) = False Then
        If NoAlerts = False Then
            MsgBox "InputArray is not an array of simple data types."
            RecursionLevel = RecursionLevel - 1
Exit Function
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ensure that the LB parameter is valid.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case LB
        Case Is < LBound(InputArray)
            If NoAlerts = False Then
                MsgBox "The LB lower bound parameter is less than the LBound of the InputArray"
            End If
            RecursionLevel = RecursionLevel - 1
Exit Function
        Case Is > UBound(InputArray)
            If NoAlerts = False Then
                MsgBox "The LB lower bound parameter is greater than the UBound of the InputArray"
            End If
            RecursionLevel = RecursionLevel - 1
Exit Function
        Case Is > UB
            If NoAlerts = False Then
                MsgBox "The LB lower bound parameter is greater than the UB upper bound parameter."
            End If
            RecursionLevel = RecursionLevel - 1
Exit Function
    End Select

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ensure the UB parameter is valid.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case UB
        Case Is > UBound(InputArray)
            If NoAlerts = False Then
                MsgBox "The UB upper bound parameter is greater than the upper bound of the InputArray."
            End If
            RecursionLevel = RecursionLevel - 1
Exit Function
        Case Is < LBound(InputArray)
            If NoAlerts = False Then
                MsgBox "The UB upper bound parameter is less than the lower bound of the InputArray."
            End If
            RecursionLevel = RecursionLevel - 1
Exit Function
        Case Is < LB
            If NoAlerts = False Then
                MsgBox "the UB upper bound parameter is less than the LB lower bound parameter."
            End If
            RecursionLevel = RecursionLevel - 1
Exit Function
    End Select

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' if UB = LB, we have nothing to sort, so get out.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If UB = LB Then
        QSortInPlace = True
        RecursionLevel = RecursionLevel - 1
Exit Function
    End If

End If ' RecursionLevel = 1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure that CompareMode is either vbBinaryCompare  or
' vbTextCompare. If it is neither, default to vbTextCompare.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If (CompareMode = vbBinaryCompare) Or (CompareMode = vbTextCompare) Then
    pCompareMode = CompareMode
Else
    pCompareMode = vbTextCompare
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Begin the actual sorting process.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CurLow = LB
CurHigh = UB

If LB = 0 Then
    CurMidpoint = ((LB + UB) \ 2) + 1
Else
    CurMidpoint = (LB + UB) \ 2 ' note integer division (\) here
End If
Temp = InputArray(CurMidpoint)

Do While (CurLow <= CurHigh)
    
    Do While QSortCompare(V1:=InputArray(CurLow), V2:=Temp, CompareMode:=pCompareMode) < 0
        CurLow = CurLow + 1
        If CurLow = UB Then
            Exit Do
        End If
    Loop
    
    Do While QSortCompare(V1:=Temp, V2:=InputArray(CurHigh), CompareMode:=pCompareMode) < 0
        CurHigh = CurHigh - 1
        If CurHigh = LB Then
           Exit Do
        End If
    Loop

    If (CurLow <= CurHigh) Then
        Buffer = InputArray(CurLow)
        InputArray(CurLow) = InputArray(CurHigh)
        InputArray(CurHigh) = Buffer
        CurLow = CurLow + 1
        CurHigh = CurHigh - 1
    End If
Loop

If LB < CurHigh Then
    QSortInPlace InputArray:=InputArray, LB:=LB, UB:=CurHigh, _
        Descending:=Descending, CompareMode:=pCompareMode, NoAlerts:=True
End If

If CurLow < UB Then
    QSortInPlace InputArray:=InputArray, LB:=CurLow, UB:=UB, _
        Descending:=Descending, CompareMode:=pCompareMode, NoAlerts:=True
End If

'''''''''''''''''''''''''''''''''''''
' If Descending is True, reverse the
' order of the array, but only if the
' recursion level is 1.
'''''''''''''''''''''''''''''''''''''
If Descending = True Then
    If RecursionLevel = 1 Then
        ReverseArrayInPlace2 InputArray, LB, UB
    End If
End If

RecursionLevel = RecursionLevel - 1
QSortInPlace = True
End Function

Public Function QSortCompare(V1 As Variant, V2 As Variant, _
    Optional CompareMode As VbCompareMethod = vbTextCompare) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' QSortCompare
' This function is used in QSortInPlace to compare two elements. If
' V1 AND V2 are both numeric data types (integer, long, single, double)
' they are converted to Doubles and compared. If V1 and V2 are BOTH strings
' that contain numeric data, they are converted to Doubles and compared.
' If either V1 or V2 is a string and does NOT contain numeric data, both
' V1 and V2 are converted to Strings and compared with StrComp.
'
' The result is -1 if V1 < V2,
'                0 if V1 = V2
'                1 if V1 > V2
' For text comparisons, case sensitivity is controlled by CompareMode.
' If this is vbBinaryCompare, the result is case SENSITIVE. If this
' is omitted or any other value, the result is case INSENSITIVE.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim d1 As Double
Dim d2 As Double
Dim S1 As String
Dim S2 As String

Dim Compare As VbCompareMethod
''''''''''''''''''''''''''''''''''''''''''''''''
' Test CompareMode. Any value other than
' vbBinaryCompare will default to vbTextCompare.
''''''''''''''''''''''''''''''''''''''''''''''''
If CompareMode = vbBinaryCompare Or CompareMode = vbTextCompare Then
    Compare = CompareMode
Else
    Compare = vbTextCompare
End If
'''''''''''''''''''''''''''''''''''''''''''''''
' If either V1 or V2 is either an array or
' an Object, raise a error 13 - Type Mismatch.
'''''''''''''''''''''''''''''''''''''''''''''''
If IsArray(V1) = True Or IsArray(V2) = True Then
    err.Raise 13
Exit Function
End If
If IsObject(V1) = True Or IsObject(V2) = True Then
    err.Raise 13
Exit Function
End If

If IsSimpleNumericType(V1) = True Then
    If IsSimpleNumericType(V2) = True Then
        '''''''''''''''''''''''''''''''''''''
        ' If BOTH V1 and V2 are numeric data
        ' types, then convert to Doubles and
        ' do an arithmetic compare and
        ' return the result.
        '''''''''''''''''''''''''''''''''''''
        d1 = CDbl(V1)
        d2 = CDbl(V2)
        If d1 = d2 Then
            QSortCompare = 0
Exit Function
        End If
        If d1 < d2 Then
            QSortCompare = -1
Exit Function
        End If
        If d1 > d2 Then
            QSortCompare = 1
Exit Function
        End If
    End If
End If
''''''''''''''''''''''''''''''''''''''''''''
' Either V1 or V2 was not numeric data type.
' Test whether BOTH V1 AND V2 are numeric
' strings. If BOTH are numeric, convert to
' Doubles and do a arithmetic comparison.
''''''''''''''''''''''''''''''''''''''''''''
If IsNumeric(V1) = True And IsNumeric(V2) = True Then
    d1 = CDbl(V1)
    d2 = CDbl(V2)
    If d1 = d2 Then
        QSortCompare = 0
Exit Function
    End If
    If d1 < d2 Then
        QSortCompare = -1
Exit Function
    End If
    If d1 > d2 Then
        QSortCompare = 1
Exit Function
    End If
End If
''''''''''''''''''''''''''''''''''''''''''''''
' Either or both V1 and V2 was not numeric
' string. In this case, convert to Strings
' and use StrComp to compare.
''''''''''''''''''''''''''''''''''''''''''''''
S1 = CStr(V1)
S2 = CStr(V2)
QSortCompare = StrComp(S1, S2, Compare)

End Function


Public Function ReverseArrayInPlace2(InputArray As Variant, _
    Optional LB As Long = -1, Optional UB As Long = -1, _
    Optional NoAlerts As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ReverseArrayInPlace2
' This reverses the order of elements in InputArray. To reverse the entire array, omit or
' set to less than 0 the LB and UB parameters. To reverse only part of tbe array, set LB and/or
' UB to the LBound and UBound of the sub array to be reversed.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim n As Long
Dim Temp As Variant
Dim Ndx As Long
Dim Ndx2 As Long
Dim OrigN As Long
Dim NewN As Long
Dim NewArr() As Variant

''''''''''''''''''''''''''''''''
' Set the default return value.
''''''''''''''''''''''''''''''''
ReverseArrayInPlace2 = False

'''''''''''''''''''''''''''''''''
' Ensure we have an array
'''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    If NoAlerts = False Then
        MsgBox "The InputArray parameter is not an array."
    End If
Exit Function
End If

''''''''''''''''''''''''''''''''''''''
' Test the number of dimensions of the
' InputArray. If 0, we have an empty,
' unallocated array. Get out with
' an error message. If greater than
' one, we have a multi-dimensional
' array, which is not allowed. Only
' an allocated 1-dimensional array is
' allowed.
''''''''''''''''''''''''''''''''''''''
Select Case NumberOfArrayDimensions(InputArray)
    Case 0
        '''''''''''''''''''''''''''''''''''''''''''
        ' Zero dimensions indicates an unallocated
        ' dynamic array.
        '''''''''''''''''''''''''''''''''''''''''''
        If NoAlerts = False Then
            MsgBox "The input array is an empty, unallocated array."
        End If
Exit Function
    Case 1
        '''''''''''''''''''''''''''''''''''''''''''
        ' We can reverse ONLY a single dimensional
        ' arrray.
        '''''''''''''''''''''''''''''''''''''''''''
    Case Else
        '''''''''''''''''''''''''''''''''''''''''''
        ' We can reverse ONLY a single dimensional
        ' arrray.
        '''''''''''''''''''''''''''''''''''''''''''
        If NoAlerts = False Then
            MsgBox "The input array multi-dimensional. ReverseArrayInPlace works only " & _
                   "on single-dimensional arrays."
        End If
Exit Function

End Select

'''''''''''''''''''''''''''''''''''''''''''''
' Ensure that we have only simple data types,
' not an array of objects or arrays.
'''''''''''''''''''''''''''''''''''''''''''''
If IsSimpleDataType(InputArray(LBound(InputArray))) = False Then
    If NoAlerts = False Then
        MsgBox "The input array contains arrays, objects, or other complex data types." & vbCrLf & _
            "ReverseArrayInPlace can reverse only arrays of simple data types."
Exit Function
    End If
End If

If LB < 0 Then
    LB = LBound(InputArray)
End If
If UB < 0 Then
    UB = UBound(InputArray)
End If

For n = LB To (LB + ((UB - LB - 1) \ 2))
    Temp = InputArray(n)
    InputArray(n) = InputArray(UB - (n - LB))
    InputArray(UB - (n - LB)) = Temp
Next n

ReverseArrayInPlace2 = True
End Function


Public Function IsSimpleNumericType(V As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsSimpleNumericType
' This returns TRUE if V is one of the following data types:
'        vbBoolean
'        vbByte
'        vbCurrency
'        vbDate
'        vbDecimal
'        vbDouble
'        vbInteger
'        vbLong
'        vbSingle
'        vbVariant if it contains a numeric value
' It returns FALSE for any other data type, including any array
' or vbEmpty.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IsSimpleDataType(V) = True Then
    Select Case VarType(V)
        Case vbBoolean, _
                vbByte, _
                vbCurrency, _
                vbDate, _
                vbDecimal, _
                vbDouble, _
                vbInteger, _
                vbLong, _
                vbSingle
            IsSimpleNumericType = True
        Case vbVariant
            If IsNumeric(V) = True Then
                IsSimpleNumericType = True
            Else
                IsSimpleNumericType = False
            End If
        Case Else
            IsSimpleNumericType = False
    End Select
Else
    IsSimpleNumericType = False
End If
End Function

Public Function IsSimpleDataType(V As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsSimpleDataType
' This function returns TRUE if V is one of the following
' variable types (as returned by the VarType function:
'    vbBoolean
'    vbByte
'    vbCurrency
'    vbDate
'    vbDecimal
'    vbDouble
'    vbEmpty
'    vbError
'    vbInteger
'    vbLong
'    vbNull
'    vbSingle
'    vbString
'    vbVariant
'
' It returns FALSE if V is any one of the following variable
' types:
'    vbArray
'    vbDataObject
'    vbObject
'    vbUserDefinedType
'    or if it is an array of any type.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Test if V is an array. We can't just use VarType(V) = vbArray
' because the VarType of an array is vbArray + VarType(type
' of array element). E.g, the VarType of an Array of Longs is
' 8195 = vbArray + vbLong.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IsArray(V) = True Then
    IsSimpleDataType = False
Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' We must also explicitly check whether V is an object, rather
' relying on VarType(V) to equal vbObject. The reason is that
' if V is an object and that object has a default proprety, VarType
' returns the data type of the default property. For example, if
' V is an Excel.Range object pointing to cell A1, and A1 contains
' 12345, VarType(V) would return vbDouble, the since Value is
' the default property of an Excel.Range object and the default
' numeric type of Value in Excel is Double. Thus, in order to
' prevent this type of behavior with default properties, we test
' IsObject(V) to see if V is an object.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IsObject(V) = True Then
    IsSimpleDataType = False
Exit Function
End If
'''''''''''''''''''''''''''''''''''''
' Test the value returned by VarType.
'''''''''''''''''''''''''''''''''''''
Select Case VarType(V)
    Case vbArray, vbDataObject, vbObject, vbUserDefinedType, vbArray + vbVariant
        '''''''''''''''''''''''
        ' not simple data types
        '''''''''''''''''''''''
        IsSimpleDataType = False
    Case Else
        ''''''''''''''''''''''''''''''''''''
        ' otherwise it is a simple data type
        ''''''''''''''''''''''''''''''''''''
        IsSimpleDataType = True
End Select

End Function

Public Function SumVariantVector(Var As Variant) As Double

   Dim Temp As Double
   Dim i As Long
   Temp = 0
   If IsVariantArrayNumeric(Var) And ArrayMod_shared.NumberOfArrayDimensions(Var) = 1 Then
      For i = LBound(Var) To UBound(Var)
          Temp = Temp + Var(i)
      Next
      SumVariantVector = Temp
   Else
      MsgBox "ArrayMod_shared.SumVariantVector: You must input a variant vector of numerics for this function to work!" & _
             " -1 has been assigned to the output."
      SumVariantVector = -1
   End If
End Function

Public Function Subtract1FromVectorElements(ByVal Vec As Variant, ByVal Ind As Long, ByVal Val As Long)
       Dim i As Long
       
       For i = Ind To UBound(Vec)
          If Val <= Vec(i) Then
             Vec(i) = Vec(i) - 1
          End If
       Next
       
       Subtract1FromVectorElements = Vec
End Function

Public Sub ReOrderVector(arr As Variant, Ind As Variant)
   Dim i As Long
   Dim TempArr As Variant
   Dim IndLB As Long
   Dim VLB As Long
   
   'Make the types match
   TempArr = arr
   
   If NumberOfArrayDimensions(arr) <> 1 Or NumberOfArrayDimensions(Ind) <> 1 Or NumElements(Ind) <> NumElements(arr) Then
      MsgBox "ArrayMod_shared.ReOrderVector: Both Inputs must be vectors with type variant and the same length!"
   Else
        IndLB = LBound(Ind) '
        VLB = LBound(arr)
        For i = VLB To UBound(arr)
             TempArr(i) = arr(Ind(IndLB + (i - VLB)))
        Next
        
        arr = TempArr
   End If
   
   
 
End Sub

Public Sub CLngArr(VarArr As Variant, LngArr() As Long, Optional ConversionSucceeded As Boolean = True)
' This function takes a variant array and converts it to a string array all other data types
' will be converted to string!
Dim i As Long, j As Long, k As Long, L As Long

If IsArray(VarArr) Then
   ConversionSucceeded = True
   On Error GoTo ErrorHappened
   If NumberOfArrayDimensions(VarArr) = 1 Then
   
       ReDim LngArr(LBound(VarArr) To UBound(VarArr))
       For i = LBound(LngArr) To UBound(LngArr)
           LngArr(i) = CLng(VarArr(i))
       Next
       
   ElseIf NumberOfArrayDimensions(VarArr) = 2 Then
   
       ReDim LngArr(LBound(VarArr, 1) To UBound(VarArr, 1), _
                    LBound(VarArr, 2) To UBound(VarArr, 2))
       For i = LBound(LngArr, 1) To UBound(LngArr, 1)
           For j = LBound(LngArr, 2) To UBound(LngArr, 2)
               LngArr(i, j) = CLng(VarArr(i, j))
           Next
       Next
       
   ElseIf NumberOfArrayDimensions(VarArr) = 3 Then
   
       ReDim LngArr(LBound(VarArr, 1) To UBound(VarArr, 1), _
                    LBound(VarArr, 2) To UBound(VarArr, 2), _
                    LBound(VarArr, 3) To UBound(VarArr, 3))
       For i = LBound(LngArr, 1) To UBound(LngArr, 1)
           For j = LBound(LngArr, 2) To UBound(LngArr, 2)
               For k = LBound(LngArr, 3) To UBound(LngArr, 3)
                    LngArr(i, j, k) = CLng(VarArr(i, j, k))
               Next
           Next
       Next
   
   ElseIf NumberOfArrayDimensions(VarArr) = 4 Then
   
       ReDim LngArr(LBound(VarArr, 1) To UBound(VarArr, 1), _
                    LBound(VarArr, 2) To UBound(VarArr, 2), _
                    LBound(VarArr, 3) To UBound(VarArr, 3), _
                    LBound(VarArr, 4) To UBound(VarArr, 4))
       For i = LBound(LngArr, 1) To UBound(LngArr, 1)
           For j = LBound(LngArr, 2) To UBound(LngArr, 2)
               For k = LBound(LngArr, 3) To UBound(LngArr, 3)
                    For L = LBound(LngArr, 4) To UBound(LngArr, 4)
                         LngArr(i, j, k, L) = CLng(VarArr(i, j, k, L))
                    Next
               Next
           Next
       Next
   Else
      ConversionSucceeded = False
   End If 'out of luck for 5-D and above
   
Else
   ConversionSucceeded = False
End If

EndOfSub:

Exit Sub
ErrorHappened:
   ConversionSucceeded = False
GoTo EndOfSub
End Sub

Public Sub CDblArr(VarArr As Variant, DblArr() As Double, Optional ConversionSucceeded As Boolean = True)
' This function takes a variant array and converts it to a string array all other data types
' will be converted to string!
Dim i As Long, j As Long, k As Long, L As Long

If IsArray(VarArr) Then

   If NumberOfArrayDimensions(VarArr) = 1 Then
   
       ReDim DblArr(LBound(VarArr) To UBound(VarArr))
       For i = LBound(DblArr) To UBound(DblArr)
           DblArr(i) = CDbl(VarArr(i))
       Next
       
   ElseIf NumberOfArrayDimensions(VarArr) = 2 Then
   
       ReDim DblArr(LBound(VarArr, 1) To UBound(VarArr, 1), _
                    LBound(VarArr, 2) To UBound(VarArr, 2))
       For i = LBound(DblArr, 1) To UBound(DblArr, 1)
           For j = LBound(DblArr, 2) To UBound(DblArr, 2)
               DblArr(i, j) = CDbl(VarArr(i, j))
           Next
       Next
       
   ElseIf NumberOfArrayDimensions(VarArr) = 3 Then
   
       ReDim DblArr(LBound(VarArr, 1) To UBound(VarArr, 1), _
                    LBound(VarArr, 2) To UBound(VarArr, 2), _
                    LBound(VarArr, 3) To UBound(VarArr, 3))
       For i = LBound(DblArr, 1) To UBound(DblArr, 1)
           For j = LBound(DblArr, 2) To UBound(DblArr, 2)
               For k = LBound(DblArr, 3) To UBound(DblArr, 3)
                    DblArr(i, j, k) = CDbl(VarArr(i, j, k))
               Next
           Next
       Next
   
   ElseIf NumberOfArrayDimensions(VarArr) = 4 Then
   
       ReDim DblArr(LBound(VarArr, 1) To UBound(VarArr, 1), _
                    LBound(VarArr, 2) To UBound(VarArr, 2), _
                    LBound(VarArr, 3) To UBound(VarArr, 3), _
                    LBound(VarArr, 4) To UBound(VarArr, 4))
       For i = LBound(DblArr, 1) To UBound(DblArr, 1)
           For j = LBound(DblArr, 2) To UBound(DblArr, 2)
               For k = LBound(DblArr, 3) To UBound(DblArr, 3)
                    For L = LBound(DblArr, 4) To UBound(DblArr, 4)
                         DblArr(i, j, k, L) = CDbl(VarArr(i, j, k, L))
                    Next
               Next
           Next
       Next
   
   
   End If 'out of luck for 5-D and above
   
   'I do not know why this is populating
   If err.Source = "VBAProject" Then
      err.Source = ""
   End If
Else
   ConversionSucceeded = False
End If

End Sub

Public Sub CStrArr(VarArr As Variant, StrArr() As String, Optional ConversionSucceeded As Boolean = True)
' This function takes a variant array and converts it to a string array all other data types
' will be converted to string!
Dim i As Long, j As Long, k As Long, L As Long

If IsArray(VarArr) Then

   If NumberOfArrayDimensions(VarArr) = 1 Then
   
       ReDim StrArr(LBound(VarArr) To UBound(VarArr))
       For i = LBound(StrArr) To UBound(StrArr)
           StrArr(i) = CStr(VarArr(i))
       Next
       
   ElseIf NumberOfArrayDimensions(VarArr) = 2 Then
   
       ReDim StrArr(LBound(VarArr, 1) To UBound(VarArr, 1), _
                    LBound(VarArr, 2) To UBound(VarArr, 2))
       For i = LBound(StrArr, 1) To UBound(StrArr, 1)
           For j = LBound(StrArr, 2) To UBound(StrArr, 2)
               StrArr(i, j) = CStr(VarArr(i, j))
           Next
       Next
       
   ElseIf NumberOfArrayDimensions(VarArr) = 3 Then
   
       ReDim StrArr(LBound(VarArr, 1) To UBound(VarArr, 1), _
                    LBound(VarArr, 2) To UBound(VarArr, 2), _
                    LBound(VarArr, 3) To UBound(VarArr, 3))
       For i = LBound(StrArr, 1) To UBound(StrArr, 1)
           For j = LBound(StrArr, 2) To UBound(StrArr, 2)
               For k = LBound(StrArr, 3) To UBound(StrArr, 3)
                    StrArr(i, j, k) = CStr(VarArr(i, j, k))
               Next
           Next
       Next
   
   ElseIf NumberOfArrayDimensions(VarArr) = 4 Then
   
       ReDim StrArr(LBound(VarArr, 1) To UBound(VarArr, 1), _
                    LBound(VarArr, 2) To UBound(VarArr, 2), _
                    LBound(VarArr, 3) To UBound(VarArr, 3), _
                    LBound(VarArr, 4) To UBound(VarArr, 4))
       For i = LBound(StrArr, 1) To UBound(StrArr, 1)
           For j = LBound(StrArr, 2) To UBound(StrArr, 2)
               For k = LBound(StrArr, 3) To UBound(StrArr, 3)
                    For L = LBound(StrArr, 4) To UBound(StrArr, 4)
                         StrArr(i, j, k, L) = CStr(VarArr(i, j, k, L))
                    Next
               Next
           Next
       Next
   
   
   End If 'out of luck for 5-D and above
   
Else
   ConversionSucceeded = False
End If

End Sub

Public Function Extract2DArrayFrom3DArray(ByVal arr As Variant, DimToSetConstant As Long, IndToSetConstant) As Variant

If DimToSetConstant < 3 And DimToSetConstant > 0 Then
   If NumberOfArrayDimensions(arr) <> 3 Then
        err.Source = "ArrayMod_shared.Extract2DArrayFrom3DArray: Input array ""Arr"" must be 3-D!"
        GoTo ErrorHappened
   ElseIf IndToSetConstant > UBound(arr, DimToSetConstant) Or IndToSetConstant < LBound(arr, DimToSetConstant) Then
        err.Source = "ArrayMod_shared.Extract2DArrayFrom3DArray: 3rd input ""IndToSetConstant""= " & IndToSetConstant & _
        " must be within the" & " indice bounds of the " & DimToSetConstant & " of ""Arr"" " & _
        " which are " & LBound(arr, DimToSetConstant) & " and " & UBound(arr, DimToSetConstant) & "."
        GoTo ErrorHappened
   Else
        Dim i As Long, j As Long
        Dim TempVariant As Variant
        Select Case DimToSetConstant
            Case 1
                 ReDim TempVariant(LBound(arr, 2) To UBound(arr, 2), LBound(arr, 3) To UBound(arr, 3))
                 For i = LBound(arr, 2) To UBound(arr, 2)
                    For j = LBound(arr, 3) To UBound(arr, 3)
                        TempVariant(i, j) = arr(IndToSetConstant, i, j)
                    Next
                 Next
            Case 2
                 ReDim TempVariant(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 3) To UBound(arr, 3))
                 For i = LBound(arr, 1) To UBound(arr, 1)
                    For j = LBound(arr, 3) To UBound(arr, 3)
                        TempVariant(i, j) = arr(i, IndToSetConstant, j)
                    Next
                 Next
            Case 3
                 ReDim TempVariant(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2))
                 For i = LBound(arr, 1) To UBound(arr, 1)
                    For j = LBound(arr, 2) To UBound(arr, 2)
                        TempVariant(i, j) = arr(i, j, IndToSetConstant)
                    Next
                 Next
        End Select
        
   
   End If


Else
   err.Source = "ArrayMod_shared.Extract2DArrayFrom3DArray: 2nd argument ""DimToSetConstant"" must be between 1 and 3"
   GoTo ErrorHappened
End If

Extract2DArrayFrom3DArray = TempVariant

FunctionExit:

Exit Function
ErrorHappened:
    Extract2DArrayFrom3DArray = Empty
    MsgBox err.Source, vbCritical, "Error!"
GoTo FunctionExit
End Function

Public Function UnFlattenFlatArray(FlatArray As Variant, DimensionColumns() As Long, _
                                   DataColumns() As Long, Optional FillValue As Variant, _
                                   Optional UniqueSets As Collection, Optional CaseInvariant As Boolean = False) As Collection
' DimensionColumns and DataColumns must both be 1D StringArrays with column numbers which exist for the
' array "FlatArray"

' WARNING

' This function takes a "Flat array" which is a 2D array and transforms it into a Collection of m, n dimensional arrays
' where m = NumElements(DataColumns) and n = Numelements(DimensionColumns) and n vectors i which contain the sorted unique.
' elements for each dimension in sorted order.  Any indices not touched by the flat table
' structure are assigned a value of FillValue.

' Catch all datatype errors
If IsMissing(FillValue) Then
    FillValue = -999999999#
End If

If Not IsArray(FlatArray) Or Not NumberOfArrayDimensions(FlatArray) = 2 Or IsObject(FlatArray) Then
     err.Source = "ArrayMod_shared.UnFlattenFlatArray: Incorrect Data Format! 1st Input ""FlatArray""" & _
                  " must be a 2-D Variant Array "
     GoTo ErrorHappened
ElseIf IsArrayEmpty(DimensionColumns) Or IsArrayEmpty(DataColumns) Then
     err.Source = "ArrayMod_shared.UnFlattenFlatArray: 2nd and 3rd inputs ""DimensionColumns"" and ""DataColumns""" & _
                  "cannot be empty arrays!"
     GoTo ErrorHappened
ElseIf NumElements(DimensionColumns) > 4 Then
     err.Source = "ArrayMod_shared.UnFlattenFlatArray: 2nd Input ""DimensionColumns"" cannot have more than 4 elements!"
     GoTo ErrorHappened
ElseIf IsArray(FillValue) Or IsObject(FillValue) Then
     err.Source = "ArrayMod_shared.UnFlattenFlatArray: 4th Input ""FillValue"" must be a number or string!"
     GoTo ErrorHappened
Else
     Dim ColumnExtracted As Boolean
     
     Dim Upper As Long
     Dim Lower As Long
     Dim i As Long, j As Long, k As Long
     Dim NumDimensions As Long, NumDataPoints As Long
     
     Dim Ind() As Long
     Dim NumElPerDim() As Long
     
     Dim CurrentColumn As Variant
     Dim UniqueSet As Variant
     Dim CurrentUniqueSet As Variant
     Dim CurArray As Variant
     Dim SortInd As Variant
     
     Set UnFlattenFlatArray = New Collection
     Set UniqueSets = New Collection
     
     Upper = UBound(FlatArray, 2)
     Lower = LBound(FlatArray, 2)
     
     ' Proceed to Unflatten
     '   Find the unique set for each column and sort it
     
     
     NumDataPoints = UBound(FlatArray, 1) - LBound(FlatArray, 1) + 1
     NumDimensions = UBound(DimensionColumns) - LBound(DimensionColumns) + 1
     ReDim Ind(NumDataPoints - 1, NumDimensions - 1)
     ReDim NumElPerDim(0 To NumDimensions - 1)
     ReDim CurrentColumn(0)
     
     For i = LBound(DimensionColumns) To UBound(DimensionColumns)
          ' Catch whether incorrect column numbers have been input
          If Lower > DimensionColumns(i) Or Upper < DimensionColumns(i) Then
               err.Source = "ArrayMod_shared.UnFlattenFlatArray: 2nd and 3rd inputs ""DimensionsColums"" and ""DataColumns""" & _
                            "must have elements within the bounds of the 2nd dimensions of 1st input ""FlatArray""!"
               GoTo ErrorHappened
          End If
          
          ' Extract the column indicated
          ColumnExtracted = ArrayMod_shared.GetColumn(FlatArray, CurrentColumn, DimensionColumns(i))
          If Not ColumnExtracted Then
              err.Source = "ArrayMod_shared.UnFlattenFlatArray: ArrayMod_shared.GetColumn Failed on ""FlatArray"" for column " & _
                           DimensionColumns(i) & "."
              GoTo ErrorHappened
          End If
          ' find the unique set
          CurrentUniqueSet = ReturnUniqueItemsForVariantStringArray(CurrentColumn)
          ' had to add case invariance, VBA and SQL in Access Sorting differently is a major pain! 6/1/2015
          SortInd = Sort(CurrentUniqueSet, CaseInvariant)
          NumElPerDim(i - LBound(DimensionColumns)) = NumElements(CurrentUniqueSet)
          UniqueSets.Add CurrentUniqueSet
     Next
     
     ' Find all of the indices positions for the unique sets (remember everything has been sorted)
     For i = 0 To NumDimensions - 1
          UniqueSet = UniqueSets(i + 1)
          For j = LBound(FlatArray, 1) To UBound(FlatArray, 1)
              Ind(j, i) = BinaryStringSearch(CStr(FlatArray(j, DimensionColumns(i))), UniqueSet, False, CaseInvariant)
              If Ind(j, i) = -1 Then
                  MsgBox "This should never happen!"
              End If
          Next
     Next
     
     'Now fill the current array
     For i = LBound(DataColumns) To UBound(DataColumns)
            ' Catch whether incorrect column numbers have been input
            If Lower > DataColumns(i) Or Upper < DataColumns(i) Then
                 err.Source = "ArrayMod_shared.UnFlattenFlatArray: 2nd and 3rd inputs ""DimensionsColums"" and ""DataColumns""" & _
                              "must have elements within the bounds of the 2nd dimensions of 1st input ""FlatArray""!"
                 GoTo ErrorHappened
            End If
            
            'Fill the Current Array
            Select Case NumDimensions
                 Case 1
                     CurArray = FillVariantArray(FillValue, NumElPerDim(0))
                 Case 2
                     CurArray = FillVariantArray(FillValue, NumElPerDim(0), NumElPerDim(1))
                 Case 3
                     CurArray = FillVariantArray(FillValue, NumElPerDim(0), NumElPerDim(1), NumElPerDim(2))
                 Case 4
                     CurArray = FillVariantArray(FillValue, NumElPerDim(0), NumElPerDim(1), NumElPerDim(2), _
                                NumElPerDim(3))
             End Select
            
            
            For j = LBound(FlatArray, 1) To UBound(FlatArray, 1)
                Select Case NumDimensions
                   Case 1
                       CurArray(Ind(j, 0)) = FlatArray(j, DataColumns(i))
                   Case 2
                       CurArray(Ind(j, 0), Ind(j, 1)) = FlatArray(j, DataColumns(i))
                   Case 3
                       CurArray(Ind(j, 0), Ind(j, 1), Ind(j, 2)) = FlatArray(j, DataColumns(i))
                   Case 4
                       CurArray(Ind(j, 0), Ind(j, 1), Ind(j, 2), Ind(j, 3)) = FlatArray(j, DataColumns(i))
                End Select
            Next
            UnFlattenFlatArray.Add CurArray
     Next
End If


FunctionExit:


Exit Function
ErrorHappened:
    
    UnFlattenFlatArray.Add False
    MsgBox err.Source, vbCritical, "ArrayMod_shared.UnFlattenFlatArray Error!"
    
GoTo FunctionExit
End Function

Public Function UniqueColumnsOf2DArray(Array2D As Variant) As Variant

Dim i As Long, j As Long
Dim Col1 As Variant
Dim Col2 As Variant
Dim TempUniqueCol() As Long
Dim NextUnAssignedCol As Long
Dim FirstUnAssignedCol As Long
Dim NumColumn As Long
Dim UB2 As Long, LB2 As Long, UB1 As Long, LB1 As Long
Dim GetColumnWorked As Boolean
Dim Ind As Long
Dim ColumnsEquivalent As Boolean
Dim NumIter As Long
Dim Ind2 As Long
Dim count As Long

Col1 = Array(0)
Col2 = Array(0) ' Make these dynamic arrays.

If NumberOfArrayDimensions(Array2D) <> 2 Then
   err.Source = "ArrayMod_shared.UniqueColumnsOf2DArray: Input must be a two dimensional array"
ElseIf UBound(Array2D, 2) - LBound(Array2D, 2) = 0 Then 'Only one column exists and it is unique!
   ReDim TempUniqueCol(0)
   TempUniqueCol(0) = 0
Else
   UB2 = UBound(Array2D, 2)
   LB2 = LBound(Array2D, 2)
   UB1 = UBound(Array2D, 1)
   LB1 = LBound(Array2D, 1)
   NumColumn = UB2 - UB1 + 1
   ReDim TempUniqueCol(LB2 To UB2)
   For i = LB2 To UB2
       TempUniqueCol(i) = -1 ' use this in searches
   Next
   ' Check for non-simple data types
    For i = LB1 To UB1
        For j = LB2 To UB2
            If Not IsSimpleDataType(Array2D(i, j)) Then
                err.Source = "ArrayMod_shared.UniqueColumnsOf2DArray: Input elements must all be simple variable types"
                GoTo ErrorHappened
            End If
        Next
    Next
    ' Perform
    ' this loop moves forward.  Starting with the first unassigned column it checks every unassigned column after
    ' it until the last column is reached.  Positive matches are assigned the next unique column number
     ' this is always the starting case
    FirstUnAssignedCol = LB2
    count = 0
    Do
       If FirstUnAssignedCol = 14 Then
          Ind = Ind
       End If
       'get first unassigned column
       GetColumnWorked = ArrayMod_shared.GetColumn(Array2D, Col1, FirstUnAssignedCol)
       
       TempUniqueCol(FirstUnAssignedCol) = count
       If Not GetColumnWorked Then
          err.Source = "ArrayMod_shared.UniqueColumnsOf2DArray: GetColumn function failed for an unknown reason"
          GoTo ErrorHappened
       End If
       
       Ind = FirstUnAssignedCol
       Do
           
            Ind = Ind + 1
            If TempUniqueCol(Ind) = -1 Then 'perform Test
                GetColumnWorked = ArrayMod_shared.GetColumn(Array2D, Col2, Ind)
                If Not GetColumnWorked Then
                   err.Source = "ArrayMod_shared.UniqueColumnsOf2DArray: GetColumn function failed for an unknown reason"
                   GoTo ErrorHappened
                End If
                If Ind = 14 Then
                   Ind = Ind
                End If
                
                
                ColumnsEquivalent = VectorsAreEquivalent(Col1, Col2)
                
                If ColumnsEquivalent Then
                   TempUniqueCol(Ind) = count
                End If
            End If
       Loop Until Ind = UB2
       
       ' Now find the next unassigned column
       Ind2 = FirstUnAssignedCol
       Do
         Ind2 = Ind2 + 1
       Loop Until TempUniqueCol(Ind2) = -1 Or Ind2 = UB2
       
       'find next unassigned column
       FirstUnAssignedCol = Ind2
       
       'Capture exception if every column assigned except the last (no comparisons to make)
       If FirstUnAssignedCol = UB2 Then
          If TempUniqueCol(UB2) = -1 Then
              TempUniqueCol(UB2) = count + 1
          End If
          FirstUnAssignedCol = UB2 + 1 'cause exit to occur
       End If
       
       count = count + 1
         
    Loop Until FirstUnAssignedCol > UB2

End If

UniqueColumnsOf2DArray = TempUniqueCol

FunctionExit:

Exit Function

ErrorHappened:
   UniqueColumnsOf2DArray = Empty
GoTo FunctionExit
End Function

Public Function VectorsAreEquivalent(Vec1 As Variant, Vec2 As Variant) As Boolean
' This only tests for equivalence of single data types (no jagged arrays!) of variant vectors
Dim i As Long
Dim iAdj As Long

VectorsAreEquivalent = True

If NumElements(Vec1) <> NumElements(Vec2) Then
   VectorsAreEquivalent = False
ElseIf NumElements(Vec2) = 0 And NumElements(Vec1) = 0 Then
    If VarType(Vec2) = VarType(Vec1) Then
        If Vec2 <> Vec1 Then
           VectorsAreEquivalent = False
        End If
    Else
        VectorsAreEquivalent = False
    End If
Else
    For i = LBound(Vec1) To UBound(Vec2)
        iAdj = i - LBound(Vec1) + LBound(Vec2)
        If Not IsSimpleDataType(Vec1(i)) Or Not IsSimpleDataType(Vec2(iAdj)) Then                           'There may be holes in this!
              err.Source = "ArrayMod_shared.VectorsAreEquivalent: Only two vectors with single entry data types in every element are accepted!" & _
              " Function Failed and has returned False when two jagged arrays may be equivalent!"
              GoTo ErrorHappened
        End If
        If VarType(Vec1(i)) = VarType(Vec2(iAdj)) Then
            If Vec1(i) <> Vec2(iAdj) Then
               VectorsAreEquivalent = False
            End If
        Else
            VectorsAreEquivalent = False
            Exit For
        End If
    Next

End If

FunctionExit:


Exit Function
ErrorHappened:
    VectorsAreEquivalent = False
    MsgBox err.Source, vbCritical, "Error!"
GoTo FunctionExit
End Function

Function ContainsNonNumeric(Vec As Variant) As Boolean
Dim i As Long
Dim j As Long
Dim k As Long
Dim L As Long
Dim Dum As Variant
   ContainsNonNumeric = False
   If IsArray(Vec) Then
      
      Select Case NumberOfArrayDimensions(Vec)
          Case 0
             If Not IsNumericDataType(Vec) Then GoTo FoundNonNumericEntry
          Case 1
             For i = LBound(Vec) To UBound(Vec)
                If Not IsNumericDataType(Vec(i)) Then GoTo FoundNonNumericEntry
             Next i
          Case 2
             For i = LBound(Vec, 1) To UBound(Vec, 1)
                For j = LBound(Vec, 2) To UBound(Vec, 2)
                   If Not IsNumericDataType(Vec(i, j)) Then GoTo FoundNonNumericEntry
                Next j
             Next i
          Case 3
             For i = LBound(Vec, 1) To UBound(Vec, 1)
                For j = LBound(Vec, 2) To UBound(Vec, 2)
                   For k = LBound(Vec, 3) To UBound(Vec, 3)
                      If Not IsNumericDataType(Vec(i, j, k)) Then GoTo FoundNonNumericEntry
                   Next k
                Next j
             Next i
          Case 4
              For i = LBound(Vec, 1) To UBound(Vec, 1)
                For j = LBound(Vec, 2) To UBound(Vec, 2)
                   For k = LBound(Vec, 3) To UBound(Vec, 3)
                      For L = LBound(Vec, 4) To UBound(Vec, 4)
                          If Not IsNumericDataType(Vec(i, j, k, L)) Then GoTo FoundNonNumericEntry
                      Next L
                   Next k
                Next j
              Next i
          Case Else
                err.Source = "ArrayMod_shared.ArrayContainsNaN: This function only handles arrays of dimension 1 to 4. The input was a " & _
                             CStr(NumberOfArrayDimensions(Vec)) & " dimensional array!"
                Dum = mdlError.ReturnError(IncludeMsgBox:=True)
       End Select
   
    Else
         err.Source = "ArrayMod_shared.ContainsNonNumeric: An Incorrect type has been input into this function.  Only arrays are allowed"
         Dum = mdlError.ReturnError(IncludeMsgBox:=True)
    End If

Exit Function
FoundNonNumericEntry:
    ContainsNonNumeric = True
End Function

Sub ConvertToZerosAndOnes(A() As Double, Out() As Double, Optional AboveZeroValue As Double = 1, Optional BelowZeroValue As Double = 1)

' This function converts all nonzero entries in A to AboveZeroValue value if A(i,j)>0 and BelowZeroValue when A(i,j)<0 and leaves all zero entries unchanged

Dim i As Long
Dim j As Long
Dim k As Long
Dim L As Long
Dim Dum As Double

If ArrayMod_shared.IsArrayAllocated(A) Then
      Select Case NumberOfArrayDimensions(A)
          Case 1
             ReDim Out(LBound(A) To UBound(A))
             For i = LBound(A) To UBound(A)
                If A(i) > 0 Then
                   Out(i) = AboveZeroValue
                ElseIf A(i) < 0 Then
                   Out(i) = BelowZeroValue
                Else
                   Out(i) = 0
                End If
             Next i
          Case 2
             ReDim Out(LBound(A, 1) To UBound(A, 1), LBound(A, 2) To UBound(A, 2))
             For i = LBound(A, 1) To UBound(A, 1)
                For j = LBound(A, 2) To UBound(A, 2)
                    If A(i, j) > 0 Then
                       Out(i, j) = AboveZeroValue
                    ElseIf A(i, j) < 0 Then
                       Out(i, j) = BelowZeroValue
                    Else
                       Out(i, j) = 0
                    End If
               Next j
            Next i
          Case 3
             ReDim Out(LBound(A, 1) To UBound(A, 1), LBound(A, 2) To UBound(A, 2), LBound(A, 3) To UBound(A, 3))
             For i = LBound(A, 1) To UBound(A, 1)
                For j = LBound(A, 2) To UBound(A, 2)
                   For k = LBound(A, 3) To UBound(A, 3)
                        If A(i, j, k) > 0 Then
                           Out(i, j, k) = AboveZeroValue
                        ElseIf A(i, j, k) < 0 Then
                           Out(i, j, k) = BelowZeroValue
                        Else
                           Out(i, j, k) = 0
                        End If
                   Next k
                Next j
             Next i
          Case 4
             ReDim Out(LBound(A, 1) To UBound(A, 1), LBound(A, 2) To UBound(A, 2), LBound(A, 3) To UBound(A, 3), _
                       LBound(A, 4) To UBound(A, 4))
              For i = LBound(A, 1) To UBound(A, 1)
                For j = LBound(A, 2) To UBound(A, 2)
                   For k = LBound(A, 3) To UBound(A, 3)
                      For L = LBound(A, 4) To UBound(A, 4)
                            If A(i, j, k, L) > 0 Then
                               Out(i, j, k, L) = AboveZeroValue
                            ElseIf A(i, j, k, L) < 0 Then
                               Out(i, j, k, L) = BelowZeroValue
                            Else
                               Out(i, j, k, L) = 0
                            End If
                      Next L
                   Next k
                Next j
              Next i
          Case Else
                err.Source = "ArrayMod_shared.OneAndZeroArray: This function only handles arrays of dimension 1 to 4. The input was a " & _
                             CStr(NumberOfArrayDimensions(A)) & " dimensional array!"
                Dum = mdlError.ReturnError(IncludeMsgBox:=True)
       End Select
Else
    err.Source = "ArrayMod_shared.OneAndZeroArray: This function only handles allocated numeric arrays!"
    Dum = mdlError.ReturnError(IncludeMsgBox:=True)
End If

End Sub

Sub IsMemberLng(A() As Long, B() As Long, Ans() As Long, Optional SortB As Boolean = True)

    Dim Ind As Long
    Dim Dum As Variant
    Dim i As Long
    
    If NumberOfArrayDimensions(A) <> 1 Or NumberOfArrayDimensions(B) <> 1 Then
       err.Source = "ArrayMod_shared.IsMemberLng: Both inputs must be 1-D arrays of vectors!"
       Dum = mdlError.ReturnError(IncludeMsgBox:=True)
    Else
       ReDim Ans(LBound(A) To UBound(A))
          
       For i = LBound(A) To UBound(A)
          If FindIndexLng(A(i), B) <> -1 Then
             Ans(i) = 1
          Else
             Ans(i) = 0
          End If
       Next
    End If
End Sub

Function FindIndexLng(ByRef IntegerToFind As Long, _
                   ByRef IntegerArray() As Long, _
                   Optional ByVal IsSorted As Boolean = False, _
                   Optional ByVal AcceptClosestMatch As Boolean = False) As Long ' Default since it will work whether it is sorted or not (just longer time to run)
'WARNING: IF THere are duplicate entries, this will find one of the duplicate entries and it gets even
' more ambiguous if AcceptClosestMatch = True
' but does not tell the user that it is a duplicate entry.
' If Accept Closest match is true then it will find the Upper memeber of the two nearest matches (assuming no duplicates)

Dim i As Long

If Not IsArray(IntegerArray) Or _
   ArrayMod_shared.NumberOfArrayDimensions(IntegerArray) <> 1 Then
     MsgBox "ArrayMod_shared.FindIndexLng: second input ""IntegerArray"" must be a vector of integers!"
     FindIndexLng = -1
Exit Function
End If

If IsSorted Then
     FindIndexLng = BinarySearchForIntegers(IntegerArray, IntegerToFind, AcceptClosestMatch)
Else
     FindIndexLng = LinearIntegerSearch(IntegerToFind, IntegerArray, AcceptClosestMatch)
End If

End Function

Private Function LinearIntegerSearch(IntegerToFind As Long, IntegerArray() As Long, Optional ByVal AcceptClosestMatch As Boolean = False) As Long
'Same rules as Integer Binary search but a linear string search instead which has to keep track of which is the closest match throughout the search

    Dim i As Long, LB As Long, UB As Long
    Dim SortInd() As Long
    Dim TempInd As Long
    ' Have to keep track of which is the closest if AcceptClosestMatch is true
    LinearIntegerSearch = -1
    
If Not IsArrayEmpty(IntegerArray) Then
    
    LB = LBound(IntegerArray)
    UB = UBound(IntegerArray)
    ReDim SortInd(LB To UB)
    
    If IsArray(IntegerArray) Then
        If AcceptClosestMatch Then
           
           'This is more complicated
           SortInd = Sort(IntegerArray)
           TempInd = BinarySearchForIntegers(IntegerArray, IntegerToFind, True)
           LinearIntegerSearch = Sort(TempInd) ' Go back to the original order
           
        Else
           For i = LBound(IntegerArray) To UBound(IntegerArray)
               If IntegerToFind = IntegerArray(i) Then
                   LinearIntegerSearch = i
                   Exit For
                End If
           Next
        
        End If
    End If

End If
End Function

Public Function SubArray1D(Var As Variant, BInd As Long, EInd As Long, Optional OutBInd As Long = 0) As Variant

Dim TempVar As Variant
Dim i As Long

If NumberOfArrayDimensions(Var) <> 1 Then
    err.Source = "ArrayMod_shared.SubArray1D: Only works for a 1-Dimensional array!"
    GoTo ErrorHappend
ElseIf BInd > EInd Then
    err.Source = "ArrayMod_shared.SubArray1D: Begin Index must be greater than end index!"
    GoTo ErrorHappend
ElseIf BInd < LBound(Var) Then
    err.Source = "ArrayMod_shared.SubArray1D:Begin Index is less than the lower boundary of the input array!"
    GoTo ErrorHappend
ElseIf EInd > UBound(Var) Then
    err.Source = "ArrayMod_shared.SubArray1D:End Index is greater than the upper boundary of the input array!"
    GoTo ErrorHappend
Else
    ReDim TempVar(0 To EInd - BInd)
    For i = BInd To EInd
        TempVar(i - BInd) = Var(i)
    Next

    SubArray1D = TempVar

End If

EndOfFunction:


Exit Function
ErrorHappend:
    MsgBox err.Source, vbCritical, "SubArray1 Error"
    SubArray1D = Empty
GoTo EndOfFunction
End Function

Public Function SubArray2D(Var As Variant, BInd1 As Long, EInd1 As Long, BInd2, EInd2) As Variant

Dim TempVar As Variant
Dim i As Long, j As Long

If NumberOfArrayDimensions(Var) <> 2 Then
    err.Source = "ArrayMod_shared.SubArray2D: Only works for a 2-Dimensional array!"
    GoTo ErrorHappend
ElseIf BInd1 > EInd1 Then
    err.Source = "ArrayMod_shared.SubArray2D: Begin Index 1 must be greater than end index 1!"
    GoTo ErrorHappend
ElseIf BInd2 > EInd2 Then
    err.Source = "ArrayMod_shared.SubArray2D: Begin Index 2 must be greater than end index 2!"
    GoTo ErrorHappend
ElseIf BInd1 < LBound(Var, 1) Then
    err.Source = "ArrayMod_shared.SubArray2D:Begin Index 1 is less than the lower boundary of the input array's 1st Dimension!"
    GoTo ErrorHappend
ElseIf EInd1 > UBound(Var, 1) Then
    err.Source = "ArrayMod_shared.SubArray2D:End Index 1 is greater than the upper boundary of the input array's 1st Dimension!"
    GoTo ErrorHappend
ElseIf BInd2 < LBound(Var, 2) Then
    err.Source = "ArrayMod_shared.SubArray2D:Begin Index 2 is less than the lower boundary of the input array's 2nd Dimension!"
    GoTo ErrorHappend
ElseIf EInd2 > UBound(Var, 2) Then
    err.Source = "ArrayMod_shared.SubArray2D:End Index 2 is greater than the upper boundary of the input array's 2nd Dimension!"
    GoTo ErrorHappend
Else
    ReDim TempVar(0 To EInd1 - BInd1, 0 To EInd2 - BInd2)
    For i = BInd1 To EInd1
        For j = BInd2 To EInd2
            TempVar(i - BInd1, j - BInd2) = Var(i, j)
        Next
    Next
    SubArray2D = TempVar
End If

EndOfFunction:


Exit Function
ErrorHappend:
    MsgBox err.Source, vbCritical, "SubArray1 Error"
    SubArray2D = Empty
GoTo EndOfFunction
End Function

Public Sub CreateRandomArrayWithZerosAndScrambledNonZeroDiagonal(FractionZeros As Double, NumberRows As Long, NumberColumns As Long, RowStartInd As Long, ColumnStartInd As Long, _
                                      HighValue As Double, LowValue As Double, OutArr() As Double)

Dim i As Long
Dim j As Long
Dim Random As Single
Dim NewRowOrder() As Single
Dim IX As Variant
Dim TempArr() As Double

ReDim NewRowOrder(1 To NumberRows)
ReDim OutArr(RowStartInd To NumberRows + RowStartInd - 1, ColumnStartInd To NumberColumns + ColumnStartInd - 1)
ReDim TempArr(RowStartInd To NumberRows + RowStartInd - 1, ColumnStartInd To NumberColumns + ColumnStartInd - 1)
' Do not allow an invalid entry -- Pass through even if the entry is invalid
If FractionZeros < 0 Then
  FractionZeros = 0
ElseIf FractionZeros > 1 Then
  FractionZeros = 1
End If

For i = RowStartInd To NumberRows + RowStartInd - 1
   For j = ColumnStartInd To NumberColumns + ColumnStartInd - 1
       Random = Rnd()
       If Random < FractionZeros And Not (i - RowStartInd = j - ColumnStartInd) Then ' Diagonals must be non-zero
           TempArr(i, j) = 0
       Else
           TempArr(i, j) = Rnd() * (HighValue - LowValue) + LowValue
       End If
       NewRowOrder(i - RowStartInd + 1) = Rnd()
   Next j
Next i

IX = Sort(NewRowOrder)

For i = RowStartInd To NumberRows + RowStartInd - 1
   For j = ColumnStartInd To NumberColumns + ColumnStartInd - 1
      OutArr(IX(i) + RowStartInd - 1, j) = TempArr(i, j)
   Next j
Next i

End Sub


