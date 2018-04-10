Attribute VB_Name = "StringModule_shared"
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

Public Function StringCanBeConvertedToLong(Str As String) As Boolean

On Error GoTo ErrorHappened

Dim L As Long

L = CLng(Str)
StringCanBeConvertedToLong = True

EndOfFunction:


Exit Function
ErrorHappened:
StringCanBeConvertedToLong = False
GoTo EndOfFunction
End Function

Public Function NextLinePosition(Str As String, ByVal CurrentPosition As Long) As Long

Dim StringLength As Long
Dim TStr As String

StringLength = Len(Str)
 
If CurrentPosition < 1 Or CurrentPosition > Len(Str) Then
    MsgBox "StringModule_shared.NextLinePosition: 2nd Argument ""CurrentPosition"" must be within" & _
           " the length of the first string argument! Lenght of string is " & StringLength & _
           " requested position is " & CurrentPosition
    NextLinePosition = -1
End If

Do
      If CurrentPosition = StringLength Then
         MsgBox "StringModule_shared.NextLinePosition: The requested Current Position is on the last line!" & _
                " no new line exists!"
      End If
      TStr = Mid(Str, CurrentPosition, 1)
      CurrentPosition = CurrentPosition + 1
      
Loop Until TStr = vbCr Or TStr = vbLf Or CurrentPosition > StringLength
 
Do
    TStr = Mid(Str, CurrentPosition, 1)
    CurrentPosition = CurrentPosition + 1
Loop Until TStr <> vbCr And TStr <> vbLf Or CurrentPosition > StringLength

NextLinePosition = CurrentPosition - 1

End Function

Public Function FindIndex(ByRef StringToFind As String, _
                   ByRef StringArray As Variant, _
                   Optional ByVal IsSorted As Boolean = False, _
                   Optional ByVal AcceptClosestMatch As Boolean = False, _
                   Optional ByVal CaseSensitive As Boolean = True) ' Default since it will work whether it is sorted or not (just longer time to run)
'WARNING: IF THere are duplicate entries, this will find one of the duplicate entries and it gets even
' more ambiguous if AcceptClosestMatch = True
' but does not tell the user that it is a duplicate entry.
' If Accept Closest match is true then it will find the Upper memeber of the two nearest matches (assuming no duplicates)

Dim i As Long

If Not IsArray(StringArray) Or _
   ArrayMod_shared.NumberOfArrayDimensions(StringArray) <> 1 Then
     MsgBox "StringModule_shared.FindIndex: second input ""StringArray"" must be a vector of type variant!"
     FindIndex = -1
Exit Function
End If

If Not CaseSensitive Then
   StringToFind = UCase(StringToFind)
   For i = LBound(StringArray) To UBound(StringArray)
       StringArray(i) = UCase(StringArray(i))
   Next
End If

If IsSorted Then
     FindIndex = BinaryStringSearch(StringToFind, StringArray, AcceptClosestMatch)
Else
     FindIndex = LinearStringSearch(StringToFind, StringArray, AcceptClosestMatch)
End If

End Function

Function FindIndices(ByRef StringToFind As String, _
                   ByRef StringArray As Variant) As Long() ' Default since it will work whether it is sorted or not (just longer time to run)
'WARNING: IF THere are duplicate entries, this will find one of the duplicate entries and it gets even
' more ambiguous if AcceptClosestMatch = True
' but does not tell the user that it is a duplicate entry.
' If Accept Closest match is true then it will find the Upper memeber of the two nearest matches (assuming no duplicates)

Dim i As Long
Dim Str As String
Dim Var As Variant

mdlGlobalConstants.InitializeGlobalConstants

Str = ""
For i = LBound(StringArray) To UBound(StringArray)
    If StringArray(i) = StringToFind Then
        Str = Str & mdlGlobalConstants.DefaultStringDelimiter & i
    End If
Next i

If Len(Str) <> 0 Then
   Str = Mid(Str, Len(mdlGlobalConstants.DefaultStringDelimiter) + 1, _
             Len(Str) - Len(mdlGlobalConstants.DefaultStringDelimiter))
   Var = Split(Str, mdlGlobalConstants.DefaultStringDelimiter)
   ArrayMod_shared.CLngArr Var, FindIndices
End If


End Function

Public Function BinaryStringSearch(StringToFind As String, StringArray As Variant, Optional ByVal AcceptClosestMatch As Boolean = True, Optional CaseInvariant As Boolean = False)
  ' A return of -1 means that no exact match was found or that an error occured
  ' during the function evaluation use AcceptClosestMatch to decide whether to accept a closest match or to return -1 if nothing matches exactly
  ' This function assumes that StringArray is a sorted array of strings starting with "A" at the top.  The algorithm is case sensitive
' Daniel Villa added case invariance on 6/1/2015
    Dim Low As Long
    Dim High As Long
    Dim Middle As Long
    Dim Iter As Long ' get rid of this once you are bug free
    Dim MaxIter As Long
    Dim SetLowEqualToMiddle As Boolean
    Dim SetHighEqualToMiddle As Boolean
    
    On Error Resume Next
    
    BinaryStringSearch = -1
    
    If IsArray(StringArray) Then
    
        MaxIter = 100
        
        Low = LBound(StringArray)
        High = UBound(StringArray)
        Middle = (High + Low) / 2
        Iter = 0
        
        Do While Abs(Low - High) > 1 And Iter < MaxIter
            If CaseInvariant Then
               SetLowEqualToMiddle = LCase(StringToFind) > LCase(StringArray(Middle))
               SetHighEqualToMiddle = LCase(StringToFind) < LCase(StringArray(Middle))
            Else
               SetLowEqualToMiddle = StringToFind > StringArray(Middle)
               SetHighEqualToMiddle = StringToFind < StringArray(Middle)
            End If
            
            
            If SetLowEqualToMiddle Then
            
               Low = Middle
            
            ElseIf SetHighEqualToMiddle Then
            
               High = Middle
            
            End If
            
            If StringToFind = StringArray(Middle) Then ' Your done!
                Exit Do
            ElseIf StringToFind = StringArray(Low) Then 'Inefficient but comprehensive
                Middle = Low
                Exit Do
            ElseIf StringToFind = StringArray(High) Then 'Inefficient but comprehensive
                Middle = High
                Exit Do
            End If
        
            Middle = (High + Low) / 2
            Iter = Iter + 1
        
        Loop
    
        If AcceptClosestMatch Then
            
            BinaryStringSearch = Middle
            If StringArray(Middle) <> StringToFind Then
                If High = UBound(StringArray) Then
                    BinaryStringSearch = High
                ElseIf Low = LBound(StringArray) Then
                    BinaryStringSearch = Low
                End If
            End If
            
        Else
            If StringToFind = StringArray(Middle) Then
                 BinaryStringSearch = Middle
            ElseIf StringToFind = StringArray(High) Then
                 BinaryStringSearch = High
            ElseIf StringToFind = StringArray(Low) Then
                 BinaryStringSearch = Low
            Else 'Its just not in the list
                 BinaryStringSearch = -1
            End If
        End If
    
    End If
End Function

Private Function LinearStringSearch(StringToFind As String, StringArray As Variant, Optional ByVal AcceptClosestMatch As Boolean = True)
'Same rules as BinaryStringSearch but a linear string search instead which has to keep track of which is the closest match throughout the search

    Dim i As Long, LB As Long, UB As Long
    Dim SortInd() As Long
    Dim TempInd As Long
    ' Have to keep track of which is the closest if AcceptClosestMatch is true
    LinearStringSearch = -1
    
    LB = LBound(StringArray)
    UB = UBound(StringArray)
    ReDim SortInd(LB To UB)
    
    If IsArray(StringArray) Then
        If AcceptClosestMatch Then
           
           'This is more complicated
           SortInd = Sort(StringArray)
           TempInd = BinaryStringSearch(StringToFind, StringArray)
           LinearStringSearch = Sort(TempInd) ' Go back to the original order
           
        Else
           For i = LBound(StringArray) To UBound(StringArray)
               If StringToFind = StringArray(i) Then
                   LinearStringSearch = i
                   Exit For
                End If
           Next
        
        End If
    End If

End Function

Function Sort(arr As Variant, Optional CaseInvariant As Boolean = False)
   Dim TempInd() As Long
   Dim i As Long
   If UBound(arr) - LBound(arr) > 100 Then 'Manual entry which may need tuning
      ReDim TempInd(LBound(arr) To UBound(arr))
      For i = LBound(arr) To UBound(arr)
         TempInd(i) = i
      Next
       QuickSort arr, LBound(arr), UBound(arr), TempInd, CaseInvariant
       Sort = TempInd
   Else
       Sort = BubbleSort(arr, CaseInvariant)
       
   End If

End Function



Private Function BubbleSort(ByRef arr As Variant, Optional CaseInvariant As Boolean = False)
'Best for < 100 items to be sorted
  Dim strTemp As String
  Dim i As Long
  Dim j As Long
  Dim lngMin As Long
  Dim lngMax As Long
  Dim TempInd() As Long
  Dim UB As Long
  Dim LB As Long
  Dim IndTemp As Long
  Dim SwitchElements As Boolean
  
If Not IsArrayEmpty(arr) Then
  
  lngMin = LBound(arr)
  lngMax = UBound(arr) '
  
  
  
ReDim TempInd(lngMin To lngMax)
  For i = lngMin To lngMax
      TempInd(i) = i
  Next
  
  For i = lngMin To lngMax - 1
    For j = i + 1 To lngMax
      If CaseInvariant Then
         SwitchElements = LCase(arr(i)) > LCase(arr(j))
      Else
         SwitchElements = arr(i) > arr(j)
      End If
         
      If SwitchElements Then
        strTemp = arr(i)
        arr(i) = arr(j)
        arr(j) = strTemp
        'Hold onto the indices too added by dlvilla
        IndTemp = TempInd(i)
        TempInd(i) = TempInd(j)
        TempInd(j) = IndTemp
      End If
    Next j
    
  Next i
  
  BubbleSort = TempInd
Else
  BubbleSort = Empty
End If

End Function



' Obtained from http://social.msdn.microsoft.com/Forums/en-US/830b42cf-8c97-4aaf-b34b-d860773281f7/sorting-an-array-in-vba-without-excel-function
' I do not even know how this function works it is best for larger groups of elements
Sub QuickSort(arr, Lo As Long, Hi As Long, ByRef TempInd() As Long, Optional CaseInvariant As Boolean = False)
  Dim varPivot As Variant
  Dim varTmp As Variant
  Dim tmpLow As Long
  Dim tmpHi As Long
  Dim IndTemp As Long
  tmpLow = Lo
  tmpHi = Hi
  
  'Daniel Villa made this function case invariant on 6/1/2015
  If CaseInvariant Then
     varPivot = LCase(arr((Lo + Hi) \ 2))
  Else
     varPivot = arr((Lo + Hi) \ 2)
  End If
  
  Do While tmpLow <= tmpHi
    If CaseInvariant Then
        Do While LCase(arr(tmpLow)) < varPivot And tmpLow < Hi
          tmpLow = tmpLow + 1
        Loop
        Do While varPivot < LCase(arr(tmpHi)) And tmpHi > Lo
          tmpHi = tmpHi - 1
        Loop
    Else
        Do While arr(tmpLow) < varPivot And tmpLow < Hi
          tmpLow = tmpLow + 1
        Loop
        Do While varPivot < arr(tmpHi) And tmpHi > Lo
          tmpHi = tmpHi - 1
        Loop
    End If
    
    If tmpLow <= tmpHi Then
      varTmp = arr(tmpLow)
      arr(tmpLow) = arr(tmpHi)
      arr(tmpHi) = varTmp
      'Inserted by Daniel Villa
      IndTemp = TempInd(tmpLow)
      TempInd(tmpLow) = TempInd(tmpHi)
      TempInd(tmpHi) = IndTemp
      'End of Insert by Daniel Villa
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
    End If
  Loop
  If Lo < tmpHi Then QuickSort arr, Lo, tmpHi, TempInd
  If tmpLow < Hi Then QuickSort arr, tmpLow, Hi, TempInd
End Sub

Public Function NumberOfMatchesInString(SubString As String, ByVal InputString As String) As Long

  Dim RExp As Object
  Dim Matches As Object
  Dim PN As Long
  
  On Error Resume Next
  
  Set RExp = CreateObject("vbscript.regexp")
  RExp.IgnoreCase = True
  RExp.Global = True
  RExp.Pattern = LiteralizeAStringForRegExpSearch(SubString)
  Set Matches = RExp.Execute(InputString)
  PN = Matches.count
  
  Set Matches = Nothing
  Set RExp = Nothing
  NumberOfMatchesInString = PN
  
End Function

Public Function FindFirstMatchInStringArray(StrToMatch As String, StrArr() As String) As Long

Dim i As Long
Dim num As Long
FindFirstMatchInStringArray = -1
For i = LBound(StrArr) To UBound(StrArr)
    num = NumberOfMatchesInString(StrToMatch, StrArr(i))
    If num <> 0 Then
        FindFirstMatchInStringArray = i
        GoTo ExitFunction
    End If
Next i

ExitFunction:

End Function

Public Function RemoveAllPatternMatches(ByRef StringToAlter As String, PatternToRemove As String)

  Dim RExp As Object
  Dim Matches As Object
  
  On Error Resume Next
  
  Set RExp = CreateObject("vbscript.regexp")
  RExp.IgnoreCase = True
  RExp.Global = True
  RExp.Pattern = PatternToRemove
  
  RemoveAllPatternMatches = RExp.Replace(StringToAlter, "")
  
End Function

Public Function Replace(ByVal OriginalText As String, _
                         ByVal FindText As String, _
                         ByVal ReplaceText As String) As String

Dim strText As String

If FindText <> "" Then
   If InStr(1, ReplaceText, FindText, vbTextCompare) = 0 Then
      strText = OriginalText
      Do While InStr(strText, FindText)
         strText = Left(strText, InStr(strText, FindText) - 1) & _
                   ReplaceText & _
                   Mid(strText, InStr(strText, FindText) + Len(FindText))
      Loop
      Replace = strText
   End If
End If

End Function


Public Function IntersectVariantStringArrays(StrArray1 As Variant, StrArray2 As Variant) As Variant
' The variants in this function are assumed to be vectors of strings
' otherwise the function will ungracefully fail!!!
Dim Str1 As String
Dim Str2 As String
Dim MatchInd As Long
Dim TMatches As Variant
Dim Matches As Variant
Dim MaxSize As Long
Dim Str1UBound As Long
Dim Str2LBound As Long
Dim Str1LBound As Long
Dim Str2UBound As Long
Dim ActualSize As Long
Dim i As Long
Dim j As Long

If IsEmpty(StrArray1) Then
    IntersectVariantStringArrays = Empty
ElseIf IsEmpty(StrArray2) Then
    IntersectVariantStringArrays = Empty
Else

    Str1UBound = UBound(StrArray1)
    Str2UBound = UBound(StrArray2)
    Str1LBound = LBound(StrArray1)
    Str2LBound = LBound(StrArray2)
    
    If Str1UBound - Str1LBound > Str2UBound - Str2LBound Then
        MaxSize = Str1UBound - Str1LBound
    Else
        MaxSize = Str2UBound - Str2LBound
    End If
    
    ReDim TMatches(0 To MaxSize)
    ActualSize = -1
       For i = Str1LBound To Str1UBound
          Str1 = StrArray1(i)
          For j = Str2LBound To Str2UBound
               Str2 = StrArray2(j)
               
               If Str1 = Str2 Then
                  ActualSize = ActualSize + 1
                  TMatches(ActualSize) = Str1
               End If
          Next
       Next
    If ActualSize = -1 Then
       IntersectVariantStringArrays = Empty
    Else
       ReDim Matches(0 To ActualSize)
       For i = 0 To ActualSize
           Matches(i) = TMatches(i)
       Next
       'Sort (Matches)
       IntersectVariantStringArrays = Matches
    End If
End If

End Function

Public Function UnionVariantStringArrays(StrArray1 As Variant, StrArray2 As Variant) As Variant
' The variants in this function are assumed to be vectors of strings
' otherwise the function will ungracefully fail!!!
Dim Str1 As String
Dim Str2 As String
Dim MatchInd As Long
Dim UnitedArr As Variant
Dim MatchFound As Boolean
Dim MaxSize As Long
Dim Str1UBound As Long
Dim Str2LBound As Long
Dim Str1LBound As Long
Dim Str2UBound As Long
Dim ActualSize As Long
Dim i As Long
Dim j As Long
Dim Ind As Variant

If Not IsArray(StrArray1) Or Not IsArray(StrArray2) Or _
   NumberOfArrayDimensions(StrArray1) <> 1 And NumberOfArrayDimensions(StrArray2) <> 1 Then
    If IsArrayEmpty(StrArray1) And Not IsArrayEmpty(StrArray2) And NumberOfArrayDimensions(StrArray2) = 1 Then
        UnionVariantStringArrays = StrArray2
        GoTo EndOfFunction
    ElseIf Not IsArrayEmpty(StrArray1) And IsArrayEmpty(StrArray2) And NumberOfArrayDimensions(StrArray1) = 1 Then
        UnionVariantStringArrays = StrArray1
        GoTo EndOfFunction
    Else
    
        err.Source = "StringModule_shared.UnionVariantStringArrays: The two inputs of this function must be 1-D arrays!"
        GoTo ErrorHappened
    End If
Else
    
    Str1UBound = UBound(StrArray1)
    Str2UBound = UBound(StrArray2)
    Str1LBound = LBound(StrArray1)
    Str2LBound = LBound(StrArray2)
    
    MaxSize = (Str1UBound - Str1LBound) + (Str2UBound - Str2LBound)
    
    ReDim UnitedArr(0 To MaxSize)
    
    
    For i = Str1LBound To Str1UBound
        UnitedArr(i - Str1LBound) = StrArray1(i)
    Next
    
    ActualSize = Str1UBound - Str1LBound
    
    For i = Str2LBound To Str2UBound
       Str2 = StrArray2(i)
       MatchFound = False
       For j = Str1LBound To Str1UBound
       
            Str1 = StrArray1(j)
            If Str1 = Str2 Then
                MatchFound = True
                Exit For
            End If
            
       Next
       
       If Not MatchFound Then 'We have a new element that needs to be added
           ActualSize = ActualSize + 1
           UnitedArr(ActualSize) = Str2
       End If
    Next
    
       ReDim SizedUnitedArr(0 To ActualSize)
       For i = 0 To ActualSize
           SizedUnitedArr(i) = UnitedArr(i)
       Next
       Ind = Sort(SizedUnitedArr)
       UnionVariantStringArrays = SizedUnitedArr
    End If

EndOfFunction:

Exit Function
ErrorHappened:
   MsgBox err.Source, vbCritical, "Error"
   UnionVariantStringArrays = Empty
GoTo EndOfFunction
End Function


Public Function SubtractVariantStringArrays(ByVal BaseArray As Variant, ByVal SubtractArray As Variant) As Variant
' The variants in this function are assumed to be vectors of strings
' otherwise the function will ungracefully fail!!!
Dim Str1 As String
Dim Str2 As String
Dim MatchInd As Long
Dim TMatches As Variant
Dim Matches As Variant
Dim MaxSize As Long
Dim Str1UBound As Long
Dim Str2LBound As Long
Dim Str1LBound As Long
Dim Str2UBound As Long
Dim ActualSize As Long
Dim i As Long
Dim j As Long
Dim Ind As Long
Dim BelongsInResult As Boolean

If IsEmpty(BaseArray) Then
   SubtractVariantStringArrays = Empty
ElseIf IsEmpty(SubtractArray) Then
   SubtractVariantStringArrays = BaseArray
Else
    
    Str1UBound = UBound(BaseArray)
    Str2UBound = UBound(SubtractArray)
    Str1LBound = LBound(BaseArray)
    Str2LBound = LBound(SubtractArray)
    
    If Str1UBound - Str1LBound > Str2UBound - Str2LBound Then
        MaxSize = Str1UBound - Str1LBound
    Else
        MaxSize = Str2UBound - Str2LBound
    End If
    
    ReDim TMatches(0 To MaxSize)
    Ind = -1
    ActualSize = Str1UBound - Str1LBound
    For i = Str1LBound To Str1UBound
       BelongsInResult = True
       Str1 = BaseArray(i)
       For j = Str2LBound To Str2UBound
            Str2 = SubtractArray(j)
            If Str1 = Str2 Then
              ' The element must be subtracted
              ActualSize = ActualSize - 1
              BelongsInResult = False
            End If
       Next
       If BelongsInResult Then
          Ind = Ind + 1
          TMatches(Ind) = Str1
       End If
    Next
       
    If Ind = -1 Then
       SubtractVariantStringArrays = Empty
    Else
       ReDim Matches(0 To ActualSize)
       For i = 0 To ActualSize
           Matches(i) = TMatches(i)
       Next
       SubtractVariantStringArrays = Matches
    End If
End If

EndOfFunction:

Exit Function
ErrorHappened:
   MsgBox err.Source, vbCritical, "Error"
   SubtractVariantStringArrays = Empty
GoTo EndOfFunction
End Function

Function ReturnUniqueItemsForVariantStringArray(VariantStringArray As Variant, Optional ReturnType As Long = 0) As Variant

    'Optional Input ReturnType = 0 'String, 1 Long, 2 Double, 3 Single, 4 Boolean

    Dim LongString As String
    Dim i As Long
    Dim j As Long
    Dim UniqueList As Variant
    Dim NumUnique As Long
    Dim LB As Long
    Dim UB As Long
    Dim MatchFound As Boolean
    Dim FinalUniqueList As Variant
    
    If IsArrayEmpty(VariantStringArray) Then
       ReturnUniqueItemsForVariantStringArray = Empty
    Else
        LB = LBound(VariantStringArray)
        UB = UBound(VariantStringArray)
        
        ReDim UniqueList(LB To UB)
        
        'The first element is always unique!
        UniqueList(LB) = VariantStringArray(LBound(VariantStringArray))
        NumUnique = 1
        
        For i = LB + 1 To UB
            MatchFound = False
            For j = LB To LB + NumUnique - 1
               If UniqueList(j) = VariantStringArray(i) Then
                   MatchFound = True
                   Exit For
               End If
            Next j
            
            If Not MatchFound Then
                UniqueList(NumUnique) = VariantStringArray(i)
                NumUnique = NumUnique + 1
            End If
        Next i
        
        ReDim FinalUniqueList(0 To NumUnique - 1)
        
        For i = 0 To NumUnique - 1
            Select Case ReturnType
               Case 0
                  FinalUniqueList(i) = CStr(UniqueList(LB + i))
               Case 1
                  FinalUniqueList(i) = CLng(UniqueList(LB + i))
               Case 2
                  FinalUniqueList(i) = CDbl(UniqueList(LB + i))
               Case 3
                  FinalUniqueList(i) = CSng(UniqueList(LB + i))
               Case 4
                  FinalUniqueList(i) = CBool(UniqueList(LB + i))
               Case Else
                  FinalUniqueList(i) = CStr(UniqueList(LB + i))
           End Select
        Next
    
        ReturnUniqueItemsForVariantStringArray = FinalUniqueList
    End If
End Function

Function LiteralizeAStringForRegExpSearch(Str As Variant) As String
Dim TempStr As String
Dim SpecialChar As String
Dim i As Long
Dim j As Long
Dim BackSlashNeeded As Boolean

SpecialChar = "[\^$.|?*+)("
       TempStr = ""
    For i = 1 To Len(Str)
       BackSlashNeeded = False
       For j = 1 To Len(SpecialChar)
            If Mid(Str, i, 1) = Mid(SpecialChar, j, 1) Then
                TempStr = TempStr & "\" & Mid(Str, i, 1)
                BackSlashNeeded = True
                Exit For
            End If
       Next
       If Not BackSlashNeeded Then
          TempStr = TempStr & Mid(Str, i, 1)
       End If
    Next
    
    LiteralizeAStringForRegExpSearch = TempStr
    
End Function

Sub FindFirstCharacterPositionInString(Str As String, Char As String, CharPos() As Long)

       Dim RExp As New RegExp
       Dim Matches As MatchCollection
       Dim m As Match
       Dim Temp() As Long
       
       RExp.Global = True
       RExp.Pattern = LiteralizeAStringForRegExpSearch(Str)
       RExp.MultiLine = True
       
       Set Matches = RExp.Execute
       ReDim Temp(Matches.count - 1)
       For i = 0 To UBound(Temp)
          Set m = Matches(i)
          Temp(i) = m.FirstIndex
       Next
       
       FindFirstCharacterPositionInString = Temp
       
  
End Sub

Public Function CreateBuildingPath(BID As String, Optional StillInDesign As Boolean = False) As String
        
    Dim Str As String
    InitializeGlobalConstants
    
    If StillInDesign Then
        If Application.Name = "Microsoft Excel" Then
           Str = ThisWorkbook.path & "\" & IXDataLocation & "\" & DesignFileLocation & "\" & BuildingInputFilesLocation & "\" & BID
        Else ' Application.Name = "Microsoft Access" Then
           Str = CurrentProject.path & "\" & DesignFileLocation & "\" & BuildingInputFilesLocation & "\" & BID
        End If
    Else
        If Application.Name = "Microsoft Excel" Then
           Str = ThisWorkbook.path & "\" & IXDataLocation & "\" & BuildingInputFilesLocation & "\" & BID
        Else ' Application.Name = "Microsoft Access" Then
           Str = CurrentProject.path & "\" & BuildingInputFilesLocation & "\" & BID
        End If
    End If
    
    CreateBuildingPath = Str
End Function

Public Function RemoveAllBDLDelimiters(InString As String) As String
    InString = StringModule_shared.RemoveAllPatternMatches(InString, " ")
    InString = StringModule_shared.RemoveAllPatternMatches(InString, vbCr)
    InString = StringModule_shared.RemoveAllPatternMatches(InString, vbLf)
    RemoveAllBDLDelimiters = StringModule_shared.RemoveAllPatternMatches(InString, "=")

End Function

Public Sub WriteStringToFile(StringToWrite As String, FName As String)

 Dim FSysObj As New FileSystemObject
 Dim TxtStream As TextStream
 
 
 Set TxtStream = FSysObj.OpenTextFile(FName, ForWriting, True)
 
 TxtStream.Write (StringToWrite)
 
 TxtStream.Close
 
 Set FSysObj = Nothing
 Set TxtStream = Nothing

End Sub

Function FixQuote(FQText As String) As String
    'Obtained from http://mikeperris.com/access/escaping-quotes-Access-VBA-SQL.html
    On Error GoTo Err_FixQuote
    FixQuote = Replace(FQText, "'", "''")
    FixQuote = Replace(FixQuote, """", """""")
Exit_FixQuote:
    If FixQuote = "" Then ' I had to add this Daniel Villa 4/10/2014
       FixQuote = FQText
    End If
Exit Function
Err_FixQuote:
    MsgBox err.Description, , "Error in Function Fix_Quotes.FixQuote"
    Resume Exit_FixQuote
    Resume 0 '.FOR TROUBLESHOOTING
End Function

Public Sub AppendString(MasterString As String, _
                  StringToAppend As Variant, _
                  Delimiter As String, _
                  Optional FinalPosition As Boolean = False)
If FinalPosition Then
    MasterString = MasterString & StringToAppend
Else
    MasterString = MasterString & StringToAppend & Delimiter
End If
                  
                  
End Sub

Public Sub EliminateReadyOnly(ParameterName As String)

    If StringModule_shared.NumberOfMatchesInString(" (read only)", ParameterName) > 0 Then
        ParameterName = Mid(ParameterName, 1, Len(ParameterName) - Len(" (read only)"))
    End If

End Sub

Function IsScenarioNameValid(ByVal Str As String) As Boolean
    Dim ForbiddenCharacters As String
    Dim i As Long
    ForbiddenCharacters = " ()!@#$%^&*`~{}[]\/?<>'"":;-+=,"
    Dim NumMatch As Long
    
    IsScenarioNameValid = True
    
    For i = 0 To Len(ForbiddenCharacters) - 1
        NumMatch = StringModule_shared.NumberOfMatchesInString(Mid(ForbiddenCharacters, i + 1, 1), Str)
        If NumMatch <> 0 Then
           IsScenarioNameValid = False
           Exit For
        End If
    Next
    
End Function
