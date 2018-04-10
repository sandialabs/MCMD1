VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConnectivity"
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

' This algorithm handles connectivity for 1. (co-flow and counter flow) and 2. (cross-flow and criss-cross flow)
' it always starts in the lower left hand corner of a rectangle and counts hot-side nodes first followed by cold side nodes working in a counter-clockwise loop
' in each control volume pair.
'                           Edge 3
'                ----------------------------
'                -                          -
'                -                          -
'                -                          -
'                -                          -
'      Edge 4    -                          -  Edge 2
'                -                          -
'                -                          -
'                -                          -
'                -                          -
'                -                          -
'                ----------------------------
'                           Edge 1

' The connectivity matrix is a mapping where each row represents a control volume pair. There are four nodes (2 cold side, and 2 hot side) that are represented by the
' columns of the matrix:   ' 1 = Edge 1 or 2 with hot nodes first ,
'                            2 = Edge 1 or 2 with hot nodes first
'                            3 = Edge 3 or 4 with hot nodes first
'                            4 = Edge 3 or 4 with hot nodes first
'                            5 = Direction of hot flow  0 = up, 1 = down, 2 = left, 3 = right
'                            6 = Direction of cold flow 0 = up, 1 = down, 2 = left, 3 = right
' Hot_Reversals - 0 = straight flow through the rectangle
'                       1 = single bend that exits from the same edge
'                       2 = two bends that exit from the opposite edge
'                       ...
' Cold_Reversals - same as Hot_Reversals for cold side
' Hot_Inflow_Edge - 1 - 4Col
' Cold_Inflow_Edge - 1 - 4
' Hot_Inflow_Edge_Side - 0 - 1  0 = when oriented to edge start from left hand side, 1 = right hand side
' Cold_Inflow_Edge_Side - 0 - 1

Option Explicit

Private Hot_Reversals As Long
Private Cold_Reversals As Long
Private Hot_Inflow_Edge As Long
Private Cold_Inflow_Edge As Long
Private Hot_Inflow_Edge_Side As Long
Private Cold_Inflow_Edge_Side As Long
Private GravDirection As Long ' Direction of gravity
                              ' 1 - (up) normal and inward facing to edge 1
                              ' 2 - (left) normal and inward facing to edge 2
                              ' 3 - (down) normal and inward facing to edge 3
                              ' 4 - (right) normal and inward facing to edge 4
                              ' etc...

                              
Private Horizontal_Divisions As Long 'This is the number of columns in the rectangle
Private Vertical_Divisions As Long ' This is the number of rows in the rectangle

Private ColdFlowPaths() As Long 'These are the 1-D flow paths (node numbers starting with an inflow node) each column represents an equivalent flow path
Private HotFlowPaths() As Long  'These are the 1-D flow paths (node numbers starting with an inflow node)

Private ColdCVFlowPaths() As Long 'Same as Nodal ColdFlowPaths but storing the CV numbers instead - each column represents an equivalent flow path.
Private HotCVFlowPaths() As Long 'Ditto for hot
Private GravityCVPaths() As Long

Private cM() As Long 'Connectivity Matrix
Private BC() As Long 'Boundary Condition Matrix with 2 columns
                     ' Column 1: Node numbers of boundary nodes
                     ' Column 2: 0 = entrance node with a set temperature
                     '           NodeNum = node number to which the column 1 node number travels to in a flow reversal
                     ' Column 3: 0 = hot side boundary condition, 1 = cold side boundary condition.
Private NumBC As Long
Private NumColdBC As Long 'Number of cold side boundary conditions
Private NumHotBC As Long ' Number of hot side boundary conditions
Private NumHotInflowBC As Long ' Number of hot inflow boundary conditions
Private NumColdInflowBC As Long

Public Property Get GravityDirection() As Long
   GravityDirection = GravDirection
End Property

Public Sub GetFlowPaths(HotPaths() As Long, ColdPaths() As Long, CVHotPaths() As Long, CVColdPaths() As Long)
'   ReDim HotPaths(1 To UBound(HotFlowPaths, 1), 1 To LBound(HotFlowPaths, 2))
'   ReDim ColdPaths(1 To UBound(ColdFlowPaths, 1), 1 To LBound(ColdFlowPaths, 2))
' THIS IS SLOWER BUT MORE CONVENIENT FOR CODING.
   CLngArr ColdFlowPaths, ColdPaths
   CLngArr HotFlowPaths, HotPaths
   CLngArr HotCVFlowPaths, CVHotPaths
   CLngArr ColdCVFlowPaths, CVColdPaths
End Sub

Public Property Get HotReversals() As Long
   HotReversals = Hot_Reversals
End Property

Public Property Get ColdReversals() As Long
   ColdReversals = Cold_Reversals
End Property

Public Sub GetBC(BCMatrix() As Long)

    ReDim BCMatrix(1 To NumBC, 1 To UBound(BC, 2))
    
    Dim i As Long
    Dim j As Long
    
    For i = 1 To NumBC
       For j = 1 To UBound(BC, 2)
          BCMatrix(i, j) = BC(i, j)
       Next
    Next

End Sub

Property Get InflowEdge(IsHot As Boolean, CVNum As Long) As Long

    Dim Ind As Long
    
    If IsHot Then
       Ind = 5
    Else
       Ind = 6
    End If

   Select Case cM(CVNum, Ind)
       Case 0
           InflowEdge = 1
       Case 1
           InflowEdge = 3
       Case 2
           InflowEdge = 2
       Case 3
           InflowEdge = 4
    End Select

End Property

Property Get HotFlowsHorizontal() As Boolean
    If Hot_Inflow_Edge = 1 Or Hot_Inflow_Edge = 3 Then
       HotFlowsHorizontal = False
    Else
       HotFlowsHorizontal = True
    End If
End Property

Property Get ColdFlowsHorizontal() As Boolean
    If Cold_Inflow_Edge = 1 Or Cold_Inflow_Edge = 3 Then
       ColdFlowsHorizontal = False
    Else
       ColdFlowsHorizontal = True
    End If
End Property

Property Get HotChannelWidth() As Long
    If Hot_Inflow_Edge = 1 Or Hot_Inflow_Edge = 3 Then
        HotChannelWidth = Horizontal_Divisions / (Hot_Reversals + 1)
    Else
        HotChannelWidth = Vertical_Divisions / (Hot_Reversals + 1)
    End If
End Property

Property Get ColdChannelWidth() As Long
    If Cold_Inflow_Edge = 1 Or Cold_Inflow_Edge = 3 Then
        ColdChannelWidth = Horizontal_Divisions / (Cold_Reversals + 1)
    Else
        ColdChannelWidth = Vertical_Divisions / (Cold_Reversals + 1)
    End If
End Property

Property Get HorizontalDivisions() As Long
    HorizontalDivisions = Horizontal_Divisions
End Property

Property Get VerticalDivisions() As Long
    VerticalDivisions = Vertical_Divisions
End Property

Property Get c(i As Long, j As Long) As Long ' This is the connectivity
    c = cM(i, j)
End Property
  
Property Get B(i As Long, j As Long) As Long ' This is the Boundary conditions
    B = BC(i, j)
End Property

Public Function NodeOrder(i As Long, CVNum As Long) As Long
  Dim NOrder() As Long
  Dim Dum As Double
  Dim Local_CV_Hot_Inflow_Edge As Long
  Dim Local_CV_Cold_Inflow_Edge As Long
  ' This function provides the relationship between inflow and outflow of each control volume.
  ' i = 1 - obtain hot inflow node
  ' i = 2 - obtain hot outflow node
  ' i = 3 - obtain cold inflow node
  ' i = 4 - obtain cold outflow node
'5 = Direction of hot flow  0 = up, 1 = down, 2 = left, 3 = right
'6 = Direction of cold flow 0 = up, 1 = down, 2 = left, 3 = right
  
 Local_CV_Hot_Inflow_Edge = InflowEdge(IsHot:=True, CVNum:=CVNum)
 Local_CV_Cold_Inflow_Edge = InflowEdge(IsHot:=False, CVNum:=CVNum)
  
  ' throw errors if incorrect input is requested.
  If i <= 0 Or i > 4 Then
     Dum = mdlError.ReturnError("clsConnectivity.NodeOrder: Only Edges 1 to 4 exist but edge " & CStr(i) & " has been requested!", , True)
     GoTo ErrorCondition
  ElseIf CVNum <= 0 Or CVNum > UBound(cM, 1) Then
     Dum = mdlError.ReturnError("clsConnectivity.NodeOrder: Only " & UBound(cM, 1) & " control volumes exist but " & CStr(CVNum) & " has been requested!", , True)
     GoTo ErrorCondition
  End If
  
  ReDim NOrder(1 To 4)
  
  ' This property returns
  Select Case Local_CV_Hot_Inflow_Edge
      Case 1
         'If so then hot is always first because it has a node on edge one and hot is always counted first.
         NOrder(1) = 1
         NOrder(2) = 3
         Select Case Local_CV_Cold_Inflow_Edge
            Case 1
                NOrder(3) = 2
                NOrder(4) = 4
            Case 2
                NOrder(3) = 2
                NOrder(4) = 4
            Case 3
                NOrder(3) = 4
                NOrder(4) = 2
            Case 4
                NOrder(3) = 4
                NOrder(4) = 2
         End Select
      Case 2
         Select Case Local_CV_Cold_Inflow_Edge
            Case 1 ' Then the cold inflow node is first!
                NOrder(1) = 2
                NOrder(2) = 4
                NOrder(3) = 1
                NOrder(4) = 3
            Case 2 ' Hot is first
                NOrder(1) = 1
                NOrder(2) = 3
                NOrder(3) = 2
                NOrder(4) = 4
            Case 3 ' Cold is first
                NOrder(1) = 2
                NOrder(2) = 4
                NOrder(3) = 3
                NOrder(4) = 1
            Case 4 ' Hot is first
                NOrder(1) = 1
                NOrder(2) = 3
                NOrder(3) = 4
                NOrder(4) = 2
         End Select
      Case 3
         'If so then hot is always first but the order is reversed w/r to the Case 1
         NOrder(1) = 3
         NOrder(2) = 1
         Select Case Local_CV_Cold_Inflow_Edge
            Case 1
                NOrder(3) = 2
                NOrder(4) = 4
            Case 2
                NOrder(3) = 2
                NOrder(4) = 4
            Case 3
                NOrder(3) = 4
                NOrder(4) = 2
            Case 4
                NOrder(3) = 4
                NOrder(4) = 2
         End Select
      Case 4
         Select Case Local_CV_Cold_Inflow_Edge
            Case 1 ' Then the cold inflow node is first!
                NOrder(1) = 4
                NOrder(2) = 2
                NOrder(3) = 1
                NOrder(4) = 3
            Case 2 ' Hot is first
                NOrder(1) = 3
                NOrder(2) = 1
                NOrder(3) = 2
                NOrder(4) = 4
            Case 3 ' Cold is first
                NOrder(1) = 4
                NOrder(2) = 2
                NOrder(3) = 3
                NOrder(4) = 1
            Case 4 ' Hot is first
                NOrder(1) = 3
                NOrder(2) = 1
                NOrder(3) = 4
                NOrder(4) = 2
         End Select
      Case Else
          Dum = mdlError.ReturnError("clsConnectivity.NodeOrder: Only edges 1 to 4 exist! Edge " & CStr(i) & " has been requested!")
  End Select
  
  NodeOrder = NOrder(i)

EndOfFunction:
  
Exit Function
ErrorCondition:
    NodeOrder = -1
    GoTo EndOfFunction
End Function
                     
Public Property Get NumberOfNodes() As Long
    
    Dim i As Long
    Dim NumCVP As Long
    
    NumCVP = UBound(cM, 1)
    NumberOfNodes = Application.WorksheetFunction.Max(cM(NumCVP, 1), cM(NumCVP, 2), cM(NumCVP, 3), cM(NumCVP, 4))
    
End Property

Public Property Get NumberOfControlVolumePairs() As Long
    NumberOfControlVolumePairs = UBound(cM, 1)
End Property

Public Property Get NumberOfBoundaryConditions() As Long
    Dim i As Long
    
    ' There is a boundary condition for 1. Mass Flow, 2. Temperature, 3. Salinity
    ' Therefore we have to multiply by three.  !@#$ If other chemical species concentrations
    ' need to be tracked this multiplier will need to be increased.
    
    NumberOfBoundaryConditions = 3 * NumHotBC + 3 * NumColdBC
End Property

Public Property Get SizeBC() As Long
    SizeBC = UBound(BC) - LBound(BC) + 1
End Property

Private Sub Class_Initialize()

  Dim TempErrMsg As String
  Dim LCM As Long
  
  Hot_Reversals = AssignValueIfInLimits(Range("Hot_Reversals"), TempErrMsg, "Hot Number of Reversals", mdlConstants.glbINVALID_VALUE, "", 0, 10)
  Cold_Reversals = AssignValueIfInLimits(Range("Cold_Reversals"), TempErrMsg, "Cold Number of Reversals", mdlConstants.glbINVALID_VALUE, "", 0, 10)
  Hot_Inflow_Edge = AssignValueIfInLimits(Range("Hot_Inflow_Edge"), TempErrMsg, "Hot Inflow Edge", mdlConstants.glbINVALID_VALUE, "", 1, 4)
  Cold_Inflow_Edge = AssignValueIfInLimits(Range("Cold_Inflow_Edge"), TempErrMsg, "Cold Inflow Edge", mdlConstants.glbINVALID_VALUE, "", 1, 4)
  Hot_Inflow_Edge_Side = AssignValueIfInLimits(Range("Hot_Inflow_Edge_Side"), TempErrMsg, "Hot Inflow Edge Side", mdlConstants.glbINVALID_VALUE, "", 0, 1)
  Cold_Inflow_Edge_Side = AssignValueIfInLimits(Range("Cold_Inflow_Edge_Side"), TempErrMsg, "Cold Inflow Edge Side", mdlConstants.glbINVALID_VALUE, "", 0, 1)
  
  Horizontal_Divisions = AssignValueIfInLimits(Range("Horizontal_Divisions"), TempErrMsg, "Number of horizontal divisions", mdlConstants.glbINVALID_VALUE, "", mdlConstants.glbNumberDivLowerLimit, mdlConstants.glbNumberDivUpperLimit)
  Vertical_Divisions = AssignValueIfInLimits(Range("Vertical_Divisions"), TempErrMsg, "Number of vertical divisions", mdlConstants.glbINVALID_VALUE, "", mdlConstants.glbNumberDivLowerLimit, mdlConstants.glbNumberDivUpperLimit)
  
  GravDirection = AssignValueIfInLimits(Range("RangeGravityDirection"), TempErrMsg, "Direction of Gravity", glbINVALID_VALUE, "", 1, 4)
  

  
  ' Now make sure that the number of divisions are compatible with the number of reversals
  err.Source = TempErrMsg
  TempErrMsg = ""
  
  If mdlError.NoError Then
  
      If (Hot_Inflow_Edge = 1 Or Hot_Inflow_Edge = 3) And (Cold_Inflow_Edge = 1 Or Cold_Inflow_Edge = 3) Then
         'We need the least common multiple of the number of reversals
         LCM = mdlMath.LeastCommonMultiple(Hot_Reversals + 1, Cold_Reversals + 1)
         TempErrMsg = ReversalError(Horizontal_Divisions, LCM, "horizontal", TempErrMsg)
    
      ElseIf (Hot_Inflow_Edge = 1 Or Hot_Inflow_Edge = 3) And (Cold_Inflow_Edge = 2 Or Cold_Inflow_Edge = 4) Then
      
         TempErrMsg = ReversalError(Horizontal_Divisions, Hot_Reversals + 1, "horizontal", TempErrMsg)
         TempErrMsg = ReversalError(Vertical_Divisions, Cold_Reversals + 1, "vertical", TempErrMsg)
         
      ElseIf (Hot_Inflow_Edge = 2 Or Hot_Inflow_Edge = 4) And (Cold_Inflow_Edge = 1 Or Cold_Inflow_Edge = 3) Then
      
         TempErrMsg = ReversalError(Vertical_Divisions, Hot_Reversals + 1, "vertical", TempErrMsg)
         TempErrMsg = ReversalError(Horizontal_Divisions, Cold_Reversals + 1, "horizontal", TempErrMsg)
    
      Else
      
         LCM = mdlMath.LeastCommonMultiple(Hot_Reversals + 1, Cold_Reversals + 1)
         TempErrMsg = ReversalError(Vertical_Divisions, LCM, "vertical", TempErrMsg)
      
      End If
  
  End If
  
  If Len(err.Source) = 0 And Len(TempErrMsg) <> 0 Then
     err.Source = TempErrMsg
  ElseIf Len(err.Source) <> 0 And Len(TempErrMsg) <> 0 Then
     err.Source = err.Source & vbCrLf & vbCrLf & TempErrMsg
  End If
  
  If Not mdlError.NoError Then
     mdlError.ReturnError PreMessage:="clsConnectivity.Class_Initialize: Input Errors!", _
                          PostMessage:="Review Configuration worksheet to correct this error.", _
                          IncludeMsgBox:=True
  Else
    'Run the connectivity function
    Connectivity
    DetermineFlowPaths
    DetermineGravityCVPaths
  End If
  
End Sub

Private Sub Connectivity()

Dim CVNum As Long
Dim NodeNum As Long
Dim i As Long
Dim j As Long

ReDim cM(1 To Horizontal_Divisions * Vertical_Divisions, 1 To 6)

CVNum = 1
NodeNum = 1

' Form the connectivity matrix
For i = 1 To Vertical_Divisions
   For j = 1 To Horizontal_Divisions
   
      ' Determine flow directions
      cM(CVNum, 5) = FlowDirection(i, j, Hot_Inflow_Edge, Hot_Inflow_Edge_Side, Hot_Reversals)
      cM(CVNum, 6) = FlowDirection(i, j, Cold_Inflow_Edge, Cold_Inflow_Edge_Side, Cold_Reversals)
      
      ' This algorithm counts edges starting with the horizontal bottom edge as edge 1 and works counter clockwise.  Any node that is hot is counted first on this edge
      ' followed by cold if there is overlap

      If (Hot_Inflow_Edge = 1 Or Hot_Inflow_Edge = 3) And (Cold_Inflow_Edge = 1 Or Cold_Inflow_Edge = 3) Then
         'We have vertically stacked nodes for hot and cold
         HotColdBothVertical cM, CVNum, NodeNum, i
         
      ElseIf (Hot_Inflow_Edge = 1 Or Hot_Inflow_Edge = 3) And (Cold_Inflow_Edge = 2 Or Cold_Inflow_Edge = 4) Then
         'We have vertically stacked hot nodes and horizontally stacked nodes for cold
         HotVerticalColdHorizontal cM, CVNum, NodeNum, i, j
      ElseIf (Hot_Inflow_Edge = 2 Or Hot_Inflow_Edge = 4) And (Cold_Inflow_Edge = 1 Or Cold_Inflow_Edge = 3) Then
         'We have horizontally stacked hot nodes and vertically stacked nodes for cold
         'this case is equivalent to the vertical, hot, horinzontal, cold with respect to connectivity.
         HotVerticalColdHorizontal cM, CVNum, NodeNum, i, j
      Else
         'We have horizontally stacked nodes for hot and cold
         HotColdBothHorizontal cM, CVNum, NodeNum, j
      End If
      
   Next
Next

' Form the boundary condition matrix
BoundaryConditions


End Sub

Private Sub BoundaryConditions()

' This function has been tested against 5 test configurations as follows
' Testing for Boundary Condition function
'Test Case:                      1   2   3   4   5
'Hot Inflow Edge                 2   4   4   2   1
'Hot Inflow Edge Side            1   1   0   1   1
'Hot Number of Reversals         1   3   3   2   1
'Cold Inflow Edge                1   3   3   4   3
'Cold Inflow Edge Side           0   0   1   0   0
'Cold Number of Reversals        6   2   1   0   3
'
'Number of Horizontal Divisions  7   3   4   4   4
'Number of Vertical Divisions    2   4   8   3   3
'
' All boundary conditions for these five cases were confirmed to be correct.
' Pictures of the boundary these boundary conditions are available on the teamforge repository md_model

' https://teamforge.sandia.gov/svn/repos/md_model/Testing/BoundaryConditionsTestCases.jpg?r=57

' and

' https://teamforge.sandia.gov/svn/repos/md_model/Testing/BoundaryConditionsTestCases2.jpg?r=57

      Dim BCCount As Long
      Dim ColdAndHotAligned As Boolean
      
      
      
      BCCount = 1
    ' this function quantifies the boundary condition matrix
    
      If (Hot_Inflow_Edge = 1 Or Hot_Inflow_Edge = 3) And (Cold_Inflow_Edge = 1 Or Cold_Inflow_Edge = 3) Then
            NumBC = 2 * Horizontal_Divisions
            NumHotBC = Horizontal_Divisions
            NumColdBC = Horizontal_Divisions
            ColdAndHotAligned = True
            
            NumHotInflowBC = Horizontal_Divisions / (Hot_Reversals + 1)
            NumColdInflowBC = Horizontal_Divisions / (Cold_Reversals + 1)
            
            
      ElseIf ((Hot_Inflow_Edge = 1 Or Hot_Inflow_Edge = 3) And (Cold_Inflow_Edge = 2 Or Cold_Inflow_Edge = 4)) Or _
             ((Hot_Inflow_Edge = 2 Or Hot_Inflow_Edge = 4) And (Cold_Inflow_Edge = 1 Or Cold_Inflow_Edge = 3)) Then
            NumBC = Horizontal_Divisions + Vertical_Divisions
            ColdAndHotAligned = False
            If (Hot_Inflow_Edge = 1 Or Hot_Inflow_Edge = 3) Then
               NumHotBC = Horizontal_Divisions
               NumColdBC = Vertical_Divisions
               
               NumHotInflowBC = Horizontal_Divisions / (Hot_Reversals + 1)
               NumColdInflowBC = Vertical_Divisions / (Cold_Reversals + 1)
               
            Else
               NumHotBC = Vertical_Divisions
               NumColdBC = Horizontal_Divisions
               
               NumHotInflowBC = Vertical_Divisions / (Hot_Reversals + 1)
               NumColdInflowBC = Horizontal_Divisions / (Cold_Reversals + 1)
            End If
      Else
            NumBC = 2 * Vertical_Divisions
            NumHotBC = Vertical_Divisions
            NumColdBC = Vertical_Divisions
            ColdAndHotAligned = True
            
            NumHotInflowBC = Vertical_Divisions / (Hot_Reversals + 1)
            NumColdInflowBC = Vertical_Divisions / (Cold_Reversals + 1)
      End If
      
      ReDim BC(1 To NumBC, 1 To 3)
      
      EdgeBoundaryCondition Hot_Inflow_Edge, Hot_Inflow_Edge_Side, Hot_Reversals, BCCount, True, ColdAndHotAligned
      EdgeBoundaryCondition Cold_Inflow_Edge, Cold_Inflow_Edge_Side, Cold_Reversals, BCCount, False, ColdAndHotAligned
      
End Sub

Private Sub EdgeBoundaryCondition(Inflow_Edge As Long, Inflow_Edge_Side As Long, NumberReversals As Long, BCCount As Long, IsHot As Boolean, ColdAndHotAligned As Boolean)

    Dim NumCVOnEdge As Long
    Dim ReversalCV As Long
    Dim CVNum As Long
    Dim NumRegion As Long
    Dim Region As Long
    Dim NumCVPerRegion As Long
    Dim r As Long 'indice for loop
    Dim c As Long 'indice for loop
    Dim e As Long 'edge indice
    Dim TInd As Long
   
    NumRegion = NumberReversals + 1
    
    'Determine the number of control volumes on the edge of interest.
    If Inflow_Edge = 1 Or Inflow_Edge = 3 Then
       NumCVOnEdge = Horizontal_Divisions
    Else
       NumCVOnEdge = Vertical_Divisions
    End If
    NumCVPerRegion = NumCVOnEdge / NumRegion  'Because of input protections this will always be an integer.
    
    Select Case Inflow_Edge
    ' Gather CV Numbers by region.
       Case 1
             DetermineRegionBC Inflow_Edge_Side, Inflow_Edge, NumRegion, Horizontal_Divisions, 0, 0, 1, IsHot, NumCVPerRegion, 0, BCCount, ColdAndHotAligned
       Case 2
            DetermineRegionBC Inflow_Edge_Side, Inflow_Edge, NumRegion, Vertical_Divisions, _
                              0, 0, Horizontal_Divisions, IsHot, NumCVPerRegion, 2, BCCount, ColdAndHotAligned
       Case 3
            DetermineRegionBC Inflow_Edge_Side, Inflow_Edge, NumRegion, Horizontal_Divisions, _
                              Horizontal_Divisions * (Vertical_Divisions - 1), 0, 1, IsHot, NumCVPerRegion, 1, BCCount, ColdAndHotAligned
       Case 4
            DetermineRegionBC Inflow_Edge_Side, Inflow_Edge, NumRegion, Vertical_Divisions, 1, 1, Horizontal_Divisions, IsHot, NumCVPerRegion, 3, BCCount, ColdAndHotAligned
    End Select
   
End Sub

Private Sub DetermineRegionBC(Inflow_Edge_Side As Long, Inflow_Edge As Long, NumRegion As Long, _
                              Divisions, Offset As Long, DivOffset As Long, Multiplier As Long, IsHot As Boolean, NumCVPerRegion As Long, _
                              InDirection As Long, BCCount As Long, ColdAndHotAligned As Boolean)
                              
    

    Dim c As Long ' column (oriented to edge) indice
    Dim ReversalCVColumn As Long
    Dim CVNum As Long
    Dim Region As Long
    Dim CVNumReversal As Long
    Dim Selector As Long
    Dim OppositeEdge As Long
    Dim OppositeCVNum As Long
    Dim OppositeCVOffset As Long
    Dim TemperatureCondition As Boolean
    Dim StartWithExitCondition As Boolean
    Dim OnOppositeEdge As Boolean
    Dim NumReversals As Long
    Dim TStart As Long
    Dim TEnd As Long
    Dim LastRegion As Long
    Dim TInd As Long
    Dim EdgeOffset As Long
    Dim CVNumToUse As Long
    
    'TInd = 5 if hot side, 6 if cold side
    If IsHot Then
       TInd = 5
    Else
       TInd = 6
    End If
    
    NumReversals = NumRegion - 1
    
    Select Case Inflow_Edge
       Case 1
          OppositeEdge = 3
          OppositeCVOffset = Horizontal_Divisions * (Vertical_Divisions - 1)
       Case 2
          OppositeEdge = 4
          OppositeCVOffset = -Horizontal_Divisions + 1
       Case 3
          OppositeEdge = 1
          OppositeCVOffset = -Horizontal_Divisions * (Vertical_Divisions - 1)
       Case 4
          OppositeEdge = 2
          OppositeCVOffset = Horizontal_Divisions - 1
    End Select
    
    ' Determine if an edge offset is needed because more than 1 node can be on an edge if hot and cold flow are aligned
    '
    If ColdAndHotAligned And IsHot Then
         If Inflow_Edge = 1 Or Inflow_Edge = 3 Then
             EdgeOffset = 0
         ElseIf Inflow_Edge = 2 Or Inflow_Edge = 4 Then
             EdgeOffset = -1
         End If
    ElseIf ColdAndHotAligned And Not IsHot Then
         If Inflow_Edge = 1 Or Inflow_Edge = 3 Then
             EdgeOffset = 1
         ElseIf Inflow_Edge = 2 Or Inflow_Edge = 4 Then
             EdgeOffset = 0
         End If
    Else
        EdgeOffset = 0
    End If
    
    
    'This inverts the orientation of Inflow_Edge_Side for edges 3 and 4
    If Inflow_Edge = 1 Or Inflow_Edge = 2 Then
       Selector = Inflow_Edge_Side
    Else
       Selector = 1 - Inflow_Edge_Side
    End If
    
    'Determine if we are starting with an exit condition (or the alternative start with a temperature condition)
    StartWithExitCondition = Selector = 1
    
    ' The opposite edge is dependent if NumReversals Mod 2 = 0.  Otherwise it is always starting on the opposite edge
    If StartWithExitCondition And (NumReversals Mod 2) = 0 Then
        OnOppositeEdge = False
    ElseIf Not StartWithExitCondition And (NumReversals Mod 2) = 0 Then
        OnOppositeEdge = True
    Else
        OnOppositeEdge = True
    End If
    
    'Temperature Loop start and end indices
    If StartWithExitCondition Then
       TStart = Divisions - NumCVPerRegion + 1
       TEnd = Divisions
    Else
       TStart = 1
       TEnd = NumCVPerRegion
    End If
    
    For c = TStart To TEnd
       CVNum = Multiplier * (c - DivOffset) + Offset
       BC(BCCount, 1) = cM(CVNum, Inflow_Edge + EdgeOffset)
       BC(BCCount, 2) = 0
       If IsHot Then
          BC(BCCount, 3) = 0 ' 0 means this is a hot side boundary condition
       Else
          BC(BCCount, 3) = 1 ' 1 means this is a cold side boundary condition
       End If
       BCCount = BCCount + 1
    Next
    
    LastRegion = mdlMath.Ceiling((NumCVPerRegion + 1) / Divisions * NumRegion)
    'Reversal Loop (only enters if NumCVPerRegion + 1 < Divisions
    For c = NumCVPerRegion + 1 To Divisions
        CVNum = Multiplier * (c - DivOffset) + Offset
        Region = mdlMath.Ceiling((c / Divisions) * NumRegion)
        ReversalCVColumn = (2 * (Region - 1) * NumCVPerRegion - c + 1)
        CVNumReversal = Multiplier * (ReversalCVColumn - DivOffset) + Offset
        
        'Weave between conditions
        If LastRegion <> Region Then
           OnOppositeEdge = Not OnOppositeEdge
        End If
                                            
        If OnOppositeEdge Then
            BC(BCCount, 1) = cM(CVNum + OppositeCVOffset, OppositeEdge + EdgeOffset)
            BC(BCCount, 2) = cM(CVNumReversal + OppositeCVOffset, OppositeEdge + EdgeOffset)
        Else
            BC(BCCount, 1) = cM(CVNum, Inflow_Edge + EdgeOffset)
            BC(BCCount, 2) = cM(CVNumReversal, Inflow_Edge + EdgeOffset)
        End If
        If IsHot Then
           BC(BCCount, 3) = 0 ' 0 means this is a hot side boundary condition
        Else
           BC(BCCount, 3) = 1 ' 1 means this is a cold side boundary condition
        End If
        BCCount = BCCount + 1
        
        LastRegion = Region
    Next c

End Sub

Private Function ReversalError(Divisions As Long, LCM As Long, VerticalOrHorizontal As String, Optional ErrMsg As String = "") As String

     If (Divisions Mod LCM <> 0) Then
        ReversalError = "The number of " & VerticalOrHorizontal & " divisions must be a multiple of the number of " & CStr(LCM) & "." & vbCrLf _
                     & "Valid input is: " & CStr(LCM) & ", " & CStr(2 * LCM) & ", " & CStr(3 * LCM) & ", ..."
     End If
     
     If Len(ErrMsg) <> 0 Then
        ReversalError = ErrMsg & vbCrLf & vbCrLf & ReversalError
     End If

End Function

Private Function FlowDirection(Row As Long, Column As Long, Inflow_Edge As Long, Inflow_Edge_Side As Long, NumReversals As Long)
    
    ' This function determines the flow direction for a given control volume's hot or cold sides.
    ' 0 = upward
    ' 1 = downward
    ' 2 = to the left
    ' 3 = to the right
    
    ' This function was tested 3/15/2017 against many test cases.
    
    Dim Region As Long
    Dim NumRegion As Long
    
    NumRegion = NumReversals + 1
    
    Select Case Inflow_Edge
        Case 1
            'Up/down 0/1
            Region = mdlMath.Ceiling((Column / Horizontal_Divisions) * NumRegion)
            FlowDirection = SelectFlowDirection(Inflow_Edge_Side, NumRegion, Region, Inflow_Edge)
        Case 2
            'left/right 2/3
            Region = mdlMath.Ceiling((Row / Vertical_Divisions) * NumRegion)
            FlowDirection = SelectFlowDirection(Inflow_Edge_Side, NumRegion, Region, Inflow_Edge, 2)
        Case 3
            'Up/down 0/1
            Region = mdlMath.Ceiling((Column / Horizontal_Divisions) * NumRegion)
            FlowDirection = SelectFlowDirection(Inflow_Edge_Side, NumRegion, Region, Inflow_Edge)
        Case 4
            'left/right 2/3
            Region = mdlMath.Ceiling((Row / Vertical_Divisions) * NumRegion)
            FlowDirection = SelectFlowDirection(Inflow_Edge_Side, NumRegion, Region, Inflow_Edge, 2)
    End Select

End Function

Private Function SelectFlowDirection(Inflow_Edge_Side As Long, NumRegion As Long, Region As Long, Inflow_Edge As Long, Optional Offset As Long = 0) As Long
    'This function selects up/down or right/left based on the region that the flow is in and the inflow edge_side
        
        Dim Selector As Long
        Dim AddOne As Long
        
        ' This factor is needed depending on whether there is an even or odd number of regions and whether the edge side
        If NumRegion Mod 2 = 0 Then
           If Inflow_Edge_Side = 1 Then
              AddOne = 0
           Else
              AddOne = 1
           End If
        Else
           If Inflow_Edge_Side = 1 Then
              AddOne = 1
           Else
              AddOne = 0
           End If
        End If
           
        ' This reverses the logic for edges 3 and 4 which are upside down. w/r to the other edges.
        If Inflow_Edge = 1 Or Inflow_Edge = 2 Then
            Selector = Inflow_Edge_Side
        Else
            Selector = 1 - Inflow_Edge_Side
        End If
        
        Select Case Selector
            Case 0
                SelectFlowDirection = ((NumRegion - Region + AddOne) Mod 2) + Offset
            Case 1
                SelectFlowDirection = ((Region + AddOne) Mod 2) + Offset
        End Select
End Function


Private Sub HotColdBothHorizontal(cM() As Long, CVNum As Long, NodeNum As Long, Column As Long)
'CVNum = Control Volume number
If Column = 1 Then 'All Numbers are brand new
    cM(CVNum, 1) = NodeNum
    cM(CVNum, 2) = NodeNum + 1
    cM(CVNum, 3) = NodeNum + 2
    cM(CVNum, 4) = NodeNum + 3
    NodeNum = NodeNum + 4
    ' Now find the flow direction for the hot side
    
Else 'Edge 4 nodes are repeats
    cM(CVNum, 1) = NodeNum
    cM(CVNum, 2) = NodeNum + 1
    cM(CVNum, 3) = cM(CVNum - 1, 1)
    cM(CVNum, 4) = cM(CVNum - 1, 2)
    NodeNum = NodeNum + 2
End If
CVNum = CVNum + 1

End Sub

Private Sub HotVerticalColdHorizontal(cM() As Long, CVNum As Long, NodeNum As Long, Row As Long, Column As Long)
'CVNum = Control Volume number
If Row = 1 And Column = 1 Then 'All numbers are brand new
    cM(CVNum, 1) = NodeNum
    cM(CVNum, 2) = NodeNum + 1
    cM(CVNum, 3) = NodeNum + 2
    cM(CVNum, 4) = NodeNum + 3
    NodeNum = NodeNum + 4
ElseIf Row = 1 Then ' Edge 1, 2, and 3 nodes are brand new
    cM(CVNum, 1) = NodeNum
    cM(CVNum, 2) = NodeNum + 1
    cM(CVNum, 3) = NodeNum + 2
    cM(CVNum, 4) = cM(CVNum - 1, 2)
    NodeNum = NodeNum + 3
ElseIf Column = 1 Then 'Edge 2, 3, and 4 are brand new
    cM(CVNum, 1) = cM(CVNum - Horizontal_Divisions, 3)
    cM(CVNum, 2) = NodeNum
    cM(CVNum, 3) = NodeNum + 1
    cM(CVNum, 4) = NodeNum + 2
    NodeNum = NodeNum + 3
Else 'Edge 2 and 3 are brand new
    cM(CVNum, 1) = cM(CVNum - Horizontal_Divisions, 3)
    cM(CVNum, 2) = NodeNum
    cM(CVNum, 3) = NodeNum + 1
    cM(CVNum, 4) = cM(CVNum - 1, 2)
    NodeNum = NodeNum + 2
End If
CVNum = CVNum + 1

End Sub

Private Sub HotColdBothVertical(cM() As Long, CVNum As Long, NodeNum As Long, Row As Long)
'CVNum = Control Volume number
If Row = 1 Then ' All numbers are brand new
    cM(CVNum, 1) = NodeNum
    cM(CVNum, 2) = NodeNum + 1
    cM(CVNum, 3) = NodeNum + 2
    cM(CVNum, 4) = NodeNum + 3
    NodeNum = NodeNum + 4
Else
    cM(CVNum, 1) = cM(CVNum - Horizontal_Divisions, 3)
    cM(CVNum, 2) = cM(CVNum - Horizontal_Divisions, 4)
    cM(CVNum, 3) = NodeNum
    cM(CVNum, 4) = NodeNum + 1
    NodeNum = NodeNum + 2
End If
CVNum = CVNum + 1

End Sub

Private Function NumNodeInFlowPath(Reversals As Long, FlowsHorizontal As Boolean)
    If FlowsHorizontal Then
       NumNodeInFlowPath = (Horizontal_Divisions + 1) * (Reversals + 1)
    Else
       NumNodeInFlowPath = (Vertical_Divisions + 1) * (Reversals + 1)
    End If
End Function

Public Sub DetermineFlowPaths()

Dim NumNodeInHotFlowPath As Long
Dim NumNodeInColdFlowPath As Long
Dim HotInd As Long
Dim ColdInd As Long
Dim i As Long

NumNodeInHotFlowPath = NumNodeInFlowPath(Hot_Reversals, HotFlowsHorizontal)
NumNodeInColdFlowPath = NumNodeInFlowPath(Cold_Reversals, ColdFlowsHorizontal)
 
' This subroutine extracts the flow path
ReDim HotFlowPaths(1 To NumNodeInHotFlowPath, 1 To HotChannelWidth)
ReDim ColdFlowPaths(1 To NumNodeInColdFlowPath, 1 To ColdChannelWidth)
ReDim HotCVFlowPaths(1 To NumNodeInHotFlowPath - (Hot_Reversals + 1), 1 To HotChannelWidth)
ReDim ColdCVFlowPaths(1 To NumNodeInColdFlowPath - (Cold_Reversals + 1), 1 To ColdChannelWidth)
HotInd = 1
ColdInd = 1

For i = 1 To UBound(BC, 1)
   If BC(i, 2) = 0 Then ' We have an inflow boundary condition
       If BC(i, 3) = 0 Then ' This is a hot boundary condition
          DetermineSingleFlowPathFromStartingBC BC(i, 1), HotFlowPaths, HotInd, HotFlowsHorizontal, 5, Hot_Reversals, HotCVFlowPaths
          HotInd = HotInd + 1
       Else ' This is a cold boundary condition
          DetermineSingleFlowPathFromStartingBC BC(i, 1), ColdFlowPaths, ColdInd, ColdFlowsHorizontal, 6, Cold_Reversals, ColdCVFlowPaths
          ColdInd = ColdInd + 1
       End If
   End If
Next

End Sub

Private Sub DetermineSingleFlowPathFromStartingBC(ByVal StartNode As Long, FlowPaths() As Long, FPInd As Long, FlowIsHorizontal As Boolean, _
                                                  DirInd As Long, Reversals As Long, CVFlowPaths() As Long)
' This function is only designed to be used in the context of "DetermineFlowPaths"
' FlowPaths and CVFlowPaths must already be initialized to the correct size or else a failure will occur.
' DirInd = indicates the direction index for CM (5 - hot flow, 6 - cold flow)

Dim CV As Long
Dim i As Long
Dim j As Long
Dim count As Long
Dim CVcount As Long
Dim MaxCount As Long
Dim CVInc As Long
Dim Edge As Long
Dim OpEdge As Long

' Find the control volume of the start node

If FlowIsHorizontal Then
   MaxCount = Horizontal_Divisions - 1
Else
   MaxCount = Vertical_Divisions - 1
End If

count = 1
CVcount = 1
For i = 1 To Reversals + 1
   ' Find the starting control volume
   CV = ControlVolumeGivenNodeNumber(StartNode, Edge)
   
   ' Establish direction to determine how to increment CV number.
   CVInc = CVIncrement(cM(CV, DirInd))

   
   ' Enter the start node for the current leg.
   FlowPaths(count, FPInd) = StartNode
   CVFlowPaths(CVcount, FPInd) = CV
   count = count + 1
   CVcount = CVcount + 1
   ' Now fill in all the nodes for the current flow path
   For j = 1 To MaxCount
      CV = CV + CVInc
      FlowPaths(count, FPInd) = cM(CV, Edge)
      CVFlowPaths(CVcount, FPInd) = CV
      count = count + 1
      CVcount = CVcount + 1
   Next j
   
   ' Add the node on the opposite edge which must also be a reversal boundary condition node
   ' unless it is an exit node.  This extra node is not needed by the CV flow paths array.
   
   OpEdge = OppositeEdge(Edge)
   
   FlowPaths(count, FPInd) = cM(CV, OpEdge)
   count = count + 1
   
   'Now find the new leg start node
   StartNode = FindReversalBCNextNode(cM(CV, OpEdge))

Next i

End Sub

Private Function CVIncrement(Dir As Long)
   If Dir = 0 Then
      CVIncrement = Horizontal_Divisions
   ElseIf Dir = 1 Then
      CVIncrement = -Horizontal_Divisions
   ElseIf Dir = 2 Then
      CVIncrement = -1
   ElseIf Dir = 3 Then
      CVIncrement = 1
   End If
End Function

Private Function FindReversalBCNextNode(Node As Long) As Long
  
  Dim i As Long
  
  For i = 1 To UBound(BC, 1)
      If BC(i, 2) <> 0 Then
         If BC(i, 1) = Node Then
            FindReversalBCNextNode = BC(i, 2)
Exit Function
         ElseIf BC(i, 2) = Node Then
            FindReversalBCNextNode = BC(i, 1)
Exit Function
         End If
      End If
  Next
  
  FindReversalBCNextNode = -1 ' This indicates that the requested node is not part of a reversal boundary condition.

End Function

Private Function OppositeEdge(Edge As Long) As Long
   Dim Dum As Double
   If Edge = 1 Then
      OppositeEdge = 3
   ElseIf Edge = 2 Then
      OppositeEdge = 4
   ElseIf Edge = 3 Then
      OppositeEdge = 1
   ElseIf Edge = 4 Then
      OppositeEdge = 2
   Else
      Dum = mdlError.ReturnError("clsConnectivity.OppositeEdge: An invalid edge number " & CStr(Edge) & _
                                 " has been requested.  Only edges 1 to 4 exist for this connectivity!", , True)
   End If
End Function

Public Function ControlVolumeGivenNodeNumber(NodeNum As Long, Optional EdgeNum As Long) As Long
' This is a brute force search approach that probably could be made a lot quicker but would require a lot of careful thinking
' This routine returns the earliest Control volume that contains a node number.

'EdgeNum is also an optional output that provides the edge number that the original node is from

     Dim i As Long
     Dim j As Long
     Dim n As Long
     Dim m As Long
     Dim Dum As Double
     
     n = UBound(cM, 1)
     m = 4
     
     If NodeNum > Me.NumberOfNodes Then
        Dum = mdlError.ReturnError("clsConnectivity.ControlVolumeGivenNodeNumber: The requested node number (" & CStr(NodeNum) & _
                             ") is greater than the number of nodes " & CStr(Me.NumberOfNodes) & "!", , True)
     End If
     
     For i = 1 To n
        For j = 1 To m
            If cM(i, j) = NodeNum Then
               ControlVolumeGivenNodeNumber = i
               EdgeNum = j
Exit Function
            End If
        Next
     Next
     ' This point should never be reached
     Dum = mdlError.ReturnError("clsConnectivity.ControlVolumeGivenNodeNumber: A node is missing from the connectivity " & _
                          "matrix? Something is wrong with the connectivity! This is a strange error!", , True)

End Function

Public Function GravityLength(Inp As clsInput) As Double
    If GravDirection = 1 Or GravDirection = 3 Then
        GravityLength = Inp.VerticalLength
    Else
        GravityLength = Inp.HorizontalLength
    End If
End Function

Private Sub DetermineGravityCVPaths()

Dim Dum As Double
Dim start_cv As Long
Dim inc_h As Long
Dim inc_v As Long
Dim i As Long
Dim j As Long
Dim num_cv_in_path As Long
Dim num_paths As Long

Select Case GravDirection
     Case 1
         ReDim GravityCVPaths(1 To Vertical_Divisions, 1 To Horizontal_Divisions)
         num_cv_in_path = Vertical_Divisions
         num_paths = Horizontal_Divisions
         start_cv = 1
         inc_h = 1
         inc_v = Horizontal_Divisions
     Case 2
         ReDim GravityCVPaths(1 To Horizontal_Divisions, 1 To Vertical_Divisions)
         num_cv_in_path = Horizontal_Divisions
         num_paths = Vertical_Divisions
         start_cv = Horizontal_Divisions
         inc_h = Horizontal_Divisions
         inc_v = -1
     Case 3
         ReDim GravityCVPaths(1 To Vertical_Divisions, 1 To Horizontal_Divisions)
         num_cv_in_path = Vertical_Divisions
         num_paths = Horizontal_Divisions
         start_cv = Horizontal_Divisions * Vertical_Divisions
         inc_h = -1
         inc_v = -Horizontal_Divisions
     Case 4  'Cases 5 and 6 may be able to be worked out (and have significant implications for pressure
             ' losses but pressure losses are not yet modeled. Also the air gap mass transport would need to be
             ' enhanced in order to accomodate this.
         ReDim GravityCVPaths(1 To Horizontal_Divisions, 1 To Vertical_Divisions)
         num_cv_in_path = Horizontal_Divisions
         num_paths = Vertical_Divisions
         start_cv = Horizontal_Divisions * (Vertical_Divisions - 1) + 1
         inc_h = -Horizontal_Divisions
         inc_v = 1
     Case Else
         Dum = mdlError.ReturnError("clsConnectivity.DetermineGravityCVPaths: A Gravity direction of " & _
                   CStr(GravDirection) & " was entered only values of 1-4 are currently accepted. Directions 5 and 6 (normal to plane) have not been developed", , True)
End Select

For i = 1 To num_paths
   If i <> 1 Then
      start_cv = start_cv + inc_h
   End If
   For j = 1 To num_cv_in_path
       GravityCVPaths(j, i) = start_cv + (j - 1) * inc_v
   Next
Next

End Sub

Public Property Get GravityControlVolumePath(layer As Long, path As Long) As Long
     GravityControlVolumePath = GravityCVPaths(layer, path)
End Property

Public Property Get NumberOfGravityLayers() As Long
    NumberOfGravityLayers = UBound(GravityCVPaths, 1)
End Property

Public Property Get NumberOfGravityPaths() As Long
   NumberOfGravityPaths = UBound(GravityCVPaths, 2)
End Property

Public Property Get FindPreviousGravityCVGivenCVNum(CVNum As Long) As Long
   Dim i As Long
   Dim j As Long
   Dim Dum As Double
   
   If CVNum > NumberOfControlVolumePairs Or CVNum < 0 Then
       Dum = mdlError.ReturnError("clsConnectivity.FindPreviousGravityCVGivenCVNum: A control volume pair number beyond the number of control volume pairs was requested", , True)
   Else
        For i = 1 To NumberOfGravityLayers
           For j = 1 To NumberOfGravityPaths
               If GravityCVPaths(i, j) = CVNum Then
                   GoTo FoundCVNum
               End If
           Next
        Next
        ' Oops we should never reach this point.
        Dum = mdlError.ReturnError("clsConnectivity.FindPreviousGravityCVGivenCVNum: The internal array GravityCVPaths does not contain control volume pair number " & _
                                CStr(CVNum) & " there must be a bug in how GravityCVPaths is formed!", , True)
   End If
FoundCVNum:
   If i = 1 Then 'There is no previous CV number above this one.  This CV is in layer 1
       FindPreviousGravityCVGivenCVNum = glbINVALID_VALUE
   Else
       FindPreviousGravityCVGivenCVNum = GravityCVPaths(i - 1, j)
   End If
       
End Property


