VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHTFuncLosses"
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

Private Const IncreaseFactor = 1 ' THIS A PURELY EMPERICAL ADJUSTMENT FACTOR TO THE TOTAL HEAT FLOW.  RETURN IT TO A VaLUE OF 1 !@#$!!!
Private PrNeedToReadConstants As Boolean
Private IsFirstTimeThrough As Boolean

Private T_Air As Double 'Ambient air temperature (K)
Private T_water As Double 'Water flow temperature (may be hot or cold side) (K)
Private M_water As Double 'Water flow temperature (may be hot or cold side) (kg/s)
Private S_water As Double 'Water salinity (grams/gram)
Private S_iw As Double ' Water Interface Salinity (grams/gram)

Private T_ia As Double 'Interface temperature (temperature at membrane wall) on the hot side (K)
Private T_iw As Double 'Interface temperature (temperature at membrane wall) on the cold side (K)
Private S_i As Double 'Interface Salinity (grams/gram)

Private P_w As Double 'Pressure water (Pa)

Private U As Double  'Overall heat transfer coefficient (W/(m2*K))
Private A_f As Double ' mass transfer area

Private Q As Double ' heat transfer rate via effectiveness method

Private L_Norm As Double ' Length normal to flow cold side (m)
Private Lz As Double ' thickness of cold side (from cold side spacers) (m)
Private L_TotalPara As Double ' Length parallel to flow cold side (m)

Private L_gravity As Double ' length in the vertical direction of the model

Private Hii As Double 'Heat transfer coefficient across the insulation material
Private Hwi As Double 'Heat transfer coefficient through water
Private Hai As Double 'Heat transfer through natural convection layer

Private ExtMat As clsMaterial ' External insulation material

Private WaterSpacer As clsSpacer ' Hot and cold spacers

Public Property Let NeedToReadConstants(Val As Boolean)
    PrNeedToReadConstants = Val
End Property

Public Property Get HeatTransferRate() As Double
    HeatTransferRate = Q
End Property

' This function name must be shared by all objects
' It represents the set of nonlinear local equations that must be solved
' In order to resolve terms in the global energy and mass transfer equations.
' There can be many different setups. In this case, the solution could be accomplished through
' successive substitution since a solution for the interface temperatures is straightforward but
' I used Newton's method anyway to keep the methodology generalized and to test the methods on a
' very simple case before moving to more complicated methods.
Public Sub EvaluateFunction(x() As Double, CV As clsControlVolumePair, SysEq As clsSystemEquations, Result() As Double, _
                                               Optional IsHotSide As Boolean, _
                                               Optional Tair As Double, _
                                               Optional Twater As Double, _
                                               Optional Mwater As Double, _
                                               Optional Swater As Double, _
                                               Optional Si_ As Double)
' Use optional variables to set values to provide a means of initialization.
    Dim Dum As Double
    If PrNeedToReadConstants Then
      If (IsMissing(IsHotSide) Or IsMissing(Tair) Or IsMissing(Twater) Or IsMissing(Mwater) Or IsMissing(Swater) Or _
         IsMissing(Si_)) Then
            Dum = mdlError.ReturnError("clsDirectContactHTFunc.EvaluateHeatTransfer: All optional arguments must be included if this is the first time that the class is being called. " _
                                      & " Subsequent calls do not need the optional arguments.", , True)
      Else
        ' Initialize values (some of these change and others stay the same value)
        T_Air = Tair
        T_water = Twater
        M_water = Mwater
        S_water = Swater
        S_i = Si_
         
         ' These are values that NEVER change and take more time to extract over and over again
         If IsFirstTimeThrough Then
            
            A_f = CV.Area
           
            Dim WS As clsWaterStream
           
            If IsHotSide Then
               Set WS = SysEq.Inputs.WaterStreams(SysEq.Inputs.HotWaterStream)
               P_w = WS.Pressure
               Set WaterSpacer = SysEq.HotSpacer
               L_Norm = CV.LengthNormalToHotFlow
               Lz = CV.HotThickness
               L_TotalPara = CV.TotalLengthParrallelToHotFlow
            Else
               Set WS = SysEq.Inputs.WaterStreams(SysEq.Inputs.ColdWaterStream)
               P_w = WS.Pressure
               Set WaterSpacer = SysEq.ColdSpacer
               L_Norm = CV.LengthNormalToColdFlow
               Lz = CV.ColdThickness
               L_TotalPara = CV.TotalLengthParrallelToColdFlow
            End If
            L_gravity = SysEq.Connectivity.GravityLength(SysEq.Inputs)
            Set ExtMat = SysEq.Inputs.ExternalInsulationMaterial
          End If
         

         IsFirstTimeThrough = False ' This never gets reset
         PrNeedToReadConstants = False ' This gets reset for each solution iteration
      End If
    Else
       ' Only the variables have to be updated.  Everything else is constant or gets recalculated.
        T_iw = x(1)
        T_ia = x(2)
        
        'Overall heat transfer Coefficient (not used anymore - we use the local heat transfer coefficients
        U = mdlTransferCoefficient.HTC_Natural_Convection(T_Air, T_water, M_water, T_ia, T_iw, Hwi, Hii, Hai, L_Norm, WaterSpacer, Lz, _
                                                          L_TotalPara, A_f, L_gravity, ExtMat, S_water, S_i)
        
        ReDim Result(1 To 2)
        ' Calculate the total heat transfer several ways these are the equations that need to be driven to zero if the interface temperatures are consistent
        Result(1) = Hai * (T_ia - T_Air) - Hii * (T_iw - T_ia)
        Result(2) = Hwi * (T_water - T_iw) - Hii * (T_iw - T_ia)
        
        Q = IncreaseFactor * U * A_f * (T_water - T_Air)
        
    End If
End Sub

Public Sub ConvergenceCriteria(ConvergCrit() As Double, CVArea As Double)

ReDim ConvergCrit(1 To 2)
Dim EnergyCritOverArea As Double

EnergyCritOverArea = mdlConstants.glbAbsConvergCriterionEnergy / CVArea

ConvergCrit(1) = EnergyCritOverArea
ConvergCrit(2) = EnergyCritOverArea

End Sub

Private Sub Class_Initialize()
   PrNeedToReadConstants = True
   IsFirstTimeThrough = True
End Sub




