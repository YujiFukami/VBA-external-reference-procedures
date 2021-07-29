VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "classProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name$
Public Code
Public KensakuCode As Object '使用プロシージャ検索用連想配列
Public UseProcedure As Collection
Public ModuleName$
Public VBProjectName$

Private Sub Class_Initialize()
    Set UseProcedure = New Collection
End Sub

Property Get MyName()
    MyName = Name
End Property
Property Let MyName(InputName)
    Name = InputName
End Property

Property Get MyCode()
    MyCode = Code
End Property
Property Let MyCode(InputCode)
    Code = InputCode
End Property

Property Get MyKensakuCode()
    Set MyKensakuCode = KensakuCode
End Property
Property Set MyKensakuCode(InputKensakuCode As Object)
    Set KensakuCode = InputKensakuCode
End Property

Public Sub AddUseProcedure(InputUseProcedure As ClassProcedure)
    UseProcedure.Add InputUseProcedure
End Sub
Public Function Item(Index&) As ClassProcedure
    Set Item = UseProcedure(Index)
End Function

Property Get MyModuleName()
    MyModuleName = ModuleName
End Property
Property Let MyModuleName(InputModuleName)
    ModuleName = InputModuleName
End Property

Property Get MyVBProjectName()
    MyVBProjectName = VBProjectName
End Property
Property Let MyVBProjectName(InputVBProjectName)
    VBProjectName = InputVBProjectName
End Property

