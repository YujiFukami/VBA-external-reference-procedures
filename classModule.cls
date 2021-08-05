VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "classModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name$
Public Procedures As Collection
Public VBProjectName$

Property Get MyName()
    MyName = Name
End Property

Property Let MyName(InputName)
    Name = InputName
End Property

Private Sub Class_Initialize()
    Set Procedures = New Collection
End Sub

Public Sub AddProcedure(InputProcedure As ClassProcedure)
    Procedures.Add InputProcedure
End Sub

Public Function Item(Index&) As ClassProcedure
    Set Item = Procedures(Index)
End Function

Property Get MyVBProjectName()
    MyVBProjectName = VBProjectName
End Property

Property Let MyVBProjectName(InputVBProjectName)
    VBProjectName = InputVBProjectName
End Property

