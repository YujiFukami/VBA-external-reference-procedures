VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExtRef 
   Caption         =   "外部参照プロシージャ"
   ClientHeight    =   7428
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10560
   OleObjectBlob   =   "frmExtRef.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmExtRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PbVBProjectNameList$()
Dim PbExtProcedureList
Dim PbTmpExtProcedureDict

Private Sub listVBProject_Click()
 
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Me.list外部参照.Clear
    Me.txtCode.Text = ""
    Dim TmpExtProcedureDict As Object, TmpExtProcedureList
    
    For I = 1 To UBound(PbVBProjectNameList, 1)
        Select Case listVBProject.List(listVBProject.ListIndex)
        Case PbVBProjectNameList(I)
            '選択したVBProjectで外部参照しているプロシージャを表示
            Set PbTmpExtProcedureDict = PbExtProcedureList(I)
            TmpExtProcedureList = PbTmpExtProcedureDict.Keys
            
            Me.list外部参照.List = TmpExtProcedureList
                
        End Select
    Next I

End Sub
Private Sub list外部参照_Click()

 
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Me.txtCode.Text = ""
    Dim TmpCode$
    Dim TmpProcedureList
    TmpProcedureList = PbTmpExtProcedureDict.Keys
    TmpProcedureList = Application.Transpose(Application.Transpose(TmpProcedureList))
    
    For I = 1 To UBound(TmpProcedureList, 1)
        Select Case Me.list外部参照.List(Me.list外部参照.ListIndex)
        Case TmpProcedureList(I)
            '選択したプロシージャのコードを表示
            TmpCode = PbTmpExtProcedureDict(TmpProcedureList(I))
            Me.txtCode.Text = TmpCode
                
        End Select
    Next I
    
End Sub

Private Sub UserForm_Initialize()

    Dim VBProjectList() As classVBProject
    VBProjectList = フォーム用VBProject作成
    
    Dim AllProcedureList
    AllProcedureList = 全プロシージャ一覧作成(VBProjectList)
    Call プロシージャ内の使用プロシージャ取得(VBProjectList, AllProcedureList)
    
    Dim ExtProcedureList
    ExtProcedureList = 外部参照プロシージャリスト作成(VBProjectList)
    
    'フォーム用パブリック変数設定
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    N = UBound(VBProjectList, 1)
    ReDim PbVBProjectNameList(1 To N)
    For I = 1 To N
        PbVBProjectNameList(I) = VBProjectList(I).Name
    Next I
    
    PbExtProcedureList = ExtProcedureList
        
    'フォーム設定
    Me.listVBProject.List = PbVBProjectNameList
        
End Sub
