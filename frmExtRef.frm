VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExtRef 
   Caption         =   "�O���Q�ƃv���V�[�W��"
   ClientHeight    =   7428
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10560
   OleObjectBlob   =   "frmExtRef.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
 
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Me.list�O���Q��.Clear
    Me.txtCode.Text = ""
    Dim TmpExtProcedureDict As Object, TmpExtProcedureList
    
    For I = 1 To UBound(PbVBProjectNameList, 1)
        Select Case listVBProject.List(listVBProject.ListIndex)
        Case PbVBProjectNameList(I)
            '�I������VBProject�ŊO���Q�Ƃ��Ă���v���V�[�W����\��
            Set PbTmpExtProcedureDict = PbExtProcedureList(I)
            TmpExtProcedureList = PbTmpExtProcedureDict.Keys
            
            Me.list�O���Q��.List = TmpExtProcedureList
                
        End Select
    Next I

End Sub
Private Sub list�O���Q��_Click()

 
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Me.txtCode.Text = ""
    Dim TmpCode$
    Dim TmpProcedureList
    TmpProcedureList = PbTmpExtProcedureDict.Keys
    TmpProcedureList = Application.Transpose(Application.Transpose(TmpProcedureList))
    
    For I = 1 To UBound(TmpProcedureList, 1)
        Select Case Me.list�O���Q��.List(Me.list�O���Q��.ListIndex)
        Case TmpProcedureList(I)
            '�I�������v���V�[�W���̃R�[�h��\��
            TmpCode = PbTmpExtProcedureDict(TmpProcedureList(I))
            Me.txtCode.Text = TmpCode
                
        End Select
    Next I
    
End Sub

Private Sub UserForm_Initialize()

    Dim VBProjectList() As classVBProject
    VBProjectList = �t�H�[���pVBProject�쐬
    
    Dim AllProcedureList
    AllProcedureList = �S�v���V�[�W���ꗗ�쐬(VBProjectList)
    Call �v���V�[�W�����̎g�p�v���V�[�W���擾(VBProjectList, AllProcedureList)
    
    Dim ExtProcedureList
    ExtProcedureList = �O���Q�ƃv���V�[�W�����X�g�쐬(VBProjectList)
    
    '�t�H�[���p�p�u���b�N�ϐ��ݒ�
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    N = UBound(VBProjectList, 1)
    ReDim PbVBProjectNameList(1 To N)
    For I = 1 To N
        PbVBProjectNameList(I) = VBProjectList(I).Name
    Next I
    
    PbExtProcedureList = ExtProcedureList
        
    '�t�H�[���ݒ�
    Me.listVBProject.List = PbVBProjectNameList
        
End Sub
