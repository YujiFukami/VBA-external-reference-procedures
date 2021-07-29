Attribute VB_Name = "ModExtProcedure"
Option Explicit
'�O���Q�ƃv���V�[�W���̎擾�p���W���[��
'frmExtRef�ƘA�g���Ă���

Function �t�H�[���pVBProject�쐬()
    
    Dim I%, J%, II%, K%, M%, N% '�����グ�p(Integer�^)
    Dim OutputVBProjectList() As classVBProject
    Dim TmpClassVBProject As classVBProject
    Dim TmpClassModule As classModule
    Dim TmpClassProcedure As ClassProcedure
    Dim VBProjectList As VBProjects, TmpVBProject As VBProject
    Dim TmpModule As VBComponent, TmpProcedureNameList, TmpCodeDict As Object
    Dim TmpProcedureName$
    
    Set VBProjectList = ActiveWorkbook.VBProject.VBE.VBProjects
    ReDim OutputVBProjectList(1 To VBProjectList.Count)
    For I = 1 To VBProjectList.Count
        Set TmpVBProject = VBProjectList(I)
        Set TmpClassVBProject = New classVBProject
        TmpClassVBProject.MyName = TmpVBProject.Name
        
        For J = 1 To TmpVBProject.VBComponents.Count
            Set TmpClassModule = New classModule
            Set TmpModule = TmpVBProject.VBComponents(J)
            
            TmpClassModule.Name = TmpModule.Name
            TmpClassModule.VBProjectName = TmpClassVBProject.Name
            
            TmpProcedureNameList = ���W���[���̃v���V�[�W�����ꗗ�擾(TmpModule)
            Set TmpCodeDict = ���W���[���̃R�[�h�ꗗ�擾(TmpModule)
            If IsEmpty(TmpProcedureNameList) = False Then
                For II = 1 To UBound(TmpProcedureNameList)
                    Set TmpClassProcedure = New ClassProcedure
                    TmpProcedureName = TmpProcedureNameList(II)
                    TmpClassProcedure.Name = TmpProcedureName
                    TmpClassProcedure.Code = TmpCodeDict(TmpProcedureName)
                    Set TmpClassProcedure.KensakuCode = �R�[�h�������p�ɕύX(TmpCodeDict(TmpProcedureName))
                    TmpClassProcedure.VBProjectName = TmpClassVBProject.Name
                    TmpClassProcedure.ModuleName = TmpClassModule.Name
                    TmpClassModule.AddProcedure TmpClassProcedure
                Next II
            End If
            
            TmpClassVBProject.AddModule TmpClassModule
            
        Next J
        
        Set OutputVBProjectList(I) = TmpClassVBProject
        
    Next I
    
    �t�H�[���pVBProject�쐬 = OutputVBProjectList
    
End Function
Function ���W���[���̃v���V�[�W�����ꗗ�擾(InputModule As VBComponent)
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim TmpStr$
    Dim Output
    ReDim Output(1 To 1)
    With InputModule.CodeModule
        K = 0
        For I = 1 To .CountOfLines
            If TmpStr <> .ProcOfLine(I, 0) Then
                TmpStr = .ProcOfLine(I, 0)
                K = K + 1
                ReDim Preserve Output(1 To K)
                Output(K) = TmpStr
            End If
        Next I
    End With
    
    If K = 0 Then '���W���[�����Ƀv���V�[�W�����Ȃ��ꍇ
        Output = Empty
    End If
    
    ���W���[���̃v���V�[�W�����ꗗ�擾 = Output
        
End Function
Function ���W���[���̃R�[�h�ꗗ�擾(InputModule As VBComponent)
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim ProcedureList
    ProcedureList = ���W���[���̃v���V�[�W�����ꗗ�擾(InputModule)
    Dim Output As Object
    Dim TmpProcedureName$, TmpStart&, TmpEnd&, TmpCode$
    If IsEmpty(ProcedureList) Then
        '�v���V�[�W������
        Set Output = Nothing
    Else
        '�v���V�[�W���L��
        N = UBound(ProcedureList, 1)
        Set Output = CreateObject("Scripting.Dictionary")
        For I = 1 To N
            TmpProcedureName = ProcedureList(I)
            With InputModule.CodeModule
                On Error Resume Next
                TmpStart = 0
                
                TmpStart = .ProcBodyLine(TmpProcedureName, 0)
                TmpEnd = .ProcCountLines(TmpProcedureName, 0)
                      
                If TmpStart = 0 Then '�N���X���W���[���̃R�[�h�擾�p
                    TmpStart = .ProcBodyLine(TmpProcedureName, vbext_pk_Get)
                    TmpEnd = .ProcCountLines(TmpProcedureName, vbext_pk_Let)
                End If
                
                On Error GoTo 0
                
                TmpCode = .Lines(TmpStart, TmpEnd)
            End With
            
            Output.Add TmpProcedureName, TmpCode
        Next I
    End If
    
    Set ���W���[���̃R�[�h�ꗗ�擾 = Output

End Function
Function �R�[�h�������p�ɕύX(InputCode) As Object
    
    Dim CodeList, TmpStr$
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    CodeList = Split(InputCode, vbLf)
    CodeList = Application.Transpose(CodeList)
    CodeList = Application.Transpose(CodeList)
    N = UBound(CodeList, 1)
    
    Dim BunkatuStrList, HenkanStr$, TmpBunkatu
    BunkatuStrList = Array(" ", ":", "_", ",", """", "(", ")")
    BunkatuStrList = Application.Transpose(Application.Transpose(BunkatuStrList))
    HenkanStr = Chr(13)
    
    Dim BunkatuDict As Object
    Set BunkatuDict = CreateObject("Scripting.Dictionary")
    Dim Output As Object
    Set Output = CreateObject("Scripting.Dictionary")
    
    For I = 1 To N
        TmpStr = CodeList(I)
        TmpStr = Trim(TmpStr) '���E�̋󔒏���
        TmpStr = StrConv(TmpStr, vbUpperCase) '�������ɕϊ�
'        TmpStr = StrConv(TmpStr, vbNarrow) '���p�ɕϊ�
        If InStr(1, TmpStr, "'") > 0 Then
            TmpStr = Split(TmpStr, "'")(0) '�R�����g�̏���
        End If
        TmpStr = Replace(TmpStr, Chr(13), "") '���s������
        
        
        If TmpStr <> "" Then
            '�w�蕶���ŕ�������
            For J = 1 To UBound(BunkatuStrList, 1)
                TmpStr = Replace(TmpStr, BunkatuStrList(J), HenkanStr)
            Next J
            TmpBunkatu = Split(TmpStr, HenkanStr)
            
            For J = 0 To UBound(TmpBunkatu)
                If BunkatuDict.Exists(TmpBunkatu(J)) = False Then
                    BunkatuDict.Add TmpBunkatu(J), ""
                End If
            Next J
        End If
    Next I
    
    Set Output = BunkatuDict
    Set �R�[�h�������p�ɕύX = Output

End Function
Function �S�v���V�[�W���ꗗ�쐬(VBProjectList)
    
    Dim I&, J&, II&, K&, M&, N& '�����グ�p(Long�^)
    Dim ProcedureCount&
    '�v���V�[�W���̌����v�Z
    Dim TmpClassVBProject As classVBProject
    Dim TmpClassModule As classModule
    Dim TmpClassProcedure As ClassProcedure
    
    ProcedureCount = 0
    For I = 1 To UBound(VBProjectList, 1)
        Set TmpClassVBProject = VBProjectList(I)
        For J = 1 To TmpClassVBProject.Modules.Count
            Set TmpClassModule = TmpClassVBProject.Modules(J)
            ProcedureCount = ProcedureCount + TmpClassModule.Procedures.Count
        Next J
    Next
    
    Dim Output
    ReDim Output(1 To ProcedureCount, 1 To 6)
    '1:VBProject��
    '2:Module��
    '3:Procedure��
    '4:VBProject�̔ԍ�
    '5:Module�̔ԍ�
    '6:Procedure�̔ԍ�
    
    K = 0
    For I = 1 To UBound(VBProjectList, 1)
        Set TmpClassVBProject = VBProjectList(I)
        For J = 1 To TmpClassVBProject.Modules.Count
            Set TmpClassModule = TmpClassVBProject.Modules(J)
            For II = 1 To TmpClassModule.Procedures.Count
                K = K + 1
                Set TmpClassProcedure = TmpClassModule.Procedures(II)
                Output(K, 1) = TmpClassVBProject.Name
                Output(K, 2) = TmpClassModule.Name
                Output(K, 3) = TmpClassProcedure.Name
                Output(K, 4) = I
                Output(K, 5) = J
                Output(K, 6) = II
            Next II
        Next J
    Next
    
    �S�v���V�[�W���ꗗ�쐬 = Output
    
End Function
Sub �v���V�[�W�����̎g�p�v���V�[�W���擾(VBProjectList() As classVBProject, AllProcedureList)
    
    Dim I&, J&, II&, JJ&, K&, M&, N& '�����グ�p(Long�^)
    N = UBound(AllProcedureList, 1)
    '�v���V�[�W���̌����v�Z
    Dim TmpClassVBProject As classVBProject
    Dim TmpClassModule As classModule
    Dim TmpClassProcedure As ClassProcedure
    Dim TmpVBProjectNum%, TmpModuleNum%, TmpProcedureNum%
    Dim TmpKensakuCode As Object
    Dim TmpVBProjectName$, TmpModuleName$, TmpProcedureName$
    Dim TmpSiyosakiList As Object
    Dim TmpSiyoProcedure As ClassProcedure
    
    For I = 1 To UBound(VBProjectList, 1)
        Set TmpClassVBProject = VBProjectList(I)
        For J = 1 To TmpClassVBProject.Modules.Count
            Set TmpClassModule = TmpClassVBProject.Modules(J)
            For II = 1 To TmpClassModule.Procedures.Count
                Set TmpClassProcedure = TmpClassModule.Procedures(II)
                Set TmpKensakuCode = TmpClassProcedure.KensakuCode
                K = 0
                For JJ = 1 To N
                    TmpVBProjectName = AllProcedureList(JJ, 1)
                    TmpModuleName = AllProcedureList(JJ, 2)
                    TmpProcedureName = AllProcedureList(JJ, 3)
                    
                    If TmpProcedureName <> TmpClassProcedure.Name Then '�������g�̃v���V�[�W���͌�������Ȃ�
                        TmpVBProjectName = StrConv(TmpVBProjectName, vbUpperCase) '�����p�ɑ啶���ɕϊ�
                        TmpModuleName = StrConv(TmpModuleName, vbUpperCase) '�����p�ɑ啶���ɕϊ�
                        TmpProcedureName = StrConv(TmpProcedureName, vbUpperCase) '�����p�ɑ啶���ɕϊ�
                        
                        If TmpKensakuCode.Exists(TmpVBProjectName & "." & TmpModuleName & "." & TmpProcedureName) Or _
                           TmpKensakuCode.Exists(TmpModuleName & "." & TmpProcedureName) Or _
                           TmpKensakuCode.Exists(TmpProcedureName) Then

                            TmpVBProjectNum = AllProcedureList(JJ, 4)
                            TmpModuleNum = AllProcedureList(JJ, 5)
                            TmpProcedureNum = AllProcedureList(JJ, 6)
                            Set TmpSiyoProcedure = VBProjectList(TmpVBProjectNum).Modules(TmpModuleNum).Procedures(TmpProcedureNum)
                            TmpClassProcedure.AddUseProcedure TmpSiyoProcedure
                            
                        End If
                    End If
                Next JJ
            Next II
        Next J
    Next

End Sub
Function �O���Q�ƃv���V�[�W�����X�g�쐬(VBProjectList() As classVBProject)
    
    Dim I&, J&, II&, K&, M&, N& '�����グ�p(Long�^)
    '�v���V�[�W���̌����v�Z
    Dim TmpClassVBProject As classVBProject
    Dim TmpClassModule As classModule
    Dim TmpClassProcedure As ClassProcedure
    
    Dim TmpVBProjectName$, TmpModuleName$, TmpProcedureName$
    Dim TmpCode$
    
    Dim TmpVBProject
    
    Dim TmpExtProcedureDict As Object
    N = UBound(VBProjectList, 1)
    ReDim Output(1 To N)
    For I = 1 To N
        Set TmpExtProcedureDict = CreateObject("Scripting.Dictionary")
        TmpVBProjectName = VBProjectList(I).Name
        Set TmpClassVBProject = VBProjectList(I)
        For J = 1 To TmpClassVBProject.Modules.Count
            Set TmpClassModule = TmpClassVBProject.Modules(J)
            For II = 1 To TmpClassModule.Procedures.Count
                Set TmpClassProcedure = TmpClassModule.Procedures(II)
                Call �v���V�[�W�����̊O���Q�ƃv���V�[�W���擾(TmpVBProjectName, TmpClassProcedure, TmpExtProcedureDict)
            Next II
        Next J
        
        Set Output(I) = TmpExtProcedureDict
        
    Next I
        
    �O���Q�ƃv���V�[�W�����X�g�쐬 = Output
    
End Function
Sub �v���V�[�W�����̊O���Q�ƃv���V�[�W���擾(MyVBProjectName$, ClassProcedure As ClassProcedure, ExtProcedureDict As Object)
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim TmpUseProcedure As ClassProcedure
    If ClassProcedure.UseProcedure.Count = 0 Then
        '�g�p���Ă���v���V�[�W�������̏ꍇ�������Ȃ�
    Else
        For I = 1 To ClassProcedure.UseProcedure.Count
            Set TmpUseProcedure = ClassProcedure.UseProcedure(I)
            Call �v���V�[�W�����̊O���Q�ƃv���V�[�W���擾(MyVBProjectName, TmpUseProcedure, ExtProcedureDict)
            
            If TmpUseProcedure.VBProjectName <> MyVBProjectName Then 'VBProject�����قȂ�ΊO���Q��
                If ExtProcedureDict.Exists(TmpUseProcedure.Name) = False Then
                    ExtProcedureDict.Add TmpUseProcedure.Name, TmpUseProcedure.Code
                End If
            End If
        Next I
    End If

End Sub
