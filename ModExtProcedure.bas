Attribute VB_Name = "ModExtProcedure"
Option Explicit
'外部参照プロシージャの取得用モジュール
'frmExtRefと連携している

Function フォーム用VBProject作成()
    
    Dim I%, J%, II%, K%, M%, N% '数え上げ用(Integer型)
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
            
            TmpProcedureNameList = モジュールのプロシージャ名一覧取得(TmpModule)
            Set TmpCodeDict = モジュールのコード一覧取得(TmpModule)
            If IsEmpty(TmpProcedureNameList) = False Then
                For II = 1 To UBound(TmpProcedureNameList)
                    Set TmpClassProcedure = New ClassProcedure
                    TmpProcedureName = TmpProcedureNameList(II)
                    TmpClassProcedure.Name = TmpProcedureName
                    TmpClassProcedure.Code = TmpCodeDict(TmpProcedureName)
                    Set TmpClassProcedure.KensakuCode = コードを検索用に変更(TmpCodeDict(TmpProcedureName))
                    TmpClassProcedure.VBProjectName = TmpClassVBProject.Name
                    TmpClassProcedure.ModuleName = TmpClassModule.Name
                    TmpClassModule.AddProcedure TmpClassProcedure
                Next II
            End If
            
            TmpClassVBProject.AddModule TmpClassModule
            
        Next J
        
        Set OutputVBProjectList(I) = TmpClassVBProject
        
    Next I
    
    フォーム用VBProject作成 = OutputVBProjectList
    
End Function
Function モジュールのプロシージャ名一覧取得(InputModule As VBComponent)
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
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
    
    If K = 0 Then 'モジュール内にプロシージャがない場合
        Output = Empty
    End If
    
    モジュールのプロシージャ名一覧取得 = Output
        
End Function
Function モジュールのコード一覧取得(InputModule As VBComponent)
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim ProcedureList
    ProcedureList = モジュールのプロシージャ名一覧取得(InputModule)
    Dim Output As Object
    Dim TmpProcedureName$, TmpStart&, TmpEnd&, TmpCode$
    If IsEmpty(ProcedureList) Then
        'プロシージャ無し
        Set Output = Nothing
    Else
        'プロシージャ有り
        N = UBound(ProcedureList, 1)
        Set Output = CreateObject("Scripting.Dictionary")
        For I = 1 To N
            TmpProcedureName = ProcedureList(I)
            With InputModule.CodeModule
                On Error Resume Next
                TmpStart = 0
                
                TmpStart = .ProcBodyLine(TmpProcedureName, 0)
                TmpEnd = .ProcCountLines(TmpProcedureName, 0)
                      
                If TmpStart = 0 Then 'クラスモジュールのコード取得用
                    TmpStart = .ProcBodyLine(TmpProcedureName, vbext_pk_Get)
                    TmpEnd = .ProcCountLines(TmpProcedureName, vbext_pk_Let)
                End If
                
                On Error GoTo 0
                
                TmpCode = .Lines(TmpStart, TmpEnd)
            End With
            
            Output.Add TmpProcedureName, TmpCode
        Next I
    End If
    
    Set モジュールのコード一覧取得 = Output

End Function
Function コードを検索用に変更(InputCode) As Object
    
    Dim CodeList, TmpStr$
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
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
        TmpStr = Trim(TmpStr) '左右の空白除去
        TmpStr = StrConv(TmpStr, vbUpperCase) '小文字に変換
'        TmpStr = StrConv(TmpStr, vbNarrow) '半角に変換
        If InStr(1, TmpStr, "'") > 0 Then
            TmpStr = Split(TmpStr, "'")(0) 'コメントの除去
        End If
        TmpStr = Replace(TmpStr, Chr(13), "") '改行を消去
        
        
        If TmpStr <> "" Then
            '指定文字で分割する
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
    Set コードを検索用に変更 = Output

End Function
Function 全プロシージャ一覧作成(VBProjectList)
    
    Dim I&, J&, II&, K&, M&, N& '数え上げ用(Long型)
    Dim ProcedureCount&
    'プロシージャの個数を計算
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
    '1:VBProject名
    '2:Module名
    '3:Procedure名
    '4:VBProjectの番号
    '5:Moduleの番号
    '6:Procedureの番号
    
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
    
    全プロシージャ一覧作成 = Output
    
End Function
Sub プロシージャ内の使用プロシージャ取得(VBProjectList() As classVBProject, AllProcedureList)
    
    Dim I&, J&, II&, JJ&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(AllProcedureList, 1)
    'プロシージャの個数を計算
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
                    
                    If TmpProcedureName <> TmpClassProcedure.Name Then '自分自身のプロシージャは検索から省く
                        TmpVBProjectName = StrConv(TmpVBProjectName, vbUpperCase) '検索用に大文字に変換
                        TmpModuleName = StrConv(TmpModuleName, vbUpperCase) '検索用に大文字に変換
                        TmpProcedureName = StrConv(TmpProcedureName, vbUpperCase) '検索用に大文字に変換
                        
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
Function 外部参照プロシージャリスト作成(VBProjectList() As classVBProject)
    
    Dim I&, J&, II&, K&, M&, N& '数え上げ用(Long型)
    'プロシージャの個数を計算
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
                Call プロシージャ内の外部参照プロシージャ取得(TmpVBProjectName, TmpClassProcedure, TmpExtProcedureDict)
            Next II
        Next J
        
        Set Output(I) = TmpExtProcedureDict
        
    Next I
        
    外部参照プロシージャリスト作成 = Output
    
End Function
Sub プロシージャ内の外部参照プロシージャ取得(MyVBProjectName$, ClassProcedure As ClassProcedure, ExtProcedureDict As Object)
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim TmpUseProcedure As ClassProcedure
    If ClassProcedure.UseProcedure.Count = 0 Then
        '使用しているプロシージャ無しの場合何もしない
    Else
        For I = 1 To ClassProcedure.UseProcedure.Count
            Set TmpUseProcedure = ClassProcedure.UseProcedure(I)
            Call プロシージャ内の外部参照プロシージャ取得(MyVBProjectName, TmpUseProcedure, ExtProcedureDict)
            
            If TmpUseProcedure.VBProjectName <> MyVBProjectName Then 'VBProject名が異なれば外部参照
                If ExtProcedureDict.Exists(TmpUseProcedure.Name) = False Then
                    ExtProcedureDict.Add TmpUseProcedure.Name, TmpUseProcedure.Code
                End If
            End If
        Next I
    End If

End Sub
