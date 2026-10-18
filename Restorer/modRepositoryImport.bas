'@Folder modRepositoryImport
Option Explicit

' ==========================================================
'  Excel VBA プロジェクト インポートツール
'  対象：アクティブなブックの VBProject
'
'  取込：
'   - BAS → 標準モジュール置換
'   - CLS → クラスモジュール置換
'   - CSV → UserForm レイアウト反映
'   - CLS → UserForm コード反映
'   - CLS → ThisWorkbook / Sheet モジュール反映
' ==========================================================

Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_Document As Long = 100

Public Sub ImportRepositoryForActiveWorkbook()

    Dim wb As Workbook
    Dim vbproj As Object
    Dim vbComp As Object
    
    Dim baseFolder As String
    Dim formsCsvFolder As String
    Dim formCodeFolder As String
    Dim modulesFolder As String
    Dim classesFolder As String
    Dim documentsFolder As String
    
    Set wb = ActiveWorkbook
    
    If wb Is Nothing Then
        MsgBox "アクティブなブックがありません。", vbExclamation
        Exit Sub
    End If
    
    If wb.Path = vbNullString Then
        MsgBox "ブックがまだ保存されていません。先に保存してください。", vbExclamation
        Exit Sub
    End If
    
    Set vbproj = wb.VBProject
    
    Dim repositoryRoot As String
    Dim vbaFolder As String
    Dim formsFolder As String
    
    repositoryRoot = wb.Path & "\ExcelRepository"
    
    baseFolder = repositoryRoot & "\" & _
                 GetNameWithoutExtension(wb.Name) & "_repository"
    
    vbaFolder = baseFolder & "\VBA"
    
    formsFolder = vbaFolder & "\Forms"
    
    formsCsvFolder = formsFolder & "\ControlsCsv"
    formCodeFolder = formsFolder & "\Code"
    
    modulesFolder = vbaFolder & "\Modules"
    classesFolder = vbaFolder & "\Classes"
    documentsFolder = vbaFolder & "\Documents"

    ImportStdModulesFromBas vbproj, modulesFolder
    ImportClassModulesFromCls vbproj, classesFolder
    
    For Each vbComp In vbproj.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_MSForm
                ImportFormLayoutFromCsv vbComp, formsCsvFolder
                ImportFormCodeFromCls vbComp, formCodeFolder
                
            Case vbext_ct_Document
                ImportDocumentModuleFromCls vbComp, documentsFolder
        End Select
    Next vbComp
    
    MsgBox "インポート完了：" & vbCrLf & baseFolder, vbInformation

End Sub

' ==========================================================
'  標準モジュール / クラスモジュール
' ==========================================================
Private Sub ImportStdModulesFromBas(ByVal vbproj As Object, ByVal folderPath As String)

    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim vbComp As Object
    Dim moduleName As String
    Dim codeText As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Exit Sub
    
    Set folder = fso.GetFolder(folderPath)
    
    For Each file In folder.Files
        If LCase$(fso.GetExtensionName(file.Name)) = "bas" Then
            
            moduleName = fso.GetBaseName(file.Name)
            
            On Error Resume Next
            Set vbComp = vbproj.VBComponents(moduleName)
            On Error GoTo 0
            
            If Not vbComp Is Nothing Then
                vbproj.VBComponents.Remove vbComp
            End If
            
            Set vbComp = vbproj.VBComponents.Add(vbext_ct_StdModule)
            vbComp.Name = moduleName
            
            codeText = ReadUtf8Text(file.Path)
            codeText = RemoveVbeExportHeader(codeText)
            
            If Len(codeText) > 0 Then
                vbComp.CodeModule.InsertLines 1, codeText
            End If
            
            Set vbComp = Nothing
        End If
    Next file

End Sub

Private Sub ImportClassModulesFromCls(ByVal vbproj As Object, ByVal folderPath As String)

    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim vbComp As Object
    Dim className As String
    Dim codeText As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Exit Sub
    
    Set folder = fso.GetFolder(folderPath)
    
    For Each file In folder.Files
        If LCase$(fso.GetExtensionName(file.Name)) = "cls" Then
            
            className = fso.GetBaseName(file.Name)
            
            On Error Resume Next
            Set vbComp = vbproj.VBComponents(className)
            On Error GoTo 0
            
            If Not vbComp Is Nothing Then
                vbproj.VBComponents.Remove vbComp
            End If
            
            Set vbComp = vbproj.VBComponents.Add(vbext_ct_ClassModule)
            vbComp.Name = className
            
            codeText = ReadUtf8Text(file.Path)
            codeText = RemoveVbeExportHeader(codeText)
            
            If Len(codeText) > 0 Then
                vbComp.CodeModule.InsertLines 1, codeText
            End If
            
            Set vbComp = Nothing
        End If
    Next file

End Sub

' ==========================================================
'  UserForm レイアウト
' ==========================================================
Private Sub ImportFormLayoutFromCsv(ByVal vbComp As Object, ByVal folderPath As String)

    Dim designer As Object
    Dim ctrl As Object
    Dim filePath As String
    Dim allText As String
    Dim lines() As String
    Dim parts() As String
    
    Dim i As Long
    Dim ctrlName As String
    Dim ctrlType As String
    Dim caption As String
    Dim leftVal As Double
    Dim topVal As Double
    Dim widthVal As Double
    Dim heightVal As Double
    Dim progId As String
    
    Dim fontName As String
    Dim fontSize As Double
    Dim fontBold As Boolean
    Dim foreColor As Variant
    Dim backColor As Variant

On Error Resume Next
    Set designer = vbComp.designer
    On Error GoTo 0
    
    If designer Is Nothing Then Exit Sub
    
    filePath = folderPath & "\" & vbComp.Name & ".csv"
    If Dir$(filePath) = vbNullString Then Exit Sub
    
    allText = ReadUtf8Text(filePath)
    If Len(allText) = 0 Then Exit Sub
    
    allText = Replace(allText, vbCrLf, vbLf)
    allText = Replace(allText, vbCr, vbLf)
    
    lines = Split(allText, vbLf)
    
    For i = 1 To UBound(lines)
        If Trim$(lines(i)) <> vbNullString Then
            parts = ParseCsvLine(lines(i))
            
            If UBound(parts) >= 6 Then
                ctrlName = parts(0)
                ctrlType = parts(1)
                caption = parts(2)
                leftVal = CDbl(parts(3))
                topVal = CDbl(parts(4))
                widthVal = CDbl(parts(5))
                heightVal = CDbl(parts(6))
            
                If UBound(parts) >= 7 Then fontName = parts(7) Else fontName = vbNullString
                If UBound(parts) >= 8 And parts(8) <> vbNullString Then fontSize = CDbl(parts(8)) Else fontSize = 0
                If UBound(parts) >= 9 And parts(9) <> vbNullString Then fontBold = CBool(parts(9)) Else fontBold = False
                If UBound(parts) >= 10 Then foreColor = parts(10) Else foreColor = vbNullString
                If UBound(parts) >= 11 Then backColor = parts(11) Else backColor = vbNullString
                Set ctrl = Nothing
                
                On Error Resume Next
                Set ctrl = designer.Controls(ctrlName)
                On Error GoTo 0
                
                If ctrl Is Nothing Then
                    progId = GetProgIdFromCtrlType(ctrlType)
                    If progId <> vbNullString Then
                        On Error Resume Next
                        Set ctrl = designer.Controls.Add(progId, ctrlName, True)
                        On Error GoTo 0
                    End If
                End If
                
                If Not ctrl Is Nothing Then
                    If ctrlType = vbNullString Or TypeName(ctrl) = ctrlType Then
                        Call ApplyControlLayout(ctrl, parts)
                    End If
                End If
            End If
        End If
    Next i

End Sub

Private Function GetProgIdFromCtrlType(ByVal ctrlType As String) As String

    Select Case LCase$(ctrlType)
        Case "textbox"
            GetProgIdFromCtrlType = "Forms.TextBox.1"
        Case "label"
            GetProgIdFromCtrlType = "Forms.Label.1"
        Case "commandbutton"
            GetProgIdFromCtrlType = "Forms.CommandButton.1"
        Case "combobox"
            GetProgIdFromCtrlType = "Forms.ComboBox.1"
        Case "listbox"
            GetProgIdFromCtrlType = "Forms.ListBox.1"
        Case "checkbox"
            GetProgIdFromCtrlType = "Forms.CheckBox.1"
        Case "optionbutton"
            GetProgIdFromCtrlType = "Forms.OptionButton.1"
        Case "frame"
            GetProgIdFromCtrlType = "Forms.Frame.1"
        Case "multipage"
            GetProgIdFromCtrlType = "Forms.MultiPage.1"
        Case "tabstrip"
            GetProgIdFromCtrlType = "Forms.TabStrip.1"
        Case "scrollbar"
            GetProgIdFromCtrlType = "Forms.ScrollBar.1"
        Case "spinbutton"
            GetProgIdFromCtrlType = "Forms.SpinButton.1"
        Case "image"
            GetProgIdFromCtrlType = "Forms.Image.1"
        Case Else
            GetProgIdFromCtrlType = vbNullString
    End Select

End Function

Private Sub ApplyControlLayout(ByVal ctrl As Object, ByRef parts() As String)

    On Error Resume Next

    ctrl.Left = CDbl(parts(3))
    ctrl.Top = CDbl(parts(4))
    ctrl.Width = CDbl(parts(5))
    ctrl.Height = CDbl(parts(6))

    ctrl.caption = parts(2)

    If UBound(parts) >= 7 Then
        If parts(7) <> vbNullString Then ctrl.Font.Name = parts(7)
    End If

    If UBound(parts) >= 8 Then
        If parts(8) <> vbNullString Then ctrl.Font.Size = CDbl(parts(8))
    End If

    If UBound(parts) >= 9 Then
        If parts(9) <> vbNullString Then ctrl.Font.Bold = CBool(parts(9))
    End If

    If UBound(parts) >= 10 Then
        If parts(10) <> vbNullString Then ctrl.foreColor = CLng(parts(10))
    End If

    If UBound(parts) >= 11 Then
        If parts(11) <> vbNullString Then ctrl.backColor = CLng(parts(11))
    End If

    On Error GoTo 0

End Sub

' ==========================================================
'  UserForm コード / Document モジュール
' ==========================================================

Private Sub ImportFormCodeFromCls(ByVal vbComp As Object, ByVal folderPath As String)

    Dim filePath As String
    Dim codeText As String
    Dim cm As Object
    Dim lineCount As Long
    
    filePath = folderPath & "\" & vbComp.Name & ".cls"
    If Dir$(filePath) = vbNullString Then Exit Sub
    
    codeText = ReadUtf8Text(filePath)
    codeText = RemoveVbeExportHeader(codeText)

    Set cm = vbComp.CodeModule
    
    lineCount = cm.CountOfLines
    If lineCount > 0 Then
        cm.DeleteLines 1, lineCount
    End If
    
    If Len(codeText) > 0 Then
        cm.InsertLines 1, codeText
    End If

End Sub

Private Sub ImportDocumentModuleFromCls(ByVal vbComp As Object, ByVal folderPath As String)

    Dim filePath As String
    Dim codeText As String
    Dim cm As Object
    Dim lineCount As Long
    
    filePath = folderPath & "\" & vbComp.Name & ".cls"
    If Dir$(filePath) = vbNullString Then Exit Sub
    
    codeText = ReadUtf8Text(filePath)
    codeText = RemoveVbeExportHeader(codeText)

    Set cm = vbComp.CodeModule
    
    lineCount = cm.CountOfLines
    If lineCount > 0 Then
        cm.DeleteLines 1, lineCount
    End If
    
    If Len(codeText) > 0 Then
        cm.InsertLines 1, codeText
    End If

End Sub

' ==========================================================
'  ヘルパー
' ==========================================================

Private Function GetNameWithoutExtension(ByVal fileName As String) As String

    Dim pos As Long
    
    pos = InStrRev(fileName, ".")
    
    If pos > 0 Then
        GetNameWithoutExtension = Left$(fileName, pos - 1)
    Else
        GetNameWithoutExtension = fileName
    End If

End Function

Private Function ParseCsvLine(ByVal line As String) As String()

    Dim result() As String
    Dim i As Long
    Dim n As Long
    Dim ch As String
    Dim current As String
    Dim inQuote As Boolean
    
    ReDim result(0)
    
    n = 0
    current = vbNullString
    inQuote = False
    i = 1
    
    Do While i <= Len(line)
        ch = Mid$(line, i, 1)
        
        If ch = """" Then
            If inQuote Then
                If i < Len(line) And Mid$(line, i + 1, 1) = """" Then
                    current = current & """"
                    i = i + 1
                Else
                    inQuote = False
                End If
            Else
                inQuote = True
            End If
        
        ElseIf ch = "," Then
            If inQuote Then
                current = current & ch
            Else
                result(n) = current
                n = n + 1
                ReDim Preserve result(n)
                current = vbNullString
            End If
        
        Else
            current = current & ch
        End If
        
        i = i + 1
    Loop
    
    result(n) = current
    ParseCsvLine = result

End Function


