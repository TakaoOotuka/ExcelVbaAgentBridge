'@Folder modRepositoryExport
Option Explicit

' ==========================================================
'  Excel VBA プロジェクト エクスポートツール
'  対象：アクティブなブックの VBProject
'
'  出力：
'   - UserForm          → FormsCsv / FormCode
'   - 標準モジュール     → Modules
'   - クラスモジュール   → Classes
'   - ThisWorkbook/Sheet → Documents
' ==========================================================

Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_Document As Long = 100

Public Sub ExportRepositoryForActiveWorkbook()

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
    baseFolder = repositoryRoot & "\" & GetNameWithoutExtension(wb.Name) & "_repository"
    vbaFolder = baseFolder & "\VBA"
    
    formsFolder = vbaFolder & "\Forms"
    formCodeFolder = formsFolder & "\Code"
    formsCsvFolder = formsFolder & "\ControlsCsv"
    modulesFolder = vbaFolder & "\Modules"
    classesFolder = vbaFolder & "\Classes"
    documentsFolder = vbaFolder & "\Documents"
    
    CreateFolderIfNotExists repositoryRoot
    CreateFolderIfNotExists baseFolder
    CreateFolderIfNotExists vbaFolder
    CreateFolderIfNotExists formsFolder
    CreateFolderIfNotExists formsCsvFolder
    CreateFolderIfNotExists formCodeFolder
    CreateFolderIfNotExists modulesFolder
    CreateFolderIfNotExists classesFolder
    CreateFolderIfNotExists documentsFolder

    For Each vbComp In vbproj.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_MSForm
                ExportFormControlsToCsv vbComp, formsCsvFolder
                ExportFormCodeToCls vbComp, formCodeFolder
                
            Case vbext_ct_StdModule
                ExportStdModuleToBas vbComp, modulesFolder
                
            Case vbext_ct_ClassModule
                ExportClassModuleToCls vbComp, classesFolder
                
            Case vbext_ct_Document
                ExportDocumentModuleCodeOnly vbComp, documentsFolder
                
            Case Else
                ' 対象外
        End Select
    Next vbComp
    
    Call ExportWorkbookRepository(wb, baseFolder)
    
    MsgBox "エクスポート完了：" & vbCrLf & baseFolder, vbInformation

End Sub

' ==========================================================
'  UserForm
' ==========================================================

Private Sub ExportFormControlsToCsv(ByVal vbComp As Object, ByVal folderPath As String)

    Dim designer As Object
    Dim ctrl As Object
    Dim filePath As String
    Dim line As String
    Dim allText As String
    
    On Error Resume Next
    Set designer = vbComp.designer
    On Error GoTo 0
    
    If designer Is Nothing Then Exit Sub
    
    filePath = folderPath & "\" & vbComp.Name & ".csv"
    
    allText = "ControlName,ControlType,Caption,Left,Top,Width,Height," & _
              "FontName,FontSize,FontBold,ForeColor,BackColor" & vbCrLf

    For Each ctrl In designer.Controls
    
        line = EscapeCsv(ctrl.Name) & "," & _
               EscapeCsv(TypeName(ctrl)) & "," & _
               EscapeCsv(GetCtrlCaptionSafe(ctrl)) & "," & _
               CStr(ctrl.Left) & "," & _
               CStr(ctrl.Top) & "," & _
               CStr(ctrl.Width) & "," & _
               CStr(ctrl.Height) & "," & _
               EscapeCsv(GetCtrlPropertySafe(ctrl, "Font.Name")) & "," & _
               EscapeCsv(GetCtrlPropertySafe(ctrl, "Font.Size")) & "," & _
               EscapeCsv(GetCtrlPropertySafe(ctrl, "Font.Bold")) & "," & _
               EscapeCsv(GetCtrlPropertySafe(ctrl, "ForeColor")) & "," & _
               EscapeCsv(GetCtrlPropertySafe(ctrl, "BackColor"))
    
        allText = allText & line & vbCrLf
    
    Next ctrl

    Call SaveUtf8NoBom(filePath, allText)

End Sub

Private Sub ExportFormCodeToCls(ByVal vbComp As Object, ByVal folderPath As String)

    Dim cm As Object
    Dim codeText As String
    Dim filePath As String
    Dim lineCount As Long
    
    Set cm = vbComp.CodeModule
    lineCount = cm.CountOfLines
    If lineCount <= 0 Then Exit Sub
    
    codeText = cm.lines(1, lineCount)
    codeText = RemoveAttributeLines(codeText)
    
    filePath = folderPath & "\" & vbComp.Name & ".cls"
    
    Call SaveUtf8NoBom(filePath, codeText)

End Sub

' ==========================================================
'  標準モジュール / クラスモジュール / Documentモジュール
' ==========================================================

Private Sub ExportStdModuleToBas(ByVal vbComp As Object, ByVal folderPath As String)

    Dim filePath As String
    
    filePath = folderPath & "\" & vbComp.Name & ".bas"
    
    vbComp.Export filePath
    Call ConvertFileToUtf8WithoutAttribute(filePath)

End Sub

Private Sub ExportClassModuleToCls(ByVal vbComp As Object, ByVal folderPath As String)

    Dim filePath As String
    
    filePath = folderPath & "\" & vbComp.Name & ".cls"
    
    vbComp.Export filePath
    Call ConvertFileToUtf8WithoutAttribute(filePath)

End Sub

Private Sub ExportDocumentModuleCodeOnly(ByVal vbComp As Object, ByVal folderPath As String)

    Dim filePath As String
    Dim codeText As String
    Dim lineCount As Long
    Dim fso As Object
    Dim ts As Object
    
    filePath = folderPath & "\" & vbComp.Name & ".cls"
    
    With vbComp.CodeModule
        lineCount = .CountOfLines
        If lineCount > 0 Then
            codeText = .lines(1, lineCount)
            codeText = RemoveAttributeLines(codeText)
        End If
    End With
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(filePath, True, False)
    ts.Write codeText
    ts.Close
    
    Call ConvertFileToUtf8(filePath)

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

Private Sub CreateFolderIfNotExists(ByVal folderPath As String)

    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If

End Sub

Private Function GetCtrlCaptionSafe(ByVal ctrl As Object) As String

    On Error Resume Next
    GetCtrlCaptionSafe = ctrl.caption
    
    If Err.Number <> 0 Then
        GetCtrlCaptionSafe = vbNullString
        Err.Clear
    End If
    
    On Error GoTo 0

End Function

Private Function EscapeCsv(ByVal text As String) As String

    Dim t As String
    
    t = Replace(text, """", """""")
    EscapeCsv = """" & t & """"

End Function

Private Sub ConvertFileToUtf8WithoutAttribute(ByVal filePath As String)

    Dim textData As String
    
    If Dir$(filePath) = vbNullString Then Exit Sub
    
    textData = ReadTextFile(filePath)
    textData = RemoveAttributeLines(textData)
    
    Call SaveUtf8NoBom(filePath, textData)

End Sub

Private Function GetCtrlPropertySafe(ByVal ctrl As Object, ByVal propertyName As String) As String

    On Error GoTo Err_Handler

    Select Case propertyName

        Case "Font.Name"
            GetCtrlPropertySafe = CStr(ctrl.Font.Name)

        Case "Font.Size"
            GetCtrlPropertySafe = CStr(ctrl.Font.Size)

        Case "Font.Bold"
            GetCtrlPropertySafe = CStr(ctrl.Font.Bold)

        Case Else
            GetCtrlPropertySafe = CStr(CallByName(ctrl, propertyName, VbGet))

    End Select

    Exit Function

Err_Handler:
    GetCtrlPropertySafe = vbNullString

End Function


