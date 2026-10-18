'@Folder modUtf8File
Option Explicit

'==========================================================
' UTF-8 ファイル入出力ユーティリティ
'==========================================================

Public Sub SaveUtf8NoBom(ByVal filePath As String, ByVal textData As String)

    Dim stm As Object
    Dim bytes() As Byte
    
    Set stm = CreateObject("ADODB.Stream")
    
    With stm
        .Type = 2                  ' text
        .Charset = "utf-8"
        .Open
        .WriteText textData
        
        .Position = 0
        .Type = 1                  ' binary
        bytes = .Read
        .Close
    End With
    
    If UBound(bytes) >= 2 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
            Dim i As Long
            Dim newBytes() As Byte
            
            ReDim newBytes(UBound(bytes) - 3)
            
            For i = 3 To UBound(bytes)
                newBytes(i - 3) = bytes(i)
            Next i
            
            bytes = newBytes
        End If
    End If
    
    Set stm = CreateObject("ADODB.Stream")
    
    With stm
        .Type = 1                  ' binary
        .Open
        .Write bytes
        .SaveToFile filePath, 2    ' adSaveCreateOverWrite
        .Close
    End With

End Sub

Public Function ReadUtf8Text(ByVal filePath As String) As String

    Dim stm As Object
    
    Set stm = CreateObject("ADODB.Stream")
    
    With stm
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile filePath
        ReadUtf8Text = .ReadText
        .Close
    End With

End Function

Public Sub ConvertFileToUtf8(ByVal filePath As String)

    Dim fso As Object
    Dim ts As Object
    Dim textData As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    '--- ファイル存在チェック ---
    If Not fso.FileExists(filePath) Then
        Debug.Print "ファイル無し: " & filePath
        Exit Sub
    End If

    On Error GoTo ReadErr

    '--- 読み込み ---
    Set ts = fso.OpenTextFile(filePath, 1, False)
    
    If ts.AtEndOfStream Then
        textData = ""
    Else
        textData = ts.ReadAll
    End If
    
    ts.Close

    '--- UTF8保存 ---
    Call SaveUtf8NoBom(filePath, textData)

    Exit Sub

ReadErr:
    Debug.Print "読み込みエラー: " & filePath
    Debug.Print Err.Number & " : " & Err.Description

    On Error Resume Next
    If Not ts Is Nothing Then ts.Close

End Sub

Public Function RemoveAttributeLines(ByVal codeText As String) As String

    Dim lines() As String
    Dim result As String
    Dim i As Long
    Dim oneLine As String
    
    codeText = Replace(codeText, vbCrLf, vbLf)
    codeText = Replace(codeText, vbCr, vbLf)
    
    lines = Split(codeText, vbLf)
    
    For i = LBound(lines) To UBound(lines)
        oneLine = lines(i)
        
        If Left$(Trim$(oneLine), 10) <> "Attribute " Then
            result = result & oneLine & vbCrLf
        End If
    Next i
    
    RemoveAttributeLines = result

End Function

Public Function ReadTextFile(ByVal filePath As String) As String

    Dim fso As Object
    Dim ts As Object
    
    If Dir$(filePath) = vbNullString Then Exit Function
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set ts = fso.OpenTextFile(filePath, 1, False)
    
    If Not ts.AtEndOfStream Then
        ReadTextFile = ts.ReadAll
    End If
    
    ts.Close

End Function


