'@Folder modWorkbookExport
Option Explicit

Public Sub ExportWorkbookRepository(ByVal wb As Workbook, ByVal baseFolder As String)

    Dim workbookFolder As String
    
    workbookFolder = baseFolder & "\Workbook"
    
    Call CreateFolderIfNotExists(workbookFolder)
    
    Call ExportWorkbookInfo(wb, workbookFolder)
    Call ExportSheetsCsv(wb, workbookFolder)
    Call ExportTablesCsv(wb, workbookFolder)
    Call ExportNamesCsv(wb, workbookFolder)

End Sub

Private Sub ExportWorkbookInfo(ByVal wb As Workbook, ByVal folderPath As String)

    Dim textData As String
    Dim filePath As String
    
    filePath = folderPath & "\WorkbookInfo.txt"
    
    textData = ""
    textData = textData & "WorkbookName=" & wb.Name & vbCrLf
    textData = textData & "FullName=" & wb.FullName & vbCrLf
    textData = textData & "Path=" & wb.Path & vbCrLf
    textData = textData & "SheetCount=" & wb.Worksheets.Count & vbCrLf
    textData = textData & "TableCount=" & GetWorkbookTableCount(wb) & vbCrLf
    textData = textData & "NameCount=" & wb.Names.Count & vbCrLf
    textData = textData & "ExportDate=" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    
    Call SaveUtf8NoBom(filePath, textData)

End Sub

Private Sub ExportSheetsCsv(ByVal wb As Workbook, ByVal folderPath As String)

    Dim ws As Worksheet
    Dim textData As String
    Dim filePath As String
    
    filePath = folderPath & "\Sheets.csv"
    
    textData = CsvLine(Array( _
                    "SheetName", _
                    "Visible", _
                    "UsedRows", _
                    "UsedCols", _
                    "TableCount", _
                    "FormulaCount", _
                    "CodeName", _
                    "ExportTarget"))
    
    For Each ws In wb.Worksheets
        textData = textData & CsvLine(Array( _
                        ws.Name, _
                        GetSheetVisibleText(ws.Visible), _
                        GetUsedRows(ws), _
                        GetUsedCols(ws), _
                        ws.ListObjects.Count, _
                        GetFormulaCount(ws), _
                        ws.CodeName, _
                        "FALSE"))
    Next ws
    
    Call SaveUtf8NoBom(filePath, textData)

End Sub

Private Sub ExportTablesCsv(ByVal wb As Workbook, ByVal folderPath As String)

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim textData As String
    Dim filePath As String
    
    filePath = folderPath & "\Tables.csv"
    
    textData = CsvLine(Array( _
                    "TableName", _
                    "SheetName", _
                    "Address", _
                    "RowCount", _
                    "ColCount", _
                    "FieldNames"))
    
    For Each ws In wb.Worksheets
        For Each tbl In ws.ListObjects
            textData = textData & CsvLine(Array( _
                            tbl.Name, _
                            ws.Name, _
                            tbl.Range.Address(False, False), _
                            GetTableRowCount(tbl), _
                            tbl.Range.Columns.Count, _
                            GetTableFieldNames(tbl)))
        Next tbl
    Next ws
    
    Call SaveUtf8NoBom(filePath, textData)

End Sub

'============================================================
' テーブルのフィールド名一覧取得
'
' ListObjectの列名を " | " 区切りで連結して返す。
'
' 例：
'   製品コード | 製品名 | 数量 | 単価
'============================================================
Private Function GetTableFieldNames(ByVal tbl As ListObject) As String

    Dim columnDefinition As ListColumn
    Dim fieldNames As String
    
    For Each columnDefinition In tbl.ListColumns
    
        If fieldNames <> vbNullString Then
            fieldNames = fieldNames & " | "
        End If
        
        fieldNames = fieldNames & columnDefinition.Name
        
    Next columnDefinition
    
    GetTableFieldNames = fieldNames

End Function

Private Sub ExportNamesCsv(ByVal wb As Workbook, ByVal folderPath As String)

    Dim nm As Name
    Dim textData As String
    Dim filePath As String
    
    filePath = folderPath & "\Names.csv"
    
    textData = CsvLine(Array( _
                    "Name", _
                    "RefersTo", _
                    "Visible", _
                    "Value"))
    
    For Each nm In wb.Names
    
        If nm.Visible = True _
           And Left$(nm.Name, 6) <> "_xlfn." Then
    
            textData = textData & CsvLine(Array( _
                            nm.Name, _
                            nm.RefersTo, _
                            CStr(nm.Visible), _
                            GetNameValueSafe(nm)))
        End If
    
    Next nm

    Call SaveUtf8NoBom(filePath, textData)

End Sub

Private Function GetWorkbookTableCount(ByVal wb As Workbook) As Long

    Dim ws As Worksheet
    Dim countValue As Long
    
    For Each ws In wb.Worksheets
        countValue = countValue + ws.ListObjects.Count
    Next ws
    
    GetWorkbookTableCount = countValue

End Function

Private Function GetSheetVisibleText(ByVal visibleValue As Long) As String

    Select Case visibleValue
        Case xlSheetVisible
            GetSheetVisibleText = "Visible"
        Case xlSheetHidden
            GetSheetVisibleText = "Hidden"
        Case xlSheetVeryHidden
            GetSheetVisibleText = "VeryHidden"
        Case Else
            GetSheetVisibleText = CStr(visibleValue)
    End Select

End Function

Private Function GetUsedRows(ByVal ws As Worksheet) As Long

    If WorksheetFunction.CountA(ws.Cells) = 0 Then
        GetUsedRows = 0
    Else
        GetUsedRows = ws.UsedRange.Rows.Count
    End If

End Function

Private Function GetUsedCols(ByVal ws As Worksheet) As Long

    If WorksheetFunction.CountA(ws.Cells) = 0 Then
        GetUsedCols = 0
    Else
        GetUsedCols = ws.UsedRange.Columns.Count
    End If

End Function

Private Function GetFormulaCount(ByVal ws As Worksheet) As Long

    On Error GoTo Err_Handler
    
    Dim rng As Range
    
    Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    GetFormulaCount = rng.CountLarge
    
    Exit Function

Err_Handler:
    GetFormulaCount = 0

End Function

Private Function GetTableRowCount(ByVal tbl As ListObject) As Long

    If tbl.DataBodyRange Is Nothing Then
        GetTableRowCount = 0
    Else
        GetTableRowCount = tbl.DataBodyRange.Rows.Count
    End If

End Function

Private Function GetNameValueSafe(ByVal nm As Name) As String

    On Error GoTo Err_Handler
    
    GetNameValueSafe = CStr(nm.RefersToRange.Value)
    Exit Function

Err_Handler:
    GetNameValueSafe = vbNullString

End Function

Private Function CsvLine(ByVal values As Variant) As String

    Dim i As Long
    Dim lineText As String
    
    For i = LBound(values) To UBound(values)
        If i > LBound(values) Then
            lineText = lineText & ","
        End If
        
        lineText = lineText & CsvValue(CStr(values(i)))
    Next i
    
    CsvLine = lineText & vbCrLf

End Function

Private Function CsvValue(ByVal valueText As String) As String

    valueText = Replace(valueText, """", """""")
    CsvValue = """" & valueText & """"

End Function

Private Sub CreateFolderIfNotExists(ByVal folderPath As String)

    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If

End Sub


