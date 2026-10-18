# Excel VBAプロジェクト リポジトリエクスポート仕様書
## Workbook情報出力仕様

### 1. 概要

VBAプロジェクトとは別に、Excelブック本体の構成情報をテキストおよびCSV形式で出力する。

メイン処理：

```vb
Public Sub ExportWorkbookRepository(ByVal wb As Workbook, ByVal baseFolder As String)
```

出力先は、

```text
対象ブック名_repository\Workbook
```

とする。

---

# 2. リポジトリ全体構造

最終的なリポジトリ構造は以下となる。

```text
Excelファイル保存フォルダ
│
├─ 対象ブック.xlsm
│
└─ ExcelRepository
    │
    └─ 対象ブック名_repository
        │
        ├─ VBA
        │   │
        │   ├─ Forms
        │   │   │
        │   │   ├─ Code
        │   │   │   ├─ UserForm1.cls
        │   │   │   └─ UserForm2.cls
        │   │   │
        │   │   └─ ControlsCsv
        │   │       ├─ UserForm1.csv
        │   │       └─ UserForm2.csv
        │   │
        │   ├─ Modules
        │   │   ├─ Module1.bas
        │   │   └─ modSample.bas
        │   │
        │   ├─ Classes
        │   │   └─ Class1.cls
        │   │
        │   └─ Documents
        │       ├─ ThisWorkbook.cls
        │       ├─ Sheet1.cls
        │       └─ Sheet2.cls
        │
        └─ Workbook
            ├─ WorkbookInfo.txt
            ├─ Sheets.csv
            ├─ Tables.csv
            └─ Names.csv
```

---

# 3. Workbookエクスポート処理

処理順序：

```text
ExportWorkbookRepository
        │
        ├─ Workbookフォルダ作成
        │
        ├─ ExportWorkbookInfo
        │
        ├─ ExportSheetsCsv
        │
        ├─ ExportTablesCsv
        │
        └─ ExportNamesCsv
```

---

# 4. WorkbookInfo.txt

出力ファイル：

```text
Workbook\WorkbookInfo.txt
```

ブック全体の基本情報を `キー=値` 形式で出力する。

出力項目：

| 項目 | 内容 |
|---|---|
| WorkbookName | Excelブック名 |
| FullName | フルパスを含むブック名 |
| Path | 保存フォルダ |
| SheetCount | ワークシート数 |
| TableCount | Excelテーブル総数 |
| NameCount | 定義名数 |
| ExportDate | エクスポート日時 |

出力例：

```text
WorkbookName=Sample.xlsm
FullName=C:\Work\Sample.xlsm
Path=C:\Work
SheetCount=5
TableCount=3
NameCount=8
ExportDate=2026-08-23 14:10:00
```

文字コードは `SaveUtf8NoBom` によりUTF-8 BOMなしで保存する。

---

# 5. Sheets.csv

出力ファイル：

```text
Workbook\Sheets.csv
```

対象ブックに存在するすべてのワークシートについて基本情報を出力する。

CSV項目：

| 列 | 内容 |
|---|---|
| SheetName | シート名 |
| Visible | シート表示状態 |
| UsedRows | 使用行数 |
| UsedCols | 使用列数 |
| TableCount | シート内Excelテーブル数 |
| FormulaCount | 数式セル数 |
| CodeName | VBA CodeName |
| ExportTarget | 将来的なシート内容出力対象フラグ |

CSVヘッダー：

```text
SheetName,Visible,UsedRows,UsedCols,TableCount,FormulaCount,CodeName,ExportTarget
```

---

## 5.1 Visible

Excelのシート表示状態を文字列へ変換する。

```text
xlSheetVisible
    → Visible

xlSheetHidden
    → Hidden

xlSheetVeryHidden
    → VeryHidden
```

未知の値の場合は数値を文字列として出力する。

---

## 5.2 UsedRows

シート内にデータが存在しない場合：

```text
0
```

データが存在する場合：

```vb
ws.UsedRange.Rows.Count
```

を出力する。

---

## 5.3 UsedCols

シート内にデータが存在しない場合：

```text
0
```

データが存在する場合：

```vb
ws.UsedRange.Columns.Count
```

を出力する。

---

## 5.4 FormulaCount

UsedRange内の数式セル数を取得する。

```vb
ws.UsedRange.SpecialCells(xlCellTypeFormulas)
```

数式セルが存在しない場合など、取得時にエラーとなった場合は、

```text
0
```

とする。

---

## 5.5 ExportTarget

現時点ではすべてのシートについて、

```text
FALSE
```

を固定出力する。

将来的にシート内容そのものをエクスポートする場合の対象指定として使用することを想定した項目である。

---

# 6. Tables.csv

出力ファイル：

```text
Workbook\Tables.csv
```

対象ブック内のすべてのExcelテーブル（ListObject）を出力する。

CSV項目：

| 列 | 内容 |
|---|---|
| TableName | Excelテーブル名 |
| SheetName | 所属シート名 |
| Address | テーブル範囲 |
| RowCount | データ行数 |
| ColCount | 列数 |

CSVヘッダー：

```text
TableName,SheetName,Address,RowCount,ColCount,FieldNames
```

すべてのワークシートを走査し、その中の `ListObjects` を出力する。

---

## 6.1 Address

Excelテーブルの範囲を、

```vb
tbl.Range.Address(False, False)
```

で取得する。

例：

```text
A1:H25
```

絶対参照記号 `$` は付けない。

---

## 6.2 RowCount

テーブルにデータ行が存在しない場合、

```vb
tbl.DataBodyRange Is Nothing
```

となるため、

```text
0
```

を出力する。

データが存在する場合は、

```vb
tbl.DataBodyRange.Rows.Count
```

を使用する。

ヘッダー行・集計行はデータ行数に含めない。

---

## 6.3 ColCount

```vb
tbl.Range.Columns.Count
```

を使用する。

---

# 7. Names.csv

出力ファイル：

```text
Workbook\Names.csv
```

対象ブックに登録されている定義名を出力する。

CSV項目：

| 列 | 内容 |
|---|---|
| Name | 定義名 |
| RefersTo | 参照式 |
| Visible | 表示状態 |
| Value | 参照先の値 |

CSVヘッダー：

```text
Name,RefersTo,Visible,Value
```

---

# 8. 定義名の出力条件

すべての `Workbook.Names` をそのまま出力するのではなく、以下の条件を満たしたものだけを対象とする。

```vb
nm.Visible = True
```

かつ、

```vb
Left$(nm.Name, 6) <> "_xlfn."
```

つまり、

- 非表示の定義名
- `_xlfn.` で始まるExcel内部的な定義名

は出力しない。

---

# 9. 定義名のValue取得

定義名がセル範囲を参照している場合、

```vb
nm.RefersToRange.Value
```

を取得する。

取得できない定義名の場合はエラーとせず、

```text
空文字
```

を出力する。

処理：

```vb
GetNameValueSafe
```

---

# 10. Excelテーブル総数

`WorkbookInfo.txt` の `TableCount` には、ブック内の全ワークシートを走査して、

```vb
ws.ListObjects.Count
```

を合計した値を使用する。

処理：

```vb
GetWorkbookTableCount
```

---

# 11. CSV生成

Workbook関連CSVは、

```vb
CsvLine
```

および、

```vb
CsvValue
```

を使用して生成する。

各値は必ずダブルクォートで囲む。

例：

```text
"Sheet1","Visible","100","12","1","25","Sheet1","FALSE"
```

値の中にダブルクォートが存在する場合は、

```text
"
```

を、

```text
""
```

へ変換する。

各CSV行の終端は `vbCrLf` とする。

最終保存時には `SaveUtf8NoBom` を使用する。

---

# 12. VBA側との役割分担

リポジトリは大きく、

```text
VBA
Workbook
```

の2つに分ける。

## VBA

VBAプログラムそのものを管理する。

```text
VBA
├─ Forms
├─ Modules
├─ Classes
└─ Documents
```

主な目的：
- コード差分確認
- UserFormコード管理
- UserFormレイアウト管理
- Gitによる変更履歴管理

## Workbook

Excelブックの構造情報を管理する。

```text
Workbook
├─ WorkbookInfo.txt
├─ Sheets.csv
├─ Tables.csv
└─ Names.csv
```

主な目的：
- シート構成確認
- テーブル構成確認
- 定義名確認
- VBA以外のExcel構造変更の把握
---

# 13. 全体処理フロー

```text
ExportRepositoryForActiveWorkbook
        │
        ├─ ActiveWorkbook取得
        │
        ├─ 出力フォルダ作成
        │
        ├─ VBProjectエクスポート
        │       │
        │       ├─ Forms
        │       │    ├─ ControlsCsv
        │       │    └─ Code
        │       │
        │       ├─ Modules
        │       ├─ Classes
        │       └─ Documents
        │
        ├─ ExportWorkbookRepository
        │       │
        │       ├─ WorkbookInfo.txt
        │       ├─ Sheets.csv
        │       ├─ Tables.csv
        │       └─ Names.csv
        │
        └─ 完了
```

---

# 14. 正規化ツールとの関係

エクスポートされたリポジトリは複数階層のフォルダ構造を持つ。

例：

```text
対象ブック名_repository
│
├─ VBA
│   ├─ Forms
│   │   ├─ Code
│   │   └─ ControlsCsv
│   ├─ Modules
│   ├─ Classes
│   └─ Documents
│
└─ Workbook
```

したがって、ファイル正規化ツールは固定フォルダを指定して処理してはならない。

仕様は以下とする。

```text
正規化スクリプトが置かれたフォルダ
        ↓
直下の全ファイル
        ↓
直下の全フォルダ
        ↓
さらにその配下のフォルダ
        ↓
以下すべて再帰的に処理
```

特定の、

```text
Source
Forms
Modules
Workbook
```

などのフォルダ名には依存しない。

ただし `.git` フォルダはリポジトリ管理情報そのものであるため処理対象外とする。

---

# 15. 現時点の出力対象

現在リポジトリ化している情報：

```text
Excelブック
│
├─ VBAコード
├─ UserFormコード
├─ UserFormコントロール配置
├─ Sheet構成
├─ Excelテーブル構成
├─ 定義名
└─ Workbook基本情報
```

現時点では、ワークシートのセルデータそのものはエクスポートしていない。

`Sheets.csv` の `ExportTarget` は現在すべて `FALSE` であり、シート内容エクスポート機能については将来拡張用とする。