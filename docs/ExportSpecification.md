# Excel VBAプロジェクト リポジトリエクスポート仕様書

## 1. 概要

Excel VBAプロジェクトの内容を、Git等で差分管理しやすいテキスト形式へエクスポートする。

対象はアクティブなExcelブックの `VBProject` とし、VBAコンポーネントの種類ごとにフォルダを分けて保存する。

また、UserFormについてはコードだけでなく、フォーム上のコントロール配置情報をCSVとして出力する。

---

## 2. 実行入口

メイン処理：

```vb
Public Sub ExportRepositoryForActiveWorkbook()
```

実行時の `ActiveWorkbook` を対象とする。

---

## 3. 実行条件

以下の場合は処理を中止する。

### アクティブブックが存在しない

```text
アクティブなブックがありません。
```

### ブックが未保存

`Workbook.Path` が空の場合、以下を表示して終了する。

```text
ブックがまだ保存されていません。先に保存してください。
```

---

## 4. 出力先

対象Excelブックと同じフォルダに `ExcelRepository` フォルダを作成する。

基本構造：

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
        └─ その他Excel本体情報
            ※ ExportWorkbookRepository により出力
```

`対象ブック名_repository` のブック名部分には拡張子を含めない。

例：

```text
売上管理.xlsm

↓

ExcelRepository
└─ 売上管理_repository
```

---

# 5. VBAコンポーネント別処理

Excel VBAの `VBComponents` を順番に確認し、コンポーネント種別によって処理を分岐する。

使用するコンポーネントタイプ：

| 種類 | 定数値 | 出力先 |
|---|---:|---|
| 標準モジュール | 1 | VBA\Modules |
| クラスモジュール | 2 | VBA\Classes |
| UserForm | 3 | VBA\Forms |
| Documentモジュール | 100 | VBA\Documents |

その他のコンポーネントは対象外とする。

---

# 6. UserForm出力

UserFormは以下の2種類に分けて出力する。

```text
VBA
└─ Forms
    ├─ Code
    └─ ControlsCsv
```

## 6.1 フォームコード

出力先：

```text
VBA\Forms\Code
```

ファイル名：

```text
フォーム名.cls
```

例：

```text
frmProductDailyReport.cls
```

処理：
1. `CodeModule` から全コードを取得する
2. Attribute行を除去する
3. `.cls` ファイルとして保存する

コードが1行も存在しない場合は出力しない。

---

## 6.2 コントロール情報CSV

出力先：

```text
VBA\Forms\ControlsCsv
```

ファイル名：

```text
フォーム名.csv
```

例：

```text
frmProductDailyReport.csv
```

CSVにはUserForm上の各コントロールについて以下を出力する。

| 列 | 内容 |
|---|---|
| ControlName | コントロール名 |
| ControlType | コントロール種類 |
| Caption | Caption |
| Left | 左位置 |
| Top | 上位置 |
| Width | 幅 |
| Height | 高さ |
| FontName | フォント名 |
| FontSize | フォントサイズ |
| FontBold | 太字 |
| ForeColor | 文字色 |
| BackColor | 背景色 |

CSVヘッダー：

```text
ControlName,ControlType,Caption,Left,Top,Width,Height,FontName,FontSize,FontBold,ForeColor,BackColor
```

文字列項目はCSV用にダブルクォートで囲む。

文字列内に `"` が存在する場合は、

```text
"
```

を

```text
""
```

へ変換する。

---

# 7. 標準モジュール

対象：

```text
vbext_ct_StdModule
```

出力先：

```text
VBA\Modules
```

ファイル形式：

```text
.bas
```

例：

```text
modRepositoryExport.bas
```

Excel VBA標準の

```vb
VBComponent.Export
```

を使用して一旦出力する。

その後、

```vb
ConvertFileToUtf8WithoutAttribute
```

を実行し、

- Attribute行除去
- UTF-8化

を行う。

---

# 8. クラスモジュール

対象：

```text
vbext_ct_ClassModule
```

出力先：

```text
VBA\Classes
```

ファイル形式：

```text
.cls
```

例：

```text
Class1.cls
```

標準モジュールと同様に、

```vb
VBComponent.Export
```

で出力後、

```vb
ConvertFileToUtf8WithoutAttribute
```

を実行する。

---

# 9. Documentモジュール

対象：

```text
vbext_ct_Document
```

主な対象：

```text
ThisWorkbook
Sheet1
Sheet2
...
```

出力先：

```text
VBA\Documents
```

ファイル形式：

```text
.cls
```

Documentモジュールについては `VBComponent.Export` を使用せず、`CodeModule` からコード部分のみを取得する。

処理：
1. `CodeModule.CountOfLines` を取得
2. 全コードを取得
3. Attribute行を除去
4. 一旦テキストファイルとして保存
5. UTF-8へ変換
---

# 10. 文字コード

Git等で管理しやすくするため、エクスポートしたソースファイルはUTF-8を基本とする。

標準モジュール・クラスモジュールについては、

```vb
ConvertFileToUtf8WithoutAttribute
```

を使用する。

UserFormコードについては、

```vb
SaveUtf8NoBom
```

を使用する。

Documentモジュールについては、

```vb
ConvertFileToUtf8
```

を使用する。

---

# 11. Attribute行

VBA標準のExportでは、ソースコード中に以下のようなAttribute情報が含まれる場合がある。

```text
Attribute VB_Name = "Module1"
Attribute VB_GlobalNameSpace = False
```

Git上で通常編集するVBAソースとして不要なため、

```vb
RemoveAttributeLines
```

によって除去する。

---

# 12. フォルダ作成

必要なフォルダが存在しない場合は、

```vb
CreateFolderIfNotExists
```

によって作成する。

作成対象：

```text
ExcelRepository
対象ブック名_repository
VBA
VBA\Forms
VBA\Forms\Code
VBA\Forms\ControlsCsv
VBA\Modules
VBA\Classes
VBA\Documents
```

既に存在する場合はそのまま使用する。

---

# 13. Excel本体情報

VBAコンポーネントの出力完了後、

```vb
ExportWorkbookRepository(wb, baseFolder)
```

を実行する。

この処理によって、VBA以外のExcelブック情報をリポジトリへ出力する。

```text
対象ブック名_repository
├─ VBA
└─ Excel本体情報
```

ただし、`ExportWorkbookRepository` の実装内容は本仕様書作成時に提示されたコードには含まれていないため、詳細な出力内容については別途確認・仕様化する。

---

# 14. 処理フロー

```text
ExportRepositoryForActiveWorkbook
        │
        ├─ ActiveWorkbook取得
        │
        ├─ 保存済みか確認
        │
        ├─ VBProject取得
        │
        ├─ 出力フォルダ作成
        │
        ├─ VBComponentsループ
        │       │
        │       ├─ UserForm
        │       │    ├─ ControlsCsv出力
        │       │    └─ FormCode出力
        │       │
        │       ├─ 標準Module
        │       │    └─ .bas出力
        │       │
        │       ├─ ClassModule
        │       │    └─ .cls出力
        │       │
        │       └─ Document
        │            └─ .cls出力
        │
        ├─ ExportWorkbookRepository
        │
        └─ 完了メッセージ
```

---

# 15. 主な関数一覧

| 関数 | 役割 |
|---|---|
| ExportRepositoryForActiveWorkbook | リポジトリエクスポート全体制御 |
| ExportFormControlsToCsv | UserFormコントロール情報CSV出力 |
| ExportFormCodeToCls | UserFormコード出力 |
| ExportStdModuleToBas | 標準モジュール出力 |
| ExportClassModuleToCls | クラスモジュール出力 |
| ExportDocumentModuleCodeOnly | ThisWorkbook・Sheetコード出力 |
| GetNameWithoutExtension | 拡張子を除いたファイル名取得 |
| CreateFolderIfNotExists | フォルダ作成 |
| GetCtrlCaptionSafe | Caption安全取得 |
| EscapeCsv | CSV文字列エスケープ |
| ConvertFileToUtf8WithoutAttribute | Attribute除去＋UTF-8変換 |
| GetCtrlPropertySafe | UserFormコントロールプロパティ安全取得 |

---

# 16. 外部・別処理として使用している関数

提示されたコード内では実装内容を確認できない関数：

```vb
ExportWorkbookRepository
SaveUtf8NoBom
ReadTextFile
RemoveAttributeLines
ConvertFileToUtf8
```

これらについては別モジュールまたは同一モジュール後半に実装されているものと考えられる。

詳細仕様については、実際のコードを確認した上で本仕様書へ追加する。

---

# 17. リポジトリの目的

このリポジトリは、Excelブックそのものでは確認しづらい変更内容をテキスト化し、

- Gitによるバージョン管理
- SourceTree等による差分確認
- AIによるコード確認
- UserFormレイアウト変更確認
- コード検索
- 過去バージョンとの比較

を行いやすくすることを目的とする。

特にUserFormについて、

```text
Code
ControlsCsv
```

を分離することで、

```text
処理ロジック
```

と

```text
画面レイアウト
```

をそれぞれ差分管理できる構造とする。

---

# 18. 関連ユーティリティ

エクスポートされたリポジトリ内のテキストファイルについて、Git差分を安定させるため、別途ファイル正規化ユーティリティを使用する。

正規化ユーティリティは、

```text
正規化スクリプト自身が置かれたフォルダ
```

を基準として、

```text
直下
および
すべてのサブフォルダ
```

を再帰的に走査する。

したがって、今回のような、

```text
VBA
├─ Forms
│  ├─ Code
│  └─ ControlsCsv
├─ Modules
├─ Classes
└─ Documents
```

という多階層構造についても処理対象とする。

特定の `Source` フォルダ名や固定されたフォルダ構成には依存しない。