# Excel VBAプロジェクト リポジトリインポート仕様書

## 1. 概要

Excelリポジトリエクスポート機能によってテキスト化されたVBAプロジェクト情報を、アクティブなExcelブックへ反映する。

主な対象は以下とする。

- 標準モジュール
- クラスモジュール
- UserFormレイアウト
- UserFormコード
- ThisWorkbook / Sheet等のDocumentモジュールコード

Workbook配下に出力される、

- WorkbookInfo.txt
- Sheets.csv
- Tables.csv
- Names.csv

については、現時点ではインポート対象としない。

---

# 2. 実行入口

メイン処理：

```vb
Public Sub ImportRepositoryForActiveWorkbook()
```

実行時の `ActiveWorkbook` を取込対象とする。

---

# 3. 実行条件

以下の場合は処理を中止する。

## 3.1 アクティブブックが存在しない

```text
アクティブなブックがありません。
```

## 3.2 ブックが未保存

`Workbook.Path` が空の場合、以下を表示して終了する。

```text
ブックがまだ保存されていません。先に保存してください。
```

---

# 4. 参照するリポジトリ

対象Excelブックと同じフォルダにある、

```text
ExcelRepository
```

以下のリポジトリを使用する。

基本パス：

```text
ブック保存フォルダ
└─ ExcelRepository
    └─ 対象ブック名_repository
```

対象ブック名には拡張子を含めない。

例：

```text
売上管理.xlsm

↓

ExcelRepository
└─ 売上管理_repository
```

---

# 5. インポート対象フォルダ

使用するVBAリポジトリ構造：

```text
対象ブック名_repository
│
├─ VBA
│   │
│   ├─ Forms
│   │   ├─ Code
│   │   └─ ControlsCsv
│   │
│   ├─ Modules
│   ├─ Classes
│   └─ Documents
│
└─ Workbook
```

インポート対象は `VBA` 配下のみ。

```text
VBA
├─ Forms
│  ├─ Code
│  └─ ControlsCsv
├─ Modules
├─ Classes
└─ Documents
```

`Workbook` フォルダは参照しない。

---

# 6. 全体処理フロー

```text
ImportRepositoryForActiveWorkbook
        │
        ├─ ActiveWorkbook取得
        │
        ├─ 保存済み確認
        │
        ├─ VBProject取得
        │
        ├─ リポジトリパス決定
        │
        ├─ 標準モジュール取込
        │
        ├─ クラスモジュール取込
        │
        ├─ VBComponentsループ
        │       │
        │       ├─ UserForm
        │       │    ├─ レイアウト反映
        │       │    └─ コード反映
        │       │
        │       └─ Document
        │            └─ コード反映
        │
        └─ 完了メッセージ
```

---

# 7. 標準モジュール取込

処理：

```vb
ImportStdModulesFromBas
```

取込元：

```text
VBA\Modules
```

対象ファイル：

```text
*.bas
```

処理内容：
1. `Modules` フォルダの全ファイルを取得
2. 拡張子 `.bas` のファイルのみ対象
3. ファイル名からモジュール名を取得
4. 同名のVBComponentが存在する場合は削除
5. 新しい標準モジュールを作成
6. ファイル名と同じモジュール名を設定
7. UTF-8テキストを読込
8. VBEエクスポートヘッダーを除去
9. コードを新規モジュールへ挿入

同名モジュールの扱いは、

```text
既存モジュール
    ↓
削除
    ↓
新規モジュール作成
    ↓
リポジトリのコードを挿入
```

となる。

標準モジュールは「編集」ではなく「置換」と考える。

---

# 8. クラスモジュール取込

処理：

```vb
ImportClassModulesFromCls
```

取込元：

```text
VBA\Classes
```

対象：

```text
*.cls
```

処理内容は標準モジュールと同じ。

```text
既存クラス
    ↓
削除
    ↓
新規クラスモジュール作成
    ↓
クラス名設定
    ↓
コード挿入
```

同名クラスが存在する場合は削除して作り直す。

---

# 9. 標準・クラスモジュールの削除仕様

現在の処理は、

```text
リポジトリに存在するファイル
```

についてのみ処理する。

したがって、

```text
Excel側には存在する
しかし
リポジトリ側にはファイルが存在しない
```

モジュールについては削除されない。

つまり現在の仕様は、

```text
リポジトリとの完全同期
```

ではなく、

```text
リポジトリに存在するモジュールをExcelへ上書き・追加
```

である。

リポジトリから削除されたモジュールをExcel側からも削除する同期処理は、現時点では実装されていない。

---

# 10. UserForm処理

既存Excelブック内のUserFormを対象とする。

`VBComponents` を走査し、

```vb
vbext_ct_MSForm
```

の場合、

```vb
ImportFormLayoutFromCsv
ImportFormCodeFromCls
```

を順番に実行する。

処理順：

```text
UserForm
   │
   ├─ ControlsCsvからレイアウト反映
   │
   └─ Codeからフォームコード反映
```

---

# 11. UserFormの新規作成

現在の処理ではUserFormそのものの新規作成は行わない。

インポート処理は、

```text
Excelブック内に既に存在するUserForm
```

を基準としている。

したがって、

```text
リポジトリには frmNew.csv / frmNew.cls が存在する
Excel側には frmNew が存在しない
```

場合でも、現在の処理だけでは `frmNew` は作成されない。

---

# 12. UserFormレイアウト取込

処理：

```vb
ImportFormLayoutFromCsv
```

取込元：

```text
VBA\Forms\ControlsCsv
```

ファイル：

```text
フォーム名.csv
```

例：

```text
frmProductDailyReport.csv
```

CSVファイルが存在しない場合は何もしない。

空ファイルの場合も何もしない。

---

# 13. UserForm CSV形式

最低限必要な列：

```text
ControlName
ControlType
Caption
Left
Top
Width
Height
```

追加項目：

```text
FontName
FontSize
FontBold
ForeColor
BackColor
```

現在想定しているCSV構造：

```text
ControlName,
ControlType,
Caption,
Left,
Top,
Width,
Height,
FontName,
FontSize,
FontBold,
ForeColor,
BackColor
```

先頭行はヘッダーとして扱い、2行目以降を処理する。

---

# 14. CSV改行処理

CSV読込後、

```text
CRLF
CR
```

をすべて、

```text
LF
```

へ統一する。

その後、

```vb
Split(allText, vbLf)
```

により行単位へ分割する。

したがって、

```text
CRLF
LF
CR
```

が混在していても読み込める構成となっている。

---

# 15. CSV解析

各CSV行は、

```vb
ParseCsvLine
```

によって解析する。

対応内容：
- カンマ区切り
- ダブルクォート囲み
- フィールド内カンマ
- `""` によるダブルクォートのエスケープ

単純な `Split(line, ",")` は使用しない。

---

# 16. UserForm既存コントロール

CSVの `ControlName` を使用して、

```vb
designer.Controls(ctrlName)
```

から既存コントロールを検索する。

存在する場合は、そのコントロールを使用する。

---

# 17. UserFormコントロールの新規追加

CSVには存在するがUserForm上に存在しないコントロールの場合、

```vb
GetProgIdFromCtrlType
```

によって `ControlType` からProgIDを取得する。

ProgIDが取得できる場合、

```vb
designer.Controls.Add
```

によって新しいコントロールを作成する。

---

# 18. 対応コントロール

現在、新規作成に対応しているコントロール：

| ControlType | ProgID |
|---|---|
| TextBox | Forms.TextBox.1 |
| Label | Forms.Label.1 |
| CommandButton | Forms.CommandButton.1 |
| ComboBox | Forms.ComboBox.1 |
| ListBox | Forms.ListBox.1 |
| CheckBox | Forms.CheckBox.1 |
| OptionButton | Forms.OptionButton.1 |
| Frame | Forms.Frame.1 |
| MultiPage | Forms.MultiPage.1 |
| TabStrip | Forms.TabStrip.1 |
| ScrollBar | Forms.ScrollBar.1 |
| SpinButton | Forms.SpinButton.1 |
| Image | Forms.Image.1 |

未対応の `ControlType` の場合は空文字を返し、コントロールを作成しない。

---

# 19. ControlType確認

既存または新規取得したコントロールについて、

```vb
TypeName(ctrl) = ctrlType
```

の場合のみレイアウトを反映する。

`ctrlType` が空の場合も反映可能とする。

したがって、

```text
同じControlNameだが種類が異なる
```

場合には、現在の処理では既存コントロールを削除・置換せず、レイアウト反映を行わない。

---

# 20. UserFormレイアウト反映項目

処理：

```vb
ApplyControlLayout
```

反映する項目：

```text
Left
Top
Width
Height
Caption
Font.Name
Font.Size
Font.Bold
ForeColor
BackColor
```

最低限必須：

```text
Left
Top
Width
Height
Caption
```

フォントや色についてはCSVに値が存在する場合のみ反映する。

---

# 21. UserFormレイアウト変更

座標・サイズはCSVの値を直接使用する。

```vb
ctrl.Left
ctrl.Top
ctrl.Width
ctrl.Height
```

したがって、PowerPoint等で編集したControlsCsvをインポートすることにより、UserFormのレイアウト変更を反映できる。

---

# 22. UserFormコントロール削除

現在の処理では、

```text
UserForm上には存在する
しかし
CSVには存在しない
```

コントロールを削除する処理はない。

つまりUserFormレイアウトについても、

```text
完全同期
```

ではない。

現在の動作は、

```text
CSVにあるコントロール
    ↓
既存なら更新
存在しなければ追加
```

である。

CSVから削除されたコントロールをUserFormから削除する処理は未実装。

---

# 23. UserFormコード取込

処理：

```vb
ImportFormCodeFromCls
```

取込元：

```text
VBA\Forms\Code
```

ファイル：

```text
フォーム名.cls
```

対象UserFormと同名の `.cls` が存在する場合のみ処理する。

---

# 24. UserFormコード置換

フォームコードについては既存コードを完全に削除する。

```text
現在のCodeModule
    ↓
全行削除
    ↓
リポジトリの.clsを読込
    ↓
コード挿入
```

したがって、UserFormコードについてはファイル内容による完全置換となる。

ただし `.cls` ファイルそのものが存在しない場合は何もしない。

---

# 25. Documentモジュール

対象：

```text
ThisWorkbook
Sheet1
Sheet2
...
```

判定：

```vb
vbext_ct_Document
```

処理：

```vb
ImportDocumentModuleFromCls
```

取込元：

```text
VBA\Documents
```

ファイル：

```text
VBComponent名.cls
```

例：

```text
ThisWorkbook.cls
Sheet1.cls
```

---

# 26. Documentコード置換

対象 `.cls` が存在する場合、

```text
既存CodeModule
    ↓
全コード削除
    ↓
リポジトリコード挿入
```

とする。

Documentモジュールそのものは削除・新規作成しない。

既存のExcel Documentオブジェクトへコードのみを反映する。

---

# 27. Documentモジュールとシート名

Documentモジュールのファイル名には、

```text
Sheet1
Sheet2
ThisWorkbook
```

などのVBComponent名を使用する。

ワークシートの表示名ではなく、

```text
CodeName
```

側の名前に対応する。

---

# 28. VBEヘッダー除去

`.bas` `.cls` 読込後、

```vb
RemoveVbeExportHeader
```

を実行する。

VBE標準Exportで出力される、

```text
VERSION
BEGIN
END
Attribute
```

等のVBE管理情報を除去し、コード部分のみをCodeModuleへ挿入することを目的とする。

具体的な除去仕様は `RemoveVbeExportHeader` の実装に依存する。

---

# 29. UTF-8読込

各ソースファイルは、

```vb
ReadUtf8Text
```

によって読み込む。

対象：

```text
*.bas
*.cls
*.csv
```

リポジトリのソースファイルはUTF-8で保存されていることを前提とする。

---

# 30. Workbook情報について

エクスポート側では、

```text
Workbook
├─ WorkbookInfo.txt
├─ Sheets.csv
├─ Tables.csv
└─ Names.csv
```

を出力するが、現時点のインポート処理ではこれらを使用しない。

したがってWorkbook情報は、

```text
差分確認
構造確認
資料
AIによる確認
```

を主目的とする。

現時点では、

```text
Sheets.csvからシート新規作成
Tables.csvからテーブル復元
Names.csvから定義名復元
```

などの処理は実装していない。

---

# 31. エクスポートとインポートの対応

| エクスポート | インポート | 状態 |
|---|---|---|
| VBA\Modules\*.bas | 標準モジュール | 対応 |
| VBA\Classes\*.cls | クラスモジュール | 対応 |
| VBA\Forms\Code\*.cls | UserFormコード | 対応 |
| VBA\Forms\ControlsCsv\*.csv | UserFormレイアウト | 対応 |
| VBA\Documents\*.cls | Documentコード | 対応 |
| Workbook\WorkbookInfo.txt | なし | 参照用 |
| Workbook\Sheets.csv | なし | 参照用 |
| Workbook\Tables.csv | なし | 参照用 |
| Workbook\Names.csv | なし | 参照用 |

---

# 32. 同期レベル

現時点のインポート処理は、対象によって同期方法が異なる。

## 標準モジュール

```text
リポジトリに存在
→ 既存削除＋新規作成

リポジトリに存在しない
→ Excel側はそのまま
```

## クラスモジュール

標準モジュールと同じ。

## UserFormコード

```text
.clsあり
→ コード完全置換

.clsなし
→ 変更なし
```

## UserFormコントロール

```text
CSVに存在＋フォームにも存在
→ 更新

CSVに存在＋フォームに存在しない
→ 対応種類なら新規追加

フォームに存在＋CSVに存在しない
→ 削除しない
```

## Documentコード

```text
.clsあり
→ コード完全置換

.clsなし
→ 変更なし
```

---

# 33. 現時点での非対応事項

以下は現在実装されていない。

- リポジトリから削除された標準モジュールの自動削除
- リポジトリから削除されたクラスモジュールの自動削除
- UserFormそのものの新規作成
- UserFormそのものの削除
- CSVから削除されたUserFormコントロールの削除
- ControlType変更時の既存コントロール置換
- Workbook情報からのブック構造復元
- シート新規作成・削除
- Excelテーブル復元
- 定義名復元
---

# 34. エラー処理

現行コードでは各処理を包括する共通エラーハンドラは設けていない。

一部の取得処理について、

```vb
On Error Resume Next
```

を限定的に使用している。

主な使用箇所：
- VBComponent存在確認
- UserForm Designer取得
- UserFormコントロール取得
- UserFormコントロール追加
- コントロールプロパティ反映

したがって、インポート処理全体で予期しないエラーが発生した場合には、その時点で処理が停止する可能性がある。

---

# 35. 使用上の前提

VBAプロジェクトへプログラムからアクセスするため、Excel側で、

```text
VBA プロジェクト オブジェクト モデルへのアクセスを信頼する
```

設定が有効である必要がある。

また、取込対象のブックは保存済みである必要がある。

---

# 36. 目的

このインポート機能の目的は、

```text
Excelファイル
    ↓
テキスト化
    ↓
Git管理
    ↓
編集・差分確認
    ↓
Excelへ反映
```

という開発サイクルを成立させることにある。

特にVBAコードについては、

- Gitによる変更履歴管理
- VS Code等での確認
- AIによるコードレビュー
- 複数バージョン比較

を可能にする。

UserFormについては、

```text
フォームコード
```

と、

```text
コントロール配置情報
```

を分離して管理する。

これにより、UserFormのレイアウトについてもExcel VBEだけに依存せず編集・管理できる。

---

# 37. 関連処理

本インポート処理は以下と対になる。

```text
ExportRepositoryForActiveWorkbook
```

また、リポジトリ内のテキストファイルについては、Git差分を安定させるため正規化ツールを使用する。

正規化ツールは、

```text
スクリプトが置かれたフォルダ以下
```

をすべて再帰的に走査する。

特定のフォルダ名や階層構造には依存しない。

`.git` フォルダは対象外とする。