# ExcelVbaAgentBridge

**Excel VBAを、GitとAIエージェントで開発するための橋渡しツール。**

Excel の VBA プロジェクトをテキストファイルと CSV に書き出し、Git による差分管理や AI エージェントによる編集を行った後、既存の Excel ブックへ取り込むための VBA マクロです。

このツールは私自身の開発作業で使うために作成しました。**Excel の機能をすべてテキスト化・復元するものではありません。** また、Excel のマクロを AI が自動実行する仕組みでもありません。

## ダウンロード

最新版のExcelマクロ（Builder・Restorer）は、[Releasesページ](https://github.com/TakaoOotuka/ExcelVbaAgentBridge/releases/latest)からダウンロードできます。

## なぜ作ったのか

Excel VBAのコードをAIに書いてもらっても、VBEへのコピー＆ペーストを繰り返すのでは、開発の流れが途切れてしまいます。そこで、VBAコードとUserFormの一部レイアウト情報をファイルとして書き出し、AIエージェントが編集したものをExcelへ戻せるようにしました。

## どんな問題を解決するか

通常、Excel VBA のコードは `.xlsm` の中に格納されています。そのままでは Git でコードの変更箇所を追跡しづらく、AI エージェントが複数のモジュールや UserForm の構成を直接編集することも難しくなります。

本ツールは、次の流れを補助します。

1. **Builder**：対象ブックの VBA コードと UserForm の一部のレイアウト情報をファイルへ書き出す
2. **Git / AI エージェント**：書き出した `.bas`、`.cls`、`.csv` を管理・編集する
3. **Restorer**：編集結果を対象の Excel ブックに反映する
4. **Excel 上での確認**：コンパイル・画面表示・動作確認を人間が行う

## 収録内容

- `Builder/modRepositoryExport.bas` — VBA プロジェクトのエクスポート
- `Builder/modWorkbookExport.bas` — ブック、シート、テーブル、名前定義などの**情報の**エクスポート
- `Builder/modUtf8File.bas` — テキストファイル入出力補助
- `Restorer/modRepositoryImport.bas` — VBA プロジェクトのインポート
- `Restorer/modUtf8File.bas` — テキストファイル入出力補助

同名の `modUtf8File.bas` は Builder と Restorer で内容が異なるため、それぞれ別フォルダーに収録しています。

## 出力の例

対象ブックと同じフォルダーに、次のような階層を作成します。

```text
ExcelRepository/
└── [ブック名]_repository/
    ├── VBA/
    │   ├── Modules/       # 標準モジュール (.bas)
    │   ├── Classes/       # クラスモジュール (.cls)
    │   ├── Documents/     # ThisWorkbook / Sheet のコード (.cls)
    │   └── Forms/
    │       ├── Code/      # 既存 UserForm のコード (.cls)
    │       └── ControlsCsv/ # UserForm のコントロール情報 (.csv)
    └── Workbook/
        ├── WorkbookInfo.txt
        ├── Sheets.csv
        ├── Tables.csv
        └── Names.csv
```

※ `Workbook/` の各ファイルは **参照用の情報** です。現行の Restorer はシート、テーブル、名前定義、セル値や数式をこれらの CSV から復元しません。

## 利用に必要なもの

- Windows版Excel（VBAが使用できる環境）
- VBAプロジェクトへのアクセス許可
- Git（履歴管理する場合）
- AIエージェント（AIで編集する場合。ツール単体の利用には不要）

## 準備と実行手順

1. **作業対象の `.xlsm` をバックアップ**してください。必ずコピーで試してください。
2. Excel の設定で、必要なマクロ実行許可と **「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」**を有効にします。後者は Excel の［ファイル］→［オプション］→［トラスト センター］→［トラスト センターの設定］→［マクロの設定］で確認できます。信頼できるファイルでのみ使用してください。
3. Builder 側の Excel ブックに `src/Builder/` の3モジュールを配置します（提供される `.xlsm` がある場合はそれを使用できます）。Restorer 側には `src/Restorer/` の2モジュールを配置します。
4. **対象ブックを保存し、アクティブにしてから** Builder の `ExportRepositoryForActiveWorkbook` を実行します。マクロは `ActiveWorkbook` を対象にします。
5. 生成されたテキストと CSV を Git で管理し、必要に応じて AI エージェントで変更します。
6. **再度対象ブックのコピーをバックアップ**したうえで、対象ブックをアクティブにし、Restorer の `ImportRepositoryForActiveWorkbook` を実行します。
7. VBE でコンパイルし、UserForm の動作・イベント処理を含めて確認してください。

> 注意：実行時に意図したブックがアクティブでないと、別のブックを対象としてしまいます。操作するブックを毎回確認してください。

## UserForm の CSV

対象は既存 UserForm 内のコントロールです。CSV は以下の列を出力します。

`ControlName, ControlType, Caption, Left, Top, Width, Height, FontName, FontSize, FontBold, ForeColor, BackColor`

Restorer は、既存のコントロールには対応するプロパティを反映します。一部の種類は、CSV にあればコントロールを新規追加できます。**UserForm 自体を新規作成する機能はありません。** CSV から行を削除しても、対応する既存コントロールが削除されるわけではありません。すべてのプロパティ、画像、複合コントロール構造などを完全再現できるわけではありません。

## 既知の制約・注意点

- Restorer は既存の標準モジュール・クラスモジュールを削除して再作成する方式を含みます。UserForm コード、ThisWorkbook / Sheet のコードも対象ファイルが存在すると置換します。**取り込み前には必ずバックアップしてください。**
- すべての参照設定、フォーム定義、外部コントロール、Workbook の内部情報を復元するものではありません。
- 出力内容に含まれる `WorkbookInfo.txt` の絶対パス、VBA コード内の機密情報、接続文字列、顧客情報などを **GitHub に公開しない**でください。実務ブックから出力したリポジトリの無加工公開は推奨しません。
- VBA プロジェクトがロック・保護されている場合や、環境によりアクセスが制限される場合は動作しません。
- CSV の項目・改行・文字コードを変更した場合、再取込できないことがあります。
- 現行マクロには取り込み前の自動バックアップ、差分プレビュー、ロールバックはありません。
- 動作対象・Office バージョンの網羅的な互換試験は実施していません。手元での実運用はしていますが、あらゆる環境での動作を保証しません。

## PowerPoint で UserForm レイアウトを調整する応用

UserForm の配置情報を CSV にしているため、別途作成した PowerPoint 連携処理などを介してレイアウトを編集する応用もできます。**PowerPoint 連携機能自体は、今回収録した5つの `.bas` には含まれていません。** 別記事で紹介する予定です。

## ライセンス

ライセンスは公開前に設定します。再利用・再配布条件については、追加する `LICENSE` を確認してください。現時点ではライセンス未設定です。

## AI エージェントを使うときの心得

差分を確認してから Restorer で反映してください。AI に既存コードの全面置換を無条件に許可せず、モジュールごとに変更内容を確認する運用をおすすめします。実際の Excel 上でのコンパイルと動作確認は不可欠です。
