# CS-17-Molecular 定期報告データ作成ツール

本プロジェクトは、臨床試験「CS-17」および「CS-17-Molecular」の登録データと、EDCから出力された各ドメインのCSVデータを統合し、定期報告用の集計報告書（Excel形式）を自動生成するRスクリプトを管理しています。

## 処理の概要

1.  **データの自動インポート**
    `input/rawdata` 内のCSVファイルをスキャンし、ファイル名から「本体登録データ」「Molecular登録データ」「各ドメイン（DM, MH, EC等）」を判別して動的に読み込みます。
2.  **症例情報の紐付け**
    共通の「登録コード」をキーとして、本体試験（CS-17）と付随研究（Molecular）の症例を結合し、解析対象集団を特定します。
3.  **変数の算出・加工**
    - 年齢、性別等の基本情報の整理。
    - 治療フェーズ（導入、地固め、サルベージ）およびコース数の整理。
    - 寛解有無、再発有無、生存日数などの臨床評価項目の算出。
4.  **Excel報告書の生成**
    `openxlsx` を使用し、施設別集計（Table1）および症例ごとの詳細一覧（Table2）を整形して出力します。

---

## 実行環境と準備

### 動作環境

- **OS**: **macOS**（パスの仕様上、Windows環境では正しく動作しない可能性があります）
- **Language**: R (version 3.6以上推奨)
- **Required Packages**: `dplyr`, `stringr`, `here`, `openxlsx`

### 事前準備

1. **リポジトリのクローン**
   ターミナルを起動し、リポジトリをローカルにクローンします。
   ```bash
   git clone https://github.com/nnh/JALSG-CS-17-Molecular.git
   ```
2. ディレクトリの作成とテンプレートの配置
   プロジェクトのルートディレクトリ（JALSG-CS-17-Molecular）にて、以下の準備を行ってください。

- output/QC フォルダを作成する。
- output フォルダ直下に、テンプレートとなるExcelファイル（CS-17-Molecular定期報告\_YYYYMMDD.xlsx）を格納する。

### 実行方法

#### ディレクトリ構造の確認

以下の構成になっていることを確認してください。

```Plaintext
JALSG-CS-17-Molecular/
├── input/
│ ├── rawdata/ # 各ドメインのCSVを格納
│ └── ext/ # diseases.csv, facilities.csv を格納
├── output/
│ ├── CS-17-Molecular定期報告\_YYYYMMDD.xlsx (テンプレート)
│ └── QC/ # 空のフォルダ
└── cs-17-molecular.R
```

#### スクリプトの実行

RStudioで cs-17-molecular.R を開き Source を実行します。

#### 出力の確認

処理完了後、output/QC/ フォルダ内に、実行日の日付が付与されたExcelファイル（例: CS-17-Molecular定期報告\_20260402.xlsx）が生成されます。

## 注意事項

- macOS 依存: パス指定にmacOS固有の形式を含んでいるため、Windows環境での動作は保証されません。

- パスの修正: スクリプト内の parent_path が自身のローカル環境のパスと一致しているか、実行前に必ず確認してください。

- データの上書き防止: 読み込まれたrawデータには lockBinding が適用されます。再読み込みが必要な場合は、Rセッションを再起動するか rm(list=ls()) を実行してください。

- 出力形式: バリデーション用のCSV（validation_table1.csv 等）は、Excelでの直接参照を考慮して CP932 (Shift-JIS) で出力されます。
