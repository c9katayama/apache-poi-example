# Apache POI サンプルプロジェクト

このプロジェクトは、Apache POIを使用してExcel、Word、PowerPointファイルを操作するサンプルコードを集めたものです。

## 機能概要

以下のような機能のサンプルコードを提供しています:

### Excel操作
- 基本的なワークブック作成 (BasicWorkbookExample)
- グラフ作成 (ChartExample) 
- 数式の設定 (FormulaExample)
- VLOOKUP関数の使用 (FormulaVLookupExample)
- セルスタイルの適用 (StyleExample)
- 入力規則の設定 (ValidationExample)
- 大容量ワークブックのエクスポート (LargeWorkbookExportExample)
- 大容量ワークブックの読み込み (LargeWorkbookReaderExample)

### Word操作
- 基本的なドキュメント作成 (BasicDocumentExample)
- フォームテンプレートの作成と値の設定 (FormExample)
- 画像の挿入 (ImageExample)
- 表の作成 (TableExample)

### PowerPoint操作
- 基本的なプレゼンテーション作成 (BasicPresentationExample)
- 画像の挿入 (ImageExample)
- 図形の追加 (ShapeExample)

### 共通機能
- ファイルの暗号化 (EncryptionExample)
- メタデータの設定 (MetadataExample)
- 印刷設定 (PrintSettingExample)

## 実行環境

- Java 11以上
- Gradle 7.0以上

## プロジェクトのセットアップ

### Eclipse

1. File > Import > Gradle > Existing Gradle Project を選択
2. プロジェクトのルートディレクトリを選択
3. Finish をクリック

### VS Code

1. VS Code を起動
2. Java Extension Pack をインストール
3. プロジェクトのルートディレクトリを開く
4. コマンドパレットから「Java: Import Java Project」を実行
5. Gradle プロジェクトとしてインポート

## ビルドと実行

```bash
# ビルド
./gradlew build

# 特定のクラスの実行例
./gradlew run --args="poi_example.excel.BasicWorkbookExample"
```

## サンプルファイルの出力先

生成されたファイルは `output` ディレクトリに保存されます。


## ライセンス

このプロジェクトは Apache License 2.0 の下で公開されています。

## 参考リンク

- [Apache POI公式サイト](https://poi.apache.org/)
- [Apache POI API Documentation](https://poi.apache.org/apidocs/index.html) 