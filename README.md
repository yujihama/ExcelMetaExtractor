# Excel Metadata Extractor

高度なExcelメタデータ抽出と解析のためのStreamlitベースのWebアプリケーション。複雑なスプレッドシート構造とXMLベースの図形情報を正確に解析します。

## 主な機能

### メタデータ解析
- Excelファイルからの包括的なメタデータ抽出
- シート構造、セル情報、フォーマット設定の詳細分析
- 数式、参照関係、依存関係の解析
- コメント、ノート、変更履歴の抽出

### 図形・グラフ解析
- XMLベースの図形情報の詳細抽出（drawing_extractor.py）
- VML要素の解析と構造化（vml_processor.py）
- チャート構造とデータ系列の分析（chart_processor.py）
- 図形の位置、サイズ、スタイル情報の抽出

### 領域分析機能
- スプレッドシートの論理的な領域検出（region_detector.py）
- セル領域の意味解析と分類（region_analyzer.py）
- データ構造の自動認識
- テーブル領域の特定と解析

### AI統合機能
- OpenAI gptによるメタデータの意味解析
- 図形とチャートの自動分類と説明生成
- データ構造の自動検出と推奨事項の提示
- 日本語コンテンツの高精度な解析と理解

### インターフェース
- Streamlitによる直感的なWebインターフェース
- リアルタイムのデータ可視化と探索機能
- インタラクティブな解析結果の表示
- カスタマイズ可能なレポート生成

## システム構成

### コアモジュール
- `excel_metadata_extractor.py`: メインの抽出エンジン
- `cell_processor.py`: セルデータの詳細処理
- `chart_processor.py`: チャートデータの解析
- `drawing_extractor.py`: 図形要素の抽出
- `vml_processor.py`: VML形式の解析
- `region_analyzer.py` & `region_detector.py`: 領域検出・分析
- `openai_helper.py`: gpt連携機能
- `logger.py`: 詳細なログ管理

### 技術スタック
- Python 3.11
- Streamlit: Webインターフェース
- OpenPyXL: Excel処理
- Pandas: データ操作
- OpenAI API: メタデータ分析
- ElementTree: XML解析
- Matplotlib: データ可視化

## インストール方法

```bash
# リポジトリのクローン
git clone https://github.com/your-username/excel-metadata-extractor.git
cd excel-metadata-extractor

# 依存パッケージのインストール
pip install -r requirements.txt

# 環境変数の設定
export OPENAI_API_KEY=your_api_key

# アプリケーションの起動
streamlit run main.py
```

## 使用方法

1. Webブラウザで`http://localhost:5000`にアクセス
2. Excelファイルをアップロード
3. 解析オプションを選択：
   - 基本メタデータ解析
   - 図形・チャート解析
   - 領域検出と分析
   - AI支援解析
4. 解析結果の表示：
   - ファイルプロパティ
   - シート構造
   - セル情報
   - 図形・チャートデータ
   - 検出された領域情報
   - AIによる分析レポート
5. 結果のエクスポートとレポート生成

## 環境設定

### 必要な環境変数
- `OPENAI_API_KEY`: OpenAI APIキー（AI解析機能に必須）

### システム要件
- Python 3.11以上
- インターネット接続（AI機能使用時）
- ディスク容量: 最低500MB（アプリケーションと依存関係用）

## パフォーマンス注意事項

- 大規模なExcelファイル（100MB以上）の処理には追加の処理時間が必要
- 複雑なチャートや多数の図形要素を含むファイルは解析に時間がかかる場合あり
- AI解析機能の使用には有効なOpenAI APIキーと十分なAPIクレジットが必要
- メモリ使用量は処理するファイルのサイズに応じて増加

## 注意事項

- 機密情報を含むExcelファイルの処理時は適切なセキュリティ対策を実施してください
- AI解析機能の使用にはAPI利用料金が発生します
- 大量のファイル処理時はシステムリソースの使用状況に注意してください
- 処理可能なファイルサイズには環境に応じた制限があります