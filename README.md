# Excel Metadata Extractor

高度なExcelメタデータ抽出と解析のためのStreamlitベースのWebアプリケーション。複雑なスプレッドシート構造とXMLベースの図形情報を正確に解析します。

## 主な機能

- Excelファイルからの高度なメタデータ抽出
- 複雑なスプレッドシート構造の解析
- XMLベースの図形情報の抽出
- リアルタイムデータ探索用のStreamlitインターフェース
- OpenAI APIを活用したインテリジェントなメタデータ分析
- 包括的なExcel図形およびシェイプメタデータの解析

## 技術スタック

- Python 3.11
- Streamlit (Webインターフェース)
- Pandas (データ操作)
- OpenPyXL (Excel処理)
- ElementTree (XML解析)
- OpenAI API (メタデータ分析)

## インストール方法

```bash
# リポジトリのクローン
git clone https://github.com/yourusername/excel-metadata-extractor.git
cd excel-metadata-extractor

# 依存パッケージのインストール
pip install -r requirements.txt

# 環境変数の設定
export OPENAI_API_KEY=your_api_key

# アプリケーションの起動
streamlit run main.py
```

## 使用方法

1. Webブラウザで`http://localhost:5000`を開く
2. Excelファイルをアップロード
3. 自動的に以下の情報が抽出・表示されます：
   - ファイルプロパティ
   - シート情報
   - テーブル構造
   - 図形・グラフ情報
   - AIによる内容分析

## 環境変数

- `OPENAI_API_KEY`: OpenAI APIキー（必須）

## ライセンス

このプロジェクトはMITライセンスの下で公開されています。

## 注意事項

- 大きなExcelファイルの処理には時間がかかる場合があります
- OpenAI APIの利用には有効なAPIキーが必要です
